import os
import re
import sys
import ipaddress
import pandas as pd
import argparse
from meraki import DashboardAPI
from dotenv import load_dotenv
from src.meraki_api_auth import APIKeyAuth, OAuthAuth

# Load environment variables from .env if present
load_dotenv()

_dashboard_instance = None


def init_sdk() -> DashboardAPI:
    """
    Initialize the Meraki Dashboard API SDK.
    Uses a singleton pattern to ensure only one instance is created.
    """
    global _dashboard_instance
    if _dashboard_instance is not None:
        return _dashboard_instance

    # Ensure logs directory exists
    os.makedirs("output/logs", exist_ok=True)

    # Use a personal API Key for authentication
    auth = APIKeyAuth()
    # If you want to use OAuth instead, check https://developer.cisco.com/meraki/api-v1/oauth-overview/ and uncomment the line below
    # auth = OAuthAuth()
    dashboard = DashboardAPI(
        api_key=auth.get_auth_token(),
        log_path="output/logs",
        print_console=False
    )

    _dashboard_instance = dashboard
    return dashboard

def validate_org_id(dashboard, org_id):
    """
    Validate that the provided org_id exists in the user's organizations.
    Returns True if valid, False otherwise.
    """
    try:
        orgs = dashboard.organizations.getOrganizations()
        for org in orgs:
            if str(org['id']) == str(org_id):
                return True
        print(f"\n❌ Organization ID {org_id} is NOT available!")
        print("Available organizations:")
        for org in orgs:
            print(f"  - {org['name']} (ID: {org['id']})")
        return False
    except Exception as e:
        print(f"Error validating organization ID: {e}")
        return False

def import_from_dashboard(org_id):
    """
    Import VLAN addressing configuration from Meraki dashboard and generate Excel file.
    
    This function retrieves the MX VLAN configuration for each network in the organization
    and creates an Excel file with the current addressing scheme.
    
    Args:
        org_id (str): Organization ID to import from
        
    Returns:
        str: Path to the generated Excel file
    """
    # Use init_sdk to get the dashboard instance (ensures consistent auth and logging setup)
    dashboard = init_sdk()
    
    # Get all networks in the organization
    print(f"Fetching networks for organization {org_id}...")
    networks = dashboard.organizations.getOrganizationNetworks(org_id)
    
    # Filter for networks with MX appliances
    mx_networks = [net for net in networks if 'appliance' in net.get('productTypes', [])]
    print(f"Found {len(mx_networks)} networks with MX appliances")
    
    # Collect VLAN data
    vlan_data = []
    
    for network in mx_networks:
        network_id = network['id']
        network_name = network['name']
        
        try:
            print(f"Processing network: {network_name}")
            
            # Get VLAN configuration for this network
            vlans = dashboard.appliance.getNetworkApplianceVlans(network_id)
            
            for vlan in vlans:
                # Get MX IP (appliance IP) from the VLAN configuration
                mx_ip = vlan.get('applianceIp', '')
                subnet = vlan.get('subnet', '')
                
                # Only include VLANs that have proper configuration
                if subnet and mx_ip:
                    vlan_data.append({
                        'Network Name': network_name,
                        'VLAN ID': vlan.get('id', ''),
                        'VLAN Name': vlan.get('name', ''),
                        'Subnet': subnet,
                        'MX IP': mx_ip
                    })
                    
        except Exception as e:
            print(f"Warning: Could not retrieve VLAN data for network {network_name}: {str(e)}")
            continue
    
    if not vlan_data:
        print("No VLAN data found in any networks.")
        return None
    
    # Create DataFrame and save to Excel
    df = pd.DataFrame(vlan_data)
    
    # Ensure output directory exists
    os.makedirs("output/spreadsheets", exist_ok=True)
    
    # Generate filename with timestamp
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"output/spreadsheets/meraki_vlan_export_{org_id}_{timestamp}.xlsx"
    
    # Save to Excel file
    df.to_excel(filename, index=False)
    print(f"✅ VLAN configuration exported to: {filename}")
    print(f"   Total VLANs exported: {len(vlan_data)}")
    print(f"   Networks processed: {len(mx_networks)}")
    
    return filename

def load_from_excel(filepath):
    """
    Load VLAN configuration from an Excel file.

    Args:
        filepath (str): Path to the Excel file.

    Returns:
        pd.DataFrame: DataFrame containing the VLAN configuration, or None if loading failed.
    """
    try:
        df = pd.read_excel(filepath)
        return df
    except FileNotFoundError:
        print(f"Excel file not found: {filepath}")
        return None
    except Exception as e:
        print(f"Error loading Excel file: {str(e)}")
        return None

def validate_excel_format(excel_data):
    """
    Check if the Excel file has the correct format.
    
    Expected format:
    - First sheet must have 5 columns: Network Name, VLAN ID, VLAN Name, Subnet, MX IP
    - Network Name: letters, numbers, spaces, and characters: . @ # _ -
    - VLAN ID: integer between 1-4094
    - VLAN Name: letters, numbers, spaces, and characters: . @ # _ -
    - Subnet: valid CIDR notation
    - MX IP: valid IP address that belongs to the subnet
    
    Args:
        filepath (str): Path to the Excel file
        
    Returns:
        tuple: (bool, list) - (is_valid, list_of_errors)
    """
    errors = []

    if excel_data is None:
        errors.append("No data found.")
        return False, errors

    # Check if required columns exist
    required_columns = ['Network Name', 'VLAN ID', 'VLAN Name', 'Subnet', 'MX IP']
    missing_columns = [col for col in required_columns if col not in excel_data.columns]
    if missing_columns:
            errors.append(f"Missing required columns: {', '.join(missing_columns)}")
            return False, errors
    
    # Pattern for allowed characters in names
    name_pattern = re.compile(r'^[a-zA-Z0-9\s.@#_-]+$')
    
    # Check each row
    for row_idx in range(len(excel_data)):
        row = excel_data.iloc[row_idx]
        row_num = row_idx + 2  # Excel rows start at 1, plus header
        
        # Check Network Name
        network_name = str(row['Network Name']).strip()
        if not network_name or pd.isna(row['Network Name']):
            errors.append(f"Row {row_num}: Network Name is empty")
        elif not name_pattern.match(network_name):
            errors.append(f"Row {row_num}: Network Name '{network_name}' contains invalid characters")
        
        # Check VLAN ID
        try:
            vlan_id = int(row['VLAN ID'])
            if vlan_id < 1 or vlan_id > 4094:
                errors.append(f"Row {row_num}: VLAN ID {vlan_id} must be between 1-4094")
        except (ValueError, TypeError):
            errors.append(f"Row {row_num}: VLAN ID must be a valid integer")
        
        # Check VLAN Name
        vlan_name = str(row['VLAN Name']).strip()
        if not vlan_name or pd.isna(row['VLAN Name']):
            errors.append(f"Row {row_num}: VLAN Name is empty")
        elif not name_pattern.match(vlan_name):
            errors.append(f"Row {row_num}: VLAN Name '{vlan_name}' contains invalid characters")
        
        # Check Subnet (CIDR notation)
        subnet_str = str(row['Subnet']).strip()
        try:
            subnet = ipaddress.IPv4Network(subnet_str, strict=False)
        except (ValueError, ipaddress.AddressValueError):
            errors.append(f"Row {row_num}: Subnet '{subnet_str}' is not valid CIDR notation")
            continue  # Skip MX IP check if subnet is invalid
        
        # Check MX IP
        mx_ip_str = str(row['MX IP']).strip()
        try:
            mx_ip = ipaddress.IPv4Address(mx_ip_str)
            # Check if MX IP belongs to the subnet
            if mx_ip not in subnet:
                errors.append(f"Row {row_num}: MX IP '{mx_ip_str}' does not belong to subnet '{subnet_str}'")
        except (ValueError, ipaddress.AddressValueError):
            errors.append(f"Row {row_num}: MX IP '{mx_ip_str}' is not a valid IP address")
    
    # Return validation result
    is_valid = len(errors) == 0
    return is_valid, errors
        
def validate_excel_data(org_id, excel_data):
    """
    Check if networks from Excel data exist in the organization and validate VLANs.

    Args:
        org_id (str): Organization ID to check.
        excel_data (pd.DataFrame): DataFrame containing VLAN configuration.

    Returns:
        dict: A dictionary with validation results:
              - 'networks': dict mapping existing network names to their IDs
              - 'vlan_validation': list of validation results for each VLAN
              - 'summary': dict with counts of valid/invalid networks and VLANs
    """
    dashboard = init_sdk()
    
    try:
        # Get all networks in the organization
        networks = dashboard.organizations.getOrganizationNetworks(org_id)
        
        # Create a mapping of network names to IDs
        network_map = {net['name']: net['id'] for net in networks}
        
        # Get unique network names from Excel data
        excel_networks = excel_data['Network Name'].unique()
        
        # Check which networks exist
        existing_networks = {}
        missing_networks = []
        vlan_validation = []
        
        for network_name in excel_networks:
            if network_name in network_map:
                existing_networks[network_name] = network_map[network_name]
            else:
                missing_networks.append(network_name)
        
        # Report network findings
        if existing_networks:
            print(f"✅ Found {len(existing_networks)} existing networks")
            for name in existing_networks.keys():
                print(f"   - {name}")
        
        if missing_networks:
            print(f"⚠️  {len(missing_networks)} networks not found in organization:")
            for name in missing_networks:
                print(f"   - {name}")
        
        # Now validate VLANs for existing networks
        print(f"\nValidating VLANs in existing networks...")
        valid_vlans = 0
        invalid_vlans = 0
        
        for network_name, network_data in excel_data.groupby('Network Name'):
            if network_name not in existing_networks:
                # Skip VLAN validation for missing networks
                for _, row in network_data.iterrows():
                    vlan_validation.append({
                        'network_name': network_name,
                        'vlan_id': row['VLAN ID'],
                        'vlan_name': row['VLAN Name'],
                        'status': 'network_missing',
                        'message': f"Network '{network_name}' does not exist"
                    })
                    invalid_vlans += 1
                continue
            
            network_id = existing_networks[network_name]
            
            try:
                # Get existing VLANs for this network
                existing_vlans = dashboard.appliance.getNetworkApplianceVlans(network_id)
                existing_vlan_ids = {str(vlan['id']): vlan for vlan in existing_vlans}
                
                print(f"   {network_name}: Found {len(existing_vlans)} existing VLANs")
                
                # Check each VLAN from Excel
                for _, row in network_data.iterrows():
                    vlan_id = str(row['VLAN ID'])
                    vlan_name = str(row['VLAN Name']).strip()
                    
                    if vlan_id in existing_vlan_ids:
                        existing_vlan = existing_vlan_ids[vlan_id]
                        vlan_validation.append({
                            'network_name': network_name,
                            'vlan_id': vlan_id,
                            'vlan_name': vlan_name,
                            'status': 'existing',
                            'message': f"VLAN {vlan_id} exists (current name: '{existing_vlan.get('name', 'Unknown')}')",
                            'current_config': existing_vlan
                        })
                        valid_vlans += 1
                        print(f"      ✅ VLAN {vlan_id} ({vlan_name}) - exists")
                    else:
                        vlan_validation.append({
                            'network_name': network_name,
                            'vlan_id': vlan_id,
                            'vlan_name': vlan_name,
                            'status': 'missing',
                            'message': f"VLAN {vlan_id} does not exist in network"
                        })
                        invalid_vlans += 1
                        print(f"      ❌ VLAN {vlan_id} ({vlan_name}) - does not exist")
                        
            except Exception as e:
                print(f"   ⚠️  Could not retrieve VLANs for network {network_name}: {str(e)}")
                # Mark all VLANs for this network as validation failed
                for _, row in network_data.iterrows():
                    vlan_validation.append({
                        'network_name': network_name,
                        'vlan_id': row['VLAN ID'],
                        'vlan_name': row['VLAN Name'],
                        'status': 'validation_error',
                        'message': f"Could not validate VLAN: {str(e)}"
                    })
                    invalid_vlans += 1
        
        # Create network validation details
        network_validation = []
        for network_name in excel_networks:
            network_validation.append({
                'name': network_name,
                'id': existing_networks.get(network_name, ''),
                'exists': network_name in existing_networks
            })
        
        # Summary
        summary = {
            'networks_found': len(existing_networks),
            'networks_missing': len(missing_networks),
            'vlans_valid': valid_vlans,
            'vlans_invalid': invalid_vlans,
            'total_networks': len(excel_networks),
            'total_vlans': len(excel_data)
        }
        
        return {
            'networks': existing_networks,
            'network_validation': network_validation,
            'vlan_validation': vlan_validation,
            'summary': summary
        }
        
    except Exception as e:
        print(f"Error validating networks and VLANs: {str(e)}")
        return {
            'networks': {},
            'network_validation': [],
            'vlan_validation': [],
            'summary': {
                'networks_found': 0,
                'networks_missing': 0,
                'vlans_valid': 0,
                'vlans_invalid': 0,
                'total_networks': 0,
                'total_vlans': 0
            }
        }

def validate_excel(org_id, filepath):
    """
    Load and comprehensively validate Excel file data against organization.
    
    Args:
        org_id (str): Organization ID
        filepath (str): Path to Excel file
        
    Returns:
        dict: Complete validation results with excel_data, validation status, and actions needed
    """
    print(f"Validating Excel file: {filepath}")
    print(f"Organization ID: {org_id}")
    
    # Step 1: Load Excel data
    print("\nStep 1: Loading Excel data...")
    excel_data = load_from_excel(filepath)
    if excel_data is None:
        print("❌ Failed to load Excel data")
        return {
            'success': False,
            'excel_data': None,
            'validation_results': None,
            'actions_needed': ['fix_excel_file'],
            'errors': ['Failed to load Excel data']
        }
    print(f"✅ Loaded {len(excel_data)} rows from Excel file")

    # Step 2: Validate Excel format
    print("\nStep 2: Validating Excel file format...")
    is_valid, errors = validate_excel_format(excel_data)
    if not is_valid:
        print("❌ Excel file format validation failed:")
        for error in errors:
            print(f"  - {error}")
        return {
            'success': False,
            'excel_data': None,
            'validation_results': None,
            'actions_needed': ['fix_excel_format'],
            'errors': errors
        }
    print("✅ Excel file format is valid!")

    # Step 3: Validate against organization
    print(f"\nStep 3: Validating against organization {org_id}...")
    validation_results = validate_excel_data(org_id, excel_data)
    
    # Determine what actions are needed
    actions_needed = []
    summary = validation_results['summary']
    vlan_validation = validation_results['vlan_validation']
    
    if summary['networks_missing'] > 0:
        actions_needed.append('create_networks')
    
    missing_vlans = [v for v in vlan_validation if v['status'] == 'missing']
    if missing_vlans:
        actions_needed.append('create_vlans')
    
    validation_errors = [v for v in vlan_validation if v['status'] == 'validation_error']
    network_missing_vlans = [v for v in vlan_validation if v['status'] == 'network_missing']
    
    if validation_errors:
        actions_needed.append('fix_validation_errors')
    
    # Determine overall success
    can_proceed = (summary['networks_found'] > 0 and 
                   len(validation_errors) == 0 and 
                   len(network_missing_vlans) == 0)
    
    return {
        'success': can_proceed,
        'excel_data': excel_data,
        'validation_results': validation_results,
        'actions_needed': actions_needed,
        'summary': {
            'total_networks': summary['total_networks'],
            'networks_found': summary['networks_found'],
            'networks_missing': summary['networks_missing'],
            'total_vlans': summary['total_vlans'],
            'vlans_existing': summary['vlans_valid'],
            'vlans_missing': len(missing_vlans),
            'vlans_errors': len(validation_errors)
        }
    }

def create_networks(org_id, excel_data, existing_networks):
    """
    Create networks that exist in Excel data but don't exist in the organization.
    
    Args:
        org_id (str): Organization ID
        excel_data (pd.DataFrame): DataFrame containing VLAN configuration
        existing_networks (dict): Dictionary of existing network names to IDs
        
    Returns:
        dict: Updated dictionary of network names to IDs (including newly created ones)
    """
    dashboard = init_sdk()
    
    # Get unique network names from Excel data
    excel_networks = excel_data['Network Name'].unique()
    missing_networks = [name for name in excel_networks if name not in existing_networks]
    
    if not missing_networks:
        print("✅ All networks already exist, no networks to create")
        return existing_networks
    
    print(f"\nCreating {len(missing_networks)} missing networks...")
    updated_networks = existing_networks.copy()
    
    for network_name in missing_networks:
        try:
            print(f"   Creating network: {network_name}")
            
            # Create network with MX appliance product type
            # Note: You may need to adjust these parameters based on your requirements
            network = dashboard.organizations.createOrganizationNetwork(
                organizationId=org_id,
                name=network_name,
                productTypes=['appliance'],  # MX appliance
                tags=[],
                timeZone='America/Los_Angeles'  # You may want to make this configurable
            )
            # Enable VLANs for this network
            dashboard.appliance.updateNetworkApplianceVlansSettings(
                networkId=network['id'],
                vlansEnabled=True
            )
            network_id = network['id']
            updated_networks[network_name] = network_id
            print(f"   ✅ Created network '{network_name}' (ID: {network_id})")
            
        except Exception as e:
            print(f"   ❌ Failed to create network '{network_name}': {str(e)}")
            continue
    
    print(f"Network creation summary:")
    print(f"   ✅ Successfully created: {len(updated_networks) - len(existing_networks)} networks")
    print(f"   ❌ Failed to create: {len(missing_networks) - (len(updated_networks) - len(existing_networks))} networks")
    
    return updated_networks

def create_vlans(excel_data, network_mapping, vlan_validation):
    """
    Create VLANs that exist in Excel data but don't exist in their respective networks.
    
    Args:
        excel_data (pd.DataFrame): DataFrame containing VLAN configuration
        network_mapping (dict): Dictionary of network names to IDs
        vlan_validation (list): List of VLAN validation results
        
    Returns:
        dict: Summary of VLAN creation results
    """
    dashboard = init_sdk()
    
    # Find VLANs that need to be created
    vlans_to_create = [v for v in vlan_validation if v['status'] == 'missing']
    
    if not vlans_to_create:
        print("✅ All VLANs already exist, no VLANs to create")
        return {'created': 0, 'failed': 0, 'total': 0}
    
    print(f"\nCreating {len(vlans_to_create)} missing VLANs...")
    
    created_count = 0
    failed_count = 0
    
    # Group VLANs by network for efficient processing
    vlans_by_network = {}
    for vlan_info in vlans_to_create:
        network_name = vlan_info['network_name']
        if network_name not in vlans_by_network:
            vlans_by_network[network_name] = []
        vlans_by_network[network_name].append(vlan_info)
    
    for network_name, vlans in vlans_by_network.items():
        if network_name not in network_mapping:
            print(f"⚠️  Skipping VLANs for network '{network_name}' - network does not exist")
            failed_count += len(vlans)
            continue
        
        network_id = network_mapping[network_name]
        print(f"\nCreating VLANs for network: {network_name}")
        
        for vlan_info in vlans:
            try:
                # Find the corresponding row in Excel data to get all VLAN details
                vlan_rows = excel_data[
                    (excel_data['Network Name'] == network_name) & 
                    (excel_data['VLAN ID'].astype(str) == str(vlan_info['vlan_id']))
                ]
                if vlan_rows.empty:
                    failed_count += 1
                    print(f"   ❌ Failed to create VLAN {vlan_info['vlan_id']}: No matching row found in Excel data for network '{network_name}' and VLAN ID '{vlan_info['vlan_id']}'")
                    continue
                vlan_row = vlan_rows.iloc[0]
                vlan_id = str(vlan_row['VLAN ID'])
                vlan_name = str(vlan_row['VLAN Name']).strip()
                subnet = str(vlan_row['Subnet']).strip()
                mx_ip = str(vlan_row['MX IP']).strip()
                print(f"   Creating VLAN {vlan_id}: {vlan_name} ({subnet})")
                # Create VLAN
                vlan = dashboard.appliance.createNetworkApplianceVlan(
                    networkId=network_id,
                    id=vlan_id,
                    name=vlan_name,
                    subnet=subnet,
                    applianceIp=mx_ip
                )
                created_count += 1
                print(f"   ✅ Created VLAN {vlan_id} successfully")
            except Exception as e:
                failed_count += 1
                print(f"   ❌ Failed to create VLAN {vlan_info['vlan_id']}: {str(e)}")
                continue
    
    print(f"\n📊 VLAN creation summary:")
    print(f"   ✅ Successfully created: {created_count} VLANs")
    print(f"   ❌ Failed to create: {failed_count} VLANs")
    print(f"   Total processed: {len(vlans_to_create)} VLANs")
    
    return {
        'created': created_count, 
        'failed': failed_count, 
        'total': len(vlans_to_create)
    }

def apply_excel_data(validation_results):
    """
    Apply VLAN addressing configuration from the provided Excel data.

    Args:
        org_id (str): Organization ID to apply addressing for.
        validation_results (dict): Results from the Excel validation process.

    Returns:
        bool: True if addressing was applied successfully, False otherwise.
    """
    dashboard = init_sdk()
    
    # Validate the existing networks and VLANs
    existing_networks = validation_results['validation_results']['networks']

    if not existing_networks:
        print("❌ No valid existing networks found.")
        return False

    print(f"\nStarting VLAN configuration updates...")
    success_count = 0
    error_count = 0

    # Group data by network for efficient processing
    for network_name, network_data in validation_results['excel_data'].groupby('Network Name'):
        if network_name not in existing_networks:
            print(f"⚠️  Skipping network '{network_name}' - does not exist")
            continue

        network_id = existing_networks[network_name]
        print(f"\nUpdating network: {network_name} (ID: {network_id})")
        
        # Process each VLAN for this network
        for _, row in network_data.iterrows():
            try:
                vlan_id = int(row['VLAN ID'])
                vlan_name = str(row['VLAN Name']).strip()
                subnet = str(row['Subnet']).strip()
                mx_ip = str(row['MX IP']).strip()
                
                print(f"   Updating VLAN {vlan_id}: {vlan_name} -> {subnet} (MX: {mx_ip})")
                
                # Update VLAN configuration
                dashboard.appliance.updateNetworkApplianceVlan(
                    networkId=network_id,
                    vlanId=str(vlan_id),
                    name=vlan_name,
                    subnet=subnet,
                    applianceIp=mx_ip
                )
                
                success_count += 1
                print(f"   ✅ VLAN {vlan_id} updated successfully")
                
            except Exception as e:
                error_count += 1
                print(f"   ❌ Failed to update VLAN {row.get('VLAN ID', 'Unknown')}: {str(e)}")
                continue

    print(f"\nSummary:")
    print(f"   ✅ Successfully updated: {success_count} VLANs")
    print(f"   ❌ Failed updates: {error_count} VLANs")
    
    return error_count == 0

def main():
    """Main function to handle command line arguments."""
    parser = argparse.ArgumentParser(description="Meraki MX VLAN Configuration Tool")
    parser.add_argument('--org', required=True, help='Organization ID')
    subparsers = parser.add_subparsers(dest='command', help='Available commands')

    # Check API connection subcommand
    api_parser = subparsers.add_parser('check-api', help='Check Meraki API connectivity')

    # Validate Excel format subcommand
    validate_parser = subparsers.add_parser('validate-excel', help='Validate Excel file format and data')
    validate_parser.add_argument('--excel-file', required=True, help='Path to Excel file')

    # Import from Excel subcommand
    apply_parser = subparsers.add_parser('apply-from-excel', help='Apply VLAN configuration from Excel to Dashboard')
    apply_parser.add_argument('--excel-file', required=True, help='Path to Excel file')

    # Export to Excel subcommand
    export_parser = subparsers.add_parser('export-to-excel', help='Export VLAN configuration from Dashboard to Excel')
    export_parser.add_argument('--excel-file', required=False, help='Path to save Excel file', default=None)

    # Create networks subcommand
    create_net_parser = subparsers.add_parser('create-networks', help='Create networks')
    create_net_parser.add_argument('--excel-file', required=True, help='Path to Excel file')

    # Create VLANs subcommand
    create_vlan_parser = subparsers.add_parser('create-vlans', help='Create VLANs')
    create_vlan_parser.add_argument('--excel-file', required=True, help='Path to Excel file')

    args = parser.parse_args()

    org_id = args.org


    # Validate org_id for all commands except check-api
    dashboard = init_sdk()
    if not validate_org_id(dashboard, org_id):
        sys.exit(1)

    if args.command == 'validate-excel':
        excel_file = args.excel_file
        validation_result = validate_excel(org_id, excel_file)
        if not validation_result['success'] and 'fix_excel_format' in validation_result['actions_needed']:
            print("❌ Excel file format validation failed:")
            for error in validation_result['errors']:
                print(f"  - {error}")
            sys.exit(1)
        if not validation_result['success'] and 'fix_excel_file' in validation_result['actions_needed']:
            print("❌ Failed to load Excel data")
            sys.exit(1)
        print("\n" + "="*80)
        print("VALIDATION RESULTS")
        print("="*80)
        validation_results = validation_result['validation_results']
        summary = validation_results['summary']
        print(f"Networks - Total: {summary['total_networks']}, Found: {summary['networks_found']}, Missing: {summary['networks_missing']}")
        print(f"VLANs - Total: {summary['total_vlans']}, Valid: {summary['vlans_valid']}, Invalid: {summary['vlans_invalid']}")
        print(f"\nNetwork Validation:")
        for network in validation_results['network_validation']:
            status = "✅" if network['exists'] else "❌"
            print(f"  {status} {network['name']} ({network['id']})")
        print(f"\nVLAN Validation:")
        vlan_counts = {'existing': 0, 'missing': 0, 'validation_error': 0, 'network_missing': 0}
        for vlan in validation_results['vlan_validation']:
            status_icons = {
                'existing': '✅',
                'missing': '❌',
                'validation_error': '⚠️',
                'network_missing': '🚫'
            }
            icon = status_icons.get(vlan['status'], '❓')
            vlan_counts[vlan['status']] += 1
            if vlan['status'] != 'existing':
                print(f"  {icon} Network: {vlan['network_name']}, VLAN: {vlan['vlan_name']} (ID: {vlan['vlan_id']}) - {vlan.get('message', vlan['status'])}")
        print(f"\nVLAN Status Summary:")
        print(f"  ✅ Existing: {vlan_counts['existing']}")
        print(f"  ❌ Missing: {vlan_counts['missing']}")
        print(f"  ⚠️ Validation Errors: {vlan_counts['validation_error']}")
        print(f"  🚫 Network Missing: {vlan_counts['network_missing']}")
        print(f"\nRecommendations:")
        if 'create_networks' in validation_result['actions_needed']:
            print(f"  - Create {summary['networks_missing']} missing networks using 'create-networks'")
        if 'create_vlans' in validation_result['actions_needed']:
            print(f"  - Create {vlan_counts['missing']} missing VLANs using 'create-vlans'")
        if 'fix_validation_errors' in validation_result['actions_needed']:
            print(f"  - Fix {vlan_counts['validation_error']} validation errors in Excel file")
        if validation_result['success']:
            print("  ✅ Ready to import configuration using 'apply-from-excel'")

    elif args.command == 'apply-from-excel':
        excel_file = args.excel_file
        print(f"Applying VLAN configuration from: {excel_file}")
        print(f"Organization ID: {org_id}")
        validation_result = validate_excel(org_id, excel_file)
        if not validation_result['success']:
            if 'fix_excel_format' in validation_result['actions_needed']:
                print("❌ Excel file format validation failed:")
                for error in validation_result['errors']:
                    print(f"  - {error}")
                sys.exit(1)
            if 'fix_excel_file' in validation_result['actions_needed']:
                print("❌ Failed to load Excel data")
                sys.exit(1)
            if 'create_networks' in validation_result['actions_needed']:
                print("❌ Missing networks found. Please create networks first using 'create-networks'")
                sys.exit(1)
            if 'fix_validation_errors' in validation_result['actions_needed']:
                print("❌ Validation errors found. Please fix these first using 'validate-excel' to see details")
                sys.exit(1)
        print("✅ Validation passed! Proceeding with import...")
        print("\nApplying VLAN configuration...")
        apply_excel_data(validation_result)
        print("\n✅ VLAN configuration import completed successfully!")

    elif args.command == 'export-to-excel':
        excel_file = args.excel_file
        print(f"Exporting VLAN configuration from organization: {org_id}")
        generated_filename = import_from_dashboard(org_id)
        if generated_filename:
            if excel_file and excel_file != generated_filename:
                import shutil
                try:
                    shutil.copy2(generated_filename, excel_file)
                    print(f"✅ File copied to requested location: {excel_file}")
                    os.remove(generated_filename)
                except Exception as e:
                    print(f"⚠️ Could not copy to {excel_file}: {e}")
                    print(f"✅ VLAN configuration exported to: {generated_filename}")
            else:
                print(f"✅ VLAN configuration exported to: {generated_filename}")
        else:
            print("❌ Failed to export VLAN configuration")
            sys.exit(1)

    elif args.command == 'create-networks':
        excel_file = args.excel_file
        print(f"Creating missing networks from: {excel_file}")
        print(f"Organization ID: {org_id}")
        validation_result = validate_excel(org_id, excel_file)
        if not validation_result['success'] and 'fix_excel_format' in validation_result['actions_needed']:
            print("❌ Excel file format validation failed:")
            for error in validation_result['errors']:
                print(f"  - {error}")
            sys.exit(1)
        if not validation_result['success'] and 'fix_excel_file' in validation_result['actions_needed']:
            print("❌ Failed to load Excel data")
            sys.exit(1)
        excel_data = validation_result['excel_data']
        existing_networks = validation_result['validation_results']['networks']
        create_networks(org_id, excel_data, existing_networks)
        print("\n✅ Network creation completed successfully!")

    elif args.command == 'create-vlans':
        excel_file = args.excel_file
        print(f"Creating missing VLANs from: {excel_file}")
        print(f"Organization ID: {org_id}")
        validation_result = validate_excel(org_id, excel_file)
        if not validation_result['success'] and 'fix_excel_format' in validation_result['actions_needed']:
            print("❌ Excel file format validation failed:")
            for error in validation_result['errors']:
                print(f"  - {error}")
            sys.exit(1)
        if not validation_result['success'] and 'fix_excel_file' in validation_result['actions_needed']:
            print("❌ Failed to load Excel data")
            sys.exit(1)
        excel_data = validation_result['excel_data']
        network_mapping = validation_result['validation_results']['networks']
        vlan_validation = validation_result['validation_results']['vlan_validation']
        create_vlans(excel_data, network_mapping, vlan_validation)
        print("\n✅ VLAN creation completed successfully!")

    elif args.command == 'check-api':
        orgs = dashboard.organizations.getOrganizations()
        print("Available organizations:")
        found = False
        for org in orgs:
            print(f"  - {org['name']} (ID: {org['id']})")
            if str(org['id']) == str(org_id):
                found = True
        if found:
            print(f"\n✅ Organization ID {org_id} is available.")
        else:
            print(f"\n❌ Organization ID {org_id} is NOT available!")

    else:
        parser.print_help()

if __name__ == "__main__":
    main()
    