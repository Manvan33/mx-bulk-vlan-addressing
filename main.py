import os
import re
import ipaddress
import pandas as pd
import argparse
from meraki import DashboardAPI
from dotenv import load_dotenv
from src.meraki_api_auth import APIKeyAuth, OAuthAuth

# Load environment variables from .env if present
load_dotenv()


def init_sdk() -> DashboardAPI:
    # Ensure logs directory exists
    os.makedirs("output/logs", exist_ok=True)
    
    # Use a personal API Key for authentication
    auth = APIKeyAuth()
    # If you want to use OAuth instead, check https://developer.cisco.com/meraki/api-v1/oauth-overview/ and uncomment the line below
    # auth = OAuthAuth()
    dashboard = DashboardAPI(
        api_key=auth.get_auth_token(),
        log_path="output/logs"
    )

    orgs = dashboard.organizations.getOrganizations()
    for org in orgs:
        print(f"Organization: {org['name']} (ID: {org['id']})")

    return dashboard

def check_excel_format(filepath):
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
    
    try:
        # Read the Excel file
        df = pd.read_excel(filepath)
        
        # Check if required columns exist
        required_columns = ['Network Name', 'VLAN ID', 'VLAN Name', 'Subnet', 'MX IP']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            errors.append(f"Missing required columns: {', '.join(missing_columns)}")
            return False, errors
        
        # Pattern for allowed characters in names
        name_pattern = re.compile(r'^[a-zA-Z0-9\s.@#_-]+$')
        
        # Check each row
        for row_idx in range(len(df)):
            row = df.iloc[row_idx]
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
        
    except FileNotFoundError:
        errors.append(f"File not found: {filepath}")
        return False, errors
    except Exception as e:
        errors.append(f"Error reading Excel file: {str(e)}")
        return False, errors


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
    os.makedirs("output/exports", exist_ok=True)
    
    # Generate filename with timestamp
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"output/exports/meraki_vlan_export_{org_id}_{timestamp}.xlsx"
    
    # Save to Excel file
    df.to_excel(filename, index=False)
    print(f"✅ VLAN configuration exported to: {filename}")
    print(f"   Total VLANs exported: {len(vlan_data)}")
    print(f"   Networks processed: {len(mx_networks)}")
    
    return filename


def main():
    """Main function to handle command line arguments."""
    parser = argparse.ArgumentParser(description='Meraki MX VLAN Configuration Tool')
    parser.add_argument('--check-excel', type=str, metavar='FILEPATH', 
                       help='Check Excel file format for VLAN configuration')
    parser.add_argument('--init-sdk', action='store_true',
                       help='Initialize Meraki SDK and list organizations')
    parser.add_argument('--import-dashboard', type=str, metavar='ORG_ID',
                       help='Import VLAN configuration from Meraki dashboard for specified organization')
    
    args = parser.parse_args()
    
    if args.check_excel:
        print(f"Checking Excel file: {args.check_excel}")
        is_valid, errors = check_excel_format(args.check_excel)
        
        if is_valid:
            print("✅ Excel file format is valid!")
        else:
            print("❌ Excel file format has errors:")
            for error in errors:
                print(f"  - {error}")
        
        return is_valid
    
    elif args.import_dashboard:
        filename = import_from_dashboard(args.import_dashboard)
        return filename is not None
    
    elif args.init_sdk:
        init_sdk()
    
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
    