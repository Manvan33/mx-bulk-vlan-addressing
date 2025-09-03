# MX Bulk VLAN Addressing

A comprehensive Python script for managing Cisco Meraki MX VLAN configurations through Excel import/export functionality.

## Use Cases

- Export VLAN configurations from Meraki dashboard to an Excel file
- Validate Excel files for proper VLAN configuration format
- Apply VLAN configurations from Excel to the Meraki dashboard

## Limitations

- The script only supports addition. Deleting networks and VLANs should be done via the Meraki dashboard.
- Networks are indexed by their names. Renaming a network in Excel and then applying to dashboard will create a new network, leaving the old one intact.
- VLANs are indexed by their IDs. Modifying a VLAN ID in Excel will create a new VLAN instead of updating the existing one.

## Setup

### Prerequisites

- Python 3.11 or higher
- Cisco Meraki API access (API key or OAuth application)

### Installation

#### 1. Clone the repository:

```bash
git clone https://github.com/Manvan33/mx-bulk-vlan-addressinggit
cd mx-bulk-vlan-addressing
```
#### 2. Create a virtual environment and install packages with pip:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

#### 3. Set up authentication:

**Option A: API Key (simpler)**

```bash
cp .env.example .env
```
Edit .env and add your MERAKI_API_KEY

**Option B: OAuth (more secure)**
- Set up an OAuth application on https://integrate.cisco.com
- Edit `.env` to include your client ID and client secret
- Uncomment a line in main.py to use OAuth

## Usage

```
usage: main.py [-h] --org ORG {check-api,validate-excel,apply-from-excel,export-to-excel,create-networks,create-vlans} ...

Meraki MX VLAN Configuration Tool

positional arguments:
  {check-api,validate-excel,apply-from-excel,export-to-excel,create-networks,create-vlans}
                        Available commands
    check-api           Check Meraki API connectivity
    validate-excel      Validate Excel file format and data
    apply-from-excel    Apply VLAN configuration from Excel to Dashboard
    export-to-excel     Export VLAN configuration from Dashboard to Excel
    create-networks     Create networks
    create-vlans        Create VLANs

options:
  -h, --help            show this help message and exit
  --org ORG             Organization ID
```

### Example Usage

0. Verify that your API key has access to your organization.

```bash
python main.py --org <your_org_id> check-api
```

1. First generate a spreadsheet of your existing configuration using the `export-to-excel` command:

```bash
python main.py --org <your_org_id> export-to-excel
```

2. Modify the Excel file as needed to reflect your desired VLAN configuration.

- You can delete lines for networks/VLANs you don't want to modify.
- You can add VLANs and networks
- You can rename VLANs, modify subnets and MX IPs.

3. Verify the Excel file.

```bash
python main.py --org <your_org_id> validate-excel --file <path_to_your_excel_file>
```

4. Create networks and VLANs from the modified Excel file.

```bash
python main.py --org <your_org_id> create-networks --file <path_to_your_excel_file>
python main.py --org <your_org_id> create-vlans --file <path_to_your_excel_file>
```

5. Apply the addressing configuration from the modified Excel file.

```bash
python main.py --org <your_org_id> apply-from-excel --file <path_to_your_excel_file>
```

## Excel File Format

The tool exports and validates Excel files with this structure:

| Network Name | VLAN ID | VLAN Name | Subnet | MX IP |
|--------------|---------|-----------|--------|-------|
| Site-A | 10 | Data | 192.168.10.0/24 | 192.168.10.1 |
| Site-A | 20 | Guest | 192.168.20.0/24 | 192.168.20.1 |
| Site-B | 10 | Data | 192.168.30.0/24 | 192.168.30.1 |

You can create new columns as needed, they will be ignored. Don't rename the existing columns.

### Validation Rules

- **Network Name**: Letters, numbers, spaces, and characters: `. @ # _ -`
- **VLAN ID**: Integer between 1-4094
- **VLAN Name**: Letters, numbers, spaces, and characters: `. @ # _ -`
- **Subnet**: Valid CIDR notation (e.g., `192.168.1.0/24`)
- **MX IP**: Valid IP address that belongs to the specified subnet

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Disclaimer

This tool is not officially supported by Cisco Meraki. Use at your own risk and always test in a lab environment before applying to production networks.
