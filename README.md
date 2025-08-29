# MX Templates Bulk Addressing

A comprehensive Python tool for managing Cisco Meraki MX VLAN configurations through Excel import/export functionality. Supports both API key and OAuth authentication methods.

## Features

- **Export VLAN configurations** from Meraki dashboard to Excel format
- **Validate Excel files** for proper VLAN configuration format
- **Multiple authentication methods**: API key and OAuth 2.0
- **HTTPS OAuth callback server** with beautiful web interface
- **Comprehensive validation**: Network names, VLAN IDs, IP addresses, and CIDR notation
- **Command-line interface** for easy automation

## Use Cases

1. **Audit existing networks**: Export current VLAN configurations to Excel for review
2. **Configuration validation**: Verify Excel files before applying changes
3. **Bulk configuration preparation**: Use exported data as template for new deployments
4. **Documentation**: Generate Excel reports of network VLAN configurations

## Setup

### Prerequisites

- Python 3.11 or higher
- Cisco Meraki API access (API key or OAuth application)

### Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/mx-templates-bulk-addressing.git
cd mx-templates-bulk-addressing
```

2. Install with uv (recommended):
```bash
uv sync
```

Or with pip:
```bash
pip install -r requirements.txt
```

3. Set up authentication:

**Option A: API Key (simpler)**
```bash
cp .env.example .env
# Edit .env and add your MERAKI_API_KEY
```

**Option B: OAuth (more secure)**
- Set up OAuth application in Meraki dashboard
- Use the built-in OAuth flow with `--oauth` flag

## Usage

### Export VLAN Configuration from Dashboard

```bash
# Using API key authentication
uv run main.py --import-dashboard YOUR_ORG_ID

# Using OAuth authentication
uv run main.py --import-dashboard YOUR_ORG_ID --oauth
```

### Validate Excel File Format

```bash
uv run main.py --check-excel your_file.xlsx
```

### List Organizations

```bash
uv run main.py --init-sdk
```

## Excel File Format

The tool exports and validates Excel files with this structure:

| Network Name | VLAN ID | VLAN Name | Subnet | MX IP |
|--------------|---------|-----------|--------|-------|
| Site-A | 10 | Data | 192.168.10.0/24 | 192.168.10.1 |
| Site-A | 20 | Guest | 192.168.20.0/24 | 192.168.20.1 |
| Site-B | 10 | Data | 192.168.30.0/24 | 192.168.30.1 |

### Validation Rules

- **Network Name**: Letters, numbers, spaces, and characters: `. @ # _ -`
- **VLAN ID**: Integer between 1-4094
- **VLAN Name**: Letters, numbers, spaces, and characters: `. @ # _ -`
- **Subnet**: Valid CIDR notation (e.g., `192.168.1.0/24`)
- **MX IP**: Valid IP address that belongs to the specified subnet

## Authentication Methods

### API Key Authentication

1. Create `.env` file:
```bash
MERAKI_API_KEY=your_api_key_here
```

2. Generate API key in Meraki Dashboard:
   - Organization > Settings > Dashboard API access
   - Generate new API key

### OAuth Authentication

1. Set up OAuth application in Meraki Dashboard
2. Use the `--oauth` flag to trigger OAuth flow
3. Built-in HTTPS server handles the callback automatically

## Project Structure

```
├── main.py                    # Main CLI application
├── src/
│   ├── meraki_api_auth.py    # Authentication classes
├── oauth_callback_flask.py   # OAuth HTTPS callback server
├── requirements.txt          # Dependencies
├── pyproject.toml           # Project configuration
└── README.md                # This file
```

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Disclaimer

This tool is not officially supported by Cisco Meraki. Use at your own risk and always test in a lab environment before applying to production networks.
