# VBScript Project with BlueZone Integration

This project provides a framework for automating mainframe interactions using VBScript and BlueZone, along with utilities for making API calls and handling data.

## Project Structure

- `main.vbs` - Main script demonstrating BlueZone integration
- `lib/` - Directory containing utility scripts
  - `api_utils.vbs` - Utilities for making HTTP API calls
  - `data_utils.vbs` - Utilities for data handling and transformation
  - `file_utils.vbs` - Utilities for file operations
- `config/` - Configuration files
  - `settings.vbs` - Global settings and configuration
- `scripts/` - Directory for specific task scripts
  - `sample_query.vbs` - Example script for querying mainframe data
  - `sample_api_call.vbs` - Example script for making external API calls

## Prerequisites

- Windows OS
- BlueZone terminal emulator installed
- Access to target mainframe systems (if applicable)

## Usage

1. Edit `config/settings.vbs` to configure your environment
2. Run the main script: `cscript main.vbs`
3. For specific tasks, run individual scripts from the `scripts/` directory

## Customization

Modify the scripts in the `scripts/` directory for your specific needs, or create new scripts based on the provided examples.
