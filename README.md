# Selective Export

SelectiveExport is a plug-in for KeePass 2.x plugin which lets you export entries  according to your selection. The current KeePass export all the entries  of the current database.

## Features

- Support for selective file export based on file extensions
- Ability to exclude specific files or directories
- Maintain original directory structure during export
- Support for custom export rules
- Command-line interface for easy operation
- Export data in multiple formats (XML and CSV)

## Build

1. Prerequisites:
   - Visual Studio 2019 or later
   - .NET Framework 4.7.2 SDK

2. Build Steps:
   - Open `SelectiveExport.sln` in Visual Studio
   - Select Release configuration
   - Build Solution (Press F6 or use Build > Build Solution)
   - Find the compiled executable in `bin/Release` directory

3. Command Line Build:
   - dotnet build -c Release

## Installation

1. Download the released DLL file
2. Copy the DLL file to KeePass Plugins directory
3. Restart KeePass

## Usage

1. Select entries or groups in KeePass
2. Click "Export Selected Items..." in the Tools menu
3. Choose save location and confirm
   
   

## Development Environment

- Visual Studio 2022 or later
- .NET Framework 4.7.2

