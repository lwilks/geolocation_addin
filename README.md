# Geolocation Addin for Revit 2024

A Revit 2024 add-in that automates geolocation workflows: takes linked models from a site model, copies them locally, applies shared coordinates via transform, and exports to IFC/NWC/DWG.

## Features

- **Automated coordinate publishing** — applies shared coordinates from the site model to each linked model using transform-based positioning
- **Multi-format export** — exports geolocated models to IFC, NWC, and DWG
- **Link mapping editor** — interactive dialog for mapping link instance names to target file names with import/export support (CSV and XML)
- **Fuzzy matching** — automatically matches link instances to imported mapping entries even when names differ slightly
- **Cloud model support** — opens models from Autodesk Construction Cloud via Desktop Connector
- **Configurable** — all settings editable through the UI or via JSON config file

## Prerequisites

- **Autodesk Revit 2024**
- **Windows 10/11**
- **.NET Framework 4.8** (included with Windows 10 1903+)
- **Navisworks NWC Export Utility** (optional, only needed for NWC export)

## Installation

### Installer

Download the latest `GeolocationAddin-Setup.exe` from [Releases](https://github.com/lwilks/geolocation_addin/releases) and run it. The installer will:

1. Copy the add-in DLL, manifest, and dependencies to `%APPDATA%\Autodesk\Revit\Addins\2024\`
2. Create the config directory at `C:\ProgramData\GeolocationAddin\`
3. Place sample configuration files if they don't already exist

### Manual installation

1. Build the solution (see [Building from source](#building-from-source))
2. Copy from `src\GeolocationAddin\bin\Release\` to `%APPDATA%\Autodesk\Revit\Addins\2024\`:
   - `GeolocationAddin.dll`
   - `GeolocationAddin.addin`
   - `Newtonsoft.Json.dll`
3. Create `C:\ProgramData\GeolocationAddin\` and copy the sample config:
   ```
   copy config\config.sample.json C:\ProgramData\GeolocationAddin\config.json
   ```

### Uninstalling

Run the uninstaller from Add/Remove Programs, or manually delete:
- `%APPDATA%\Autodesk\Revit\Addins\2024\GeolocationAddin.*`
- `%APPDATA%\Autodesk\Revit\Addins\2024\Newtonsoft.Json.dll`
- `C:\ProgramData\GeolocationAddin\` (if you no longer need the config)

## Configuration

Settings are stored at `C:\ProgramData\GeolocationAddin\config.json`. Edit directly or use the **Settings** button in Revit.

| Field | Description |
|-------|-------------|
| `csvMappingPath` | Path to a CSV mapping file (optional — can import via dialog instead) |
| `linkSourceFolder` | Folder to search for link source files (fallback for path resolution) |
| `outputFolder` | Where geolocated .rvt copies are saved |
| `ifcOutputFolder` | IFC export destination |
| `nwcOutputFolder` | NWC export destination |
| `dwgOutputFolder` | DWG export destination |
| `exportSettings` | Toggle IFC/NWC/DWG export individually |
| `fuzzyMatchSettings` | Enable/disable fuzzy matching and set thresholds |

### CSV mapping format

```csv
LinkInstanceName,TargetFileName
Building A - Arch Model,SiteX_BuildingA_Arch.rvt
Building B - Struct Model,SiteX_BuildingB_Struct.rvt
```

### XML mapping format

```xml
<LinkMappings>
  <Mapping LinkName="Building A - Arch Model" TargetFileName="SiteX_BuildingA_Arch.rvt"/>
  <Mapping LinkName="Building B - Struct Model" TargetFileName="SiteX_BuildingB_Struct.rvt"/>
</LinkMappings>
```

## Usage

1. Open the **site model** in Revit 2024 (the model containing all the linked models)
2. Go to the **Geolocation** ribbon tab and click **Run Geolocation**
3. The **Link Mapping Editor** dialog appears with all linked models listed
4. Assign target file names:
   - **Import** a CSV or XML mapping file, or
   - **Type** target names directly in the grid
   - Fuzzy matching auto-fills suggestions when importing
5. Review the mappings — rows are color-coded:
   - Green = exact match from import
   - Yellow = fuzzy match (review before selecting)
   - Red = duplicate target name (must resolve before processing)
6. Check the boxes for links you want to process
7. Click **Process Selected**

The add-in will then, for each selected link:
1. Open the linked model (detached from central)
2. Save a local copy to the output folder
3. Apply shared coordinates from the site model
4. Export to enabled formats (IFC/NWC/DWG)
5. Display a summary of results

## Building from source

### Requirements

- [.NET SDK](https://dotnet.microsoft.com/download) (any version that supports .NET Framework 4.8 targeting)
- Revit 2024 installed (for API assemblies at `C:\Program Files\Autodesk\Revit 2024\`)

### Build

```powershell
dotnet restore GeolocationAddin.sln
dotnet build GeolocationAddin.sln -c Release --no-restore
```

Output is in `src\GeolocationAddin\bin\Release\`.

### Creating the installer

Requires [Inno Setup 6](https://jrsoftware.org/isinfo.php) installed.

```powershell
.\installer\Build-Installer.ps1
```

This builds the solution in Release mode and compiles the Inno Setup script. The resulting installer is placed in `installer\Output\`.

## Project structure

```
geolocation_addin/
  GeolocationAddin.sln
  config/
    config.sample.json          # Sample runtime config
    mapping.sample.csv          # Sample CSV mapping
  deploy/
    Deploy-GeolocationAddin.ps1 # Developer deployment script
  installer/
    GeolocationAddin.iss        # Inno Setup installer script
    Build-Installer.ps1         # Build + package script
  src/GeolocationAddin/
    Application/
      GeolocationApp.cs         # Revit add-in entry point, ribbon setup
    Commands/
      GeolocationCommand.cs     # Main workflow command
      SettingsCommand.cs        # Standalone settings command
    Config/
      AddinConfig.cs            # Config model
      ConfigLoader.cs           # JSON config load/save
      CsvMapping.cs             # CSV parser with consume-once semantics
    Core/
      GeolocationWorkflow.cs    # Main orchestrator
      CoordinatePublisher.cs    # Shared coordinate strategies
      FileCopyManager.cs        # Link file path resolution
      ModelExporter.cs          # IFC/NWC/DWG export
    Helpers/
      FuzzyMatcher.cs           # Token overlap + Levenshtein matching
      LogHelper.cs              # File logging
      MappingSerializer.cs      # CSV + XML import/export
      PathHelper.cs             # Path utilities
      RevitDocumentHelper.cs    # Revit document open/save/close helpers
    Models/
      LinkInstanceInfo.cs       # Full link data for processing
      LinkMatchInfo.cs          # Link + match state for UI binding
      FuzzyMatchResult.cs       # Fuzzy match result
      ProcessingResult.cs       # Per-link processing outcome
    UI/
      LinkMappingWindow.xaml    # Mapping editor + settings dialog
      SettingsWindow.xaml       # Standalone settings dialog
    GeolocationAddin.addin      # Revit manifest
    GeolocationAddin.csproj     # Project file (.NET Framework 4.8)
```

## Log file

All operations are logged to `C:\ProgramData\GeolocationAddin\geolocation.log`.

## License

All rights reserved.
