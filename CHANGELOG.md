# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.1] - 2026-08-29

### Fixed
- **Date preservation in EML output**: The converter now preserves both the sent date/time and received date/time when converting PST to EML files.

### Changed
- **`PstToEmlConverter/Core/OutlookPstReader.cs`** - Enhanced date/time handling:
  - **Lines 20-22**: Added MAPI property tag constants for reliable date extraction:
    - `PR_CLIENT_SUBMIT_TIME` (0x00390040) - SentOn property
    - `PR_MESSAGE_DELIVERY_TIME` (0x0E060040) - ReceivedTime property
  - **Lines 39-57**: Added `TryGetPropertyDateTime()` helper method to extract DateTime values from MAPI properties via PropertyAccessor, handling COM boxing and type conversion
  - **Lines 201-224**: Enhanced `ResolveSentDate()` method with fallback chain:
    1. Try MAPI `PR_CLIENT_SUBMIT_TIME` (most reliable)
    2. Fall back to COM `SentOn` property
    3. Try MAPI `PR_MESSAGE_DELIVERY_TIME` (received time)
    4. Fall back to COM `ReceivedTime` property
    5. Default to `DateTime.Now` if all fail
  - **Lines 473-478**: Enhanced received time extraction in `WriteEmlFromMailItem()`:
    - Try MAPI `PR_MESSAGE_DELIVERY_TIME` first
    - Fall back to COM `ReceivedTime` property
    - Validate with `IsPlausibleDate()` check
  - **Lines 485-492**: Added `Received` header generation with original received timestamp
  - **Line 502**: Write `Received` header to EML output if available

- **`PstToEmlConverter/PstToEmlConverter.csproj`** - **Line 5**: Changed target framework from `net10.0-windows` to `net8.0-windows` for compatibility with installed .NET SDK
  - **Note**: This change is a temporary downgrade due to the .NET 10 SDK not being available in the current build environment. When .NET 10 is released and available, the target framework can be reverted to `net10.0-windows` along with any associated package updates. Future contributors should verify SDK availability and update the `<TargetFramework>` element and any NuGet package versions accordingly.

- **`installer/PstToEmlConverter.iss`** - Installer script updates:
  - **Line 2**: Updated comment from ".NET 10" to ".NET 8"
  - **Line 6**: Updated version from `1.0.0` to `1.0.1`
  - **Line 41**: Updated source path from `net10.0-windows` to `net8.0-windows` for published output

### Technical Details

The changes ensure that EML files generated from PST conversion now contain:
1. **`Date:` header** - The sent date/time (RFC 5322 format with timezone offset)
2. **`Received:` header** - The received date/time (preserved from original message)

Both timestamps include full precision (hours, minutes, seconds) and timezone information, resolving the issue where date/time information was lost during conversion.

### Build
- Project builds successfully with MSBuild (Visual Studio 2022)
- Self-contained publish for win-x64 works correctly
- Inno Setup installer compiles to `installer/Output/PstToEmlConverter-1.0.1-Setup.exe`