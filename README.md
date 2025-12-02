# Taskbar Middle-Click Close

A Windows utility that allows you to close applications by middle-clicking their taskbar buttons.

## Features
- Middle-click any taskbar button to close the application
- Runs silently in the system tray
- Single executable file
- No dependencies

## Create Production release
```shell
dotnet build -c Release
```

## Installation
1. Download the latest release
2. Run `MiddleClickCloser.exe`
3. (Optional) Add to startup folder: `Win+R` → `shell:startup`

## Usage
- The app runs in the system tray
- Right-click the tray icon to exit
- Middle-click any taskbar button to close that app