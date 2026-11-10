# Taskbar Middle-Click Close

A Windows utility that allows you to close applications by middle-clicking their taskbar buttons.

## Features
- Middle-click any taskbar button to close the application
- Runs silently in the system tray
- Single executable file
- No dependencies

## Create Production release
```shell
dotnet publish -c Release -r win-x64 --self-contained true
```

## Installation
1. ### Package
```shell
winapp package `
    .\bin\Release\net10.0-windows\win-x64\publish\ `
    --manifest .\Package.appxmanifest `
    --cert .\TeddyG.pfx `
    --cert-password password `
    --output .\TaskbarMiddleClickClose.msix
```

2. ### Install as Windows app
```shell
.\TaskbarMiddleClickClose.msix
```

## Usage
- The app runs in the system tray
- Right-click the tray icon to exit
- Middle-click any taskbar button to close that app