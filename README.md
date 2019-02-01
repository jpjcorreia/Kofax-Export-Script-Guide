# Kofax-Export-Script-Guide

## Development Environment

The guide relies on **Visual Studio**

## Project Settings

### Project Type

.NET class library

### Framework

.NET 4.0 Framework

```
=> Application => Target framwork
```

### COM (Component Object Model)

In order for the interfaces to communicate with Kofax, COM visibility must be enabled

```
=> Application => Assembly Information => Make assembly COM-Visible
```

### Target Platform

As a 32-bit application, the target platform is x86

```
=> Build => Platform target
```

### Build

The files required for Kofax must be stored in the Bin directory of Kofax. This can be found at

```
.....\Program Files (x86)\Kofax\CaptureSS\ServLib\Bin\
```

In order to optimize the development, the output path of the . dll file can be set directly via Visual Studio

```
=> Build => Output path
```

### Debugging

#### Debug Information

To get complete debug information, the option must be specified

```
=> Build => Advanced => Debugging information
```

#### Local Tests

An external program can be started to debug the respective script

```
// Administration
C:\Program Files (x86)\Kofax\CaptureSS\ServLib\Bin\Admin.exe
```

```
// Release
C:\Program Files (x86)\Kofax\CaptureSS\ServLib\Bin\Release.exe
```

The path can be defined under

```
=> Debug => Start external program
```

### Dependencies

The **Kofax.ReleaseLib.Interop.dll** must be included in the references

```
=> References => Add Reference => Browse => C:\Program Files (x86)\Kofax\CaptureSS\ServLib\Bin\Kofax.ReleaseLib.Interop.dll
```

### Registration

A [.inf file](https://github.com/matthiashermsen/Kofax-Export-Script-Guide/blob/master/src/LeeresExportScript.inf) is required to install the script. This is then stored in the Kofax Bin directory

**The file must not be created in the UTF-8 format, but must use the UTF-8 without BOM Fortmat**

### Setup Script

#### Interface

The [setup script](https://github.com/matthiashermsen/Kofax-Export-Script-Guide/blob/master/src/KfxReleaseScriptSetup.cs) is used by the administration module

#### Setup Form

The setup script starts a [form](https://github.com/matthiashermsen/Kofax-Export-Script-Guide/blob/master/src/FrmSetup.cs)

### Release Script

The [release script](https://github.com/matthiashermsen/Kofax-Export-Script-Guide/blob/master/src/KfxReleaseScript.cs) is executed during the export

### Register the project on the machine

To register the project locally, RegAsm must be run once as administrator in the console

```
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" LeeresExportScript.dll /codebase /tlb:LeeresExportScript.tlb
```

### Install the Script

The script can be installed via the administration module

```
Extras Tab => Export Scripte => Hinzufügen => .inf Datei im Kofax Bin Verzeichnis
```

## Rollout

To deliver the script to the customer, the project . dll and . inf file must be placed in the customer's Kofax Bin directory
