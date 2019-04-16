# Kofax-Export-Script-Guide

## <a name=Content></a> Table of Contents
1. [Development Environment](#DevEnv)
2. [Project Settings](#Settings)  
  2.1. [Project Type](#ProjectType)  
  2.2. [Framework](#Framework)  
  2.3. [COM (Component Object Model)](#COM)  
  2.4. [Target Platform](#Target)  
  2.5. [Build](#Build)  
  2.6. [Debugging](#Debugging)  
    &emsp;&ensp;&nbsp;2.6.1. [Debug Information](#DebugInfo)  
    &emsp;&ensp;&nbsp;2.6.2. [Local Tests](#Tests)  
  2.7. [Dependencies](#Dependencies)  
  2.8. [Registration](#Registration)  
  2.9. [Setup Script](#SetupScript)  
    &emsp;&ensp;&nbsp;2.9.1. [Interface](#Interface)  
    &emsp;&ensp;&nbsp;2.9.1. [Setup Form](#SetupForm)  
  2.10. [Release Script](#ReleaseScript)  
  2.11. [Register the project on the machine](#ProjectRegistration)  
  2.12. [Install the Script](#Installation)
3. [Rollout](#Rollout)  

## <a name=DevEnv></a> 1. Development Environment

The guide relies on **Visual Studio**

## <a name=Settings></a> 2. Project Settings

### <a name=ProjectType></a> 2.1 Project Type

.NET class library

### <a name=Framework></a> 2.2 Framework

.NET 4.0 Framework

```
=> Application => Target framwork
```

### <a name=COM></a> 2.3 COM (Component Object Model)

In order for the interfaces to communicate with Kofax, COM visibility must be enabled

```
=> Application => Assembly Information => Make assembly COM-Visible
```

### <a name=Target></a> 2.4 Target Platform

As a 32-bit application, the target platform is x86

```
=> Build => Platform target
```

### <a name=Build></a> 2.5 Build

The files required for Kofax must be stored in the Bin directory of Kofax. This can be found at

```
.....\Program Files (x86)\Kofax\CaptureSS\ServLib\Bin\
```

In order to optimize the development, the output path of the . dll file can be set directly via Visual Studio

```
=> Build => Output path
```

### <a name=Debugging></a> 2.6 Debugging

#### <a name=DebugInfo></a> 2.6.1 Debug Information

To get complete debug information, the option must be specified

```
=> Build => Advanced => Debugging information
```

#### <a name=Tests></a> 2.6.2 Local Tests

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

### <a name=Dependencies></a> 2.7 Dependencies

The **Kofax.ReleaseLib.Interop.dll** must be included in the references

```
=> References => Add Reference => Browse => C:\Program Files (x86)\Kofax\CaptureSS\ServLib\Bin\Kofax.ReleaseLib.Interop.dll
```

### <a name=Registration></a> 2.8 Registration

A [.inf file](https://github.com/matthiashermsen/Kofax-Export-Script-Guide/blob/master/src/LeeresExportScript.inf) is required to install the script. This is then stored in the Kofax Bin directory

**The file must not be created in the UTF-8 format, but must use the UTF-8 without BOM Fortmat**

### <a name=SetupScript></a> 2.9 Setup Script

#### <a name=Interface></a> 2.9.1. Interface

The [setup script](https://github.com/matthiashermsen/Kofax-Export-Script-Guide/blob/master/src/KfxReleaseScriptSetup.cs) is used by the administration module

#### <a name=SetupForm></a> 2.9.2 Setup Form

The setup script starts a [form](https://github.com/matthiashermsen/Kofax-Export-Script-Guide/blob/master/src/FrmSetup.cs)

### <a name=ReleaseScript></a> 2.10. Release Script

The [release script](https://github.com/matthiashermsen/Kofax-Export-Script-Guide/blob/master/src/KfxReleaseScript.cs) is executed during the export

### <a name=ProjectRegistration></a> 2.11. Register the project on the machine

To register the project locally, RegAsm must be run once as administrator in the console

```
"C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" LeeresExportScript.dll /codebase /tlb:LeeresExportScript.tlb
```

### <a name=Installation></a> 2.12. Install the Script

The script can be installed via the administration module

```
Extras Tab => Export Scripts => Add => .inf File from the Kofax Bin Directory
```

## <a name=Rollout></a> 3. Rollout

To deliver the script to the customer, the project . dll and . inf file must be placed in the customer's Kofax Bin directory
