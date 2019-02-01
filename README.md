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

A .inf file is required to install the script. This is then stored in the Kofax Bin directory

**The file must not be created in the UTF-8 format, but must use the UTF-8 without BOM Fortmat**

An example file looks like this

```
[Scripts]
LeeresExportScript

[LeeresExportScript]
SetupModule=LeeresExportScript.dll
SetupProgID=LeeresExportScript.Setup
SetupVersion=1.0.0
ReleaseModule=LeeresExportScript.dll
ReleaseProgID=LeeresExportScript.Release
ReleaseVersion=1.0.0
SupportsNonImageFiles=True
SupportsKofaxPDF=True
RemainLoaded=True
SupportsOriginalFileName=False
SupportsMultipleInstances=True
SupportsCustomEncryption=False
PathSubstitution=ASCII File Name
DisplayName=Leeres Export Script
```

### Setup Script

#### Interface

The setup script is used by the administration module

```csharp
using Kofax.ReleaseLib;
using LeeresExportScript.Extensions;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace LeeresExportScript
{
    /// <summary>
    /// Setup script for configuration
    /// </summary>
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("LeeresExportScript.Setup")]
    public class KfxReleaseScriptSetup
    {
        #region Variables

        /// <summary>
        /// Provides configuration information
        /// </summary>
        public ReleaseSetupData setupData;

        #endregion Variables

        #region Methods

        /// <summary>
        /// Gets called once on setup start
        /// </summary>
        /// <returns>Returns a KfxReturnValue describing whether the action was successful</returns>
        public KfxReturnValue OpenScript()
        {
            try
            {
                return KfxReturnValue.KFX_REL_SUCCESS;
            }
            catch (Exception e)
            {
                setupData.LogError(e.ToString());
                return KfxReturnValue.KFX_REL_ERROR;
            }
        }

        /// <summary>
        /// Gets called on setup close
        /// </summary>
        /// <returns>Returns a KfxReturnValue describing whether the action was successful</returns>
        public KfxReturnValue CloseScript()
        {
            try
            {
                return KfxReturnValue.KFX_REL_SUCCESS;
            }
            catch (Exception e)
            {
                setupData.LogError(e.ToString());
                return KfxReturnValue.KFX_REL_ERROR;
            }
        }

        /// <summary>
        /// Show setup UI and apply the configured configuration data
        /// </summary>
        /// <returns>Returns a KfxReturnValue describing whether the action was successful</returns>
        public KfxReturnValue RunUI()
        {
            FrmSetup frmSetup = new FrmSetup();

            try
            {
                if (frmSetup.ShowSetupForm(ref setupData) == DialogResult.OK)
                {
                    setupData.Apply();
                }

                return KfxReturnValue.KFX_REL_SUCCESS;
            }
            catch (Exception e)
            {
                setupData.LogError(e.ToString());
                return KfxReturnValue.KFX_REL_ERROR;
            }
        }

        /// <summary>
        /// Handle Kofax Action and open up UI on demand
        /// </summary>
        /// <param name="actionID">the ID of the Kofax action</param>
        /// <param name="data1"></param>
        /// <param name="data2"></param>
        /// <returns>Returns a KfxReturnValue describing whether the action was successful</returns>
        public KfxReturnValue ActionEvent(KfxActionValue actionID, string data1, string data2)
        {
            try
            {
                bool showUI = false;

                switch (actionID)
                {
                    case KfxActionValue.KFX_REL_INDEXFIELD_INSERT:
                    case KfxActionValue.KFX_REL_INDEXFIELD_DELETE:
                    case KfxActionValue.KFX_REL_BATCHFIELD_INSERT:
                    case KfxActionValue.KFX_REL_BATCHFIELD_DELETE:
                        showUI = true;
                        break;

                    case KfxActionValue.KFX_REL_UNDEFINED_ACTION:
                    case KfxActionValue.KFX_REL_DOCCLASS_RENAME:
                    case KfxActionValue.KFX_REL_BATCHCLASS_RENAME:
                    case KfxActionValue.KFX_REL_INDEXFIELD_RENAME:
                    case KfxActionValue.KFX_REL_BATCHFIELD_RENAME:
                    case KfxActionValue.KFX_REL_RELEASESETUP_DELETE:
                    case KfxActionValue.KFX_REL_IMPORT:
                    case KfxActionValue.KFX_REL_UPGRADE:
                    case KfxActionValue.KFX_REL_PUBLISH_CHECK:
                    case KfxActionValue.KFX_REL_START:
                    case KfxActionValue.KFX_REL_END:
                    case KfxActionValue.KFX_REL_FOLDERCLASS_INSERT:
                    case KfxActionValue.KFX_REL_FOLDERCLASS_RENAME:
                    case KfxActionValue.KFX_REL_FOLDERCLASS_DELETE:
                    case KfxActionValue.KFX_REL_TABLE_DELETE:
                    case KfxActionValue.KFX_REL_TABLE_INSERT:
                    case KfxActionValue.KFX_REL_TABLE_RENAME:
                    default:
                        break;
                }

                if (showUI)
                {
                    return RunUI();
                }

                return KfxReturnValue.KFX_REL_SUCCESS;
            }
            catch (Exception e)
            {
                setupData.LogError(e.ToString());
                return KfxReturnValue.KFX_REL_ERROR;
            }
        }

        #endregion Methods
    }
}
```

#### Setup Form

The setup script starts a form

```csharp
using Kofax.ReleaseLib;
using System;
using System.Windows.Forms;

namespace LeeresExportScript
{
    /// <summary>
    /// Setup configuration UI
    /// </summary>
    public partial class FrmSetup : Form
    {
        #region Constructor

        /// <summary>
        /// Default Constructor
        /// </summary>
        public FrmSetup()
        {
            InitializeComponent();
        }

        #endregion Constructor

        #region Variables

        /// <summary>
        /// Provides configuration information
        /// </summary>
        private ReleaseSetupData releaseSetupData;

        #endregion Variables

        #region Methods

        /// <summary>
        /// Creates a reference to the setup data and initializes the UI
        /// </summary>
        /// <param name="objectSetupData">ReleaseSetupData</param>
        /// <returns></returns>
        public DialogResult ShowSetupForm(ref ReleaseSetupData objectSetupData)
        {
            releaseSetupData = objectSetupData;
            return ShowDialog();
        }

        /// <summary>
        /// Closes the UI and returns the configuration data to the setup script
        /// </summary>
        private void ApplySettings()
        {
            if (VerifySettings())
            {
                SaveSettings();

                DialogResult = DialogResult.OK;
                Close();
            }
        }

        /// <summary>
        /// Stores all the configuration information
        /// </summary>
        private void SaveSettings()
        {
            releaseSetupData.CustomProperties.RemoveAll();

            // releaseSetupData.CustomProperties.Add("key", "value");

            releaseSetupData.Links.RemoveAll();

            foreach (IndexField indexField in releaseSetupData.IndexFields)
            {
                releaseSetupData.Links.Add(indexField.Name, KfxLinkSourceType.KFX_REL_INDEXFIELD, indexField.Name);
            }

            foreach (BatchField batchField in releaseSetupData.BatchFields)
            {
                releaseSetupData.Links.Add(batchField.Name, KfxLinkSourceType.KFX_REL_BATCHFIELD, batchField.Name);
            }

            foreach (dynamic batchVariable in releaseSetupData.BatchVariableNames)
            {
                releaseSetupData.Links.Add(batchVariable, KfxLinkSourceType.KFX_REL_VARIABLE, batchVariable);
            }
        }

        /// <summary>
        /// Validate the configuration
        /// </summary>
        /// <returns>Returns a bool describing the validation</returns>
        private bool VerifySettings()
        {
            return true;
        }

        #endregion Methods

        #region Events

        /// <summary>
        /// Close the setup UI and cancel the setup script
        /// </summary>
        /// <param name="sender">Sender</param>
        /// <param name="e">EventArgs</param>
        private void btnCancelSetup_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        #endregion Events
    }
}
```

### Release Script

The release script is executed during the export

```csharp
using Kofax.ReleaseLib;
using LeeresExportScript.Extensions;
using System;
using System.Runtime.InteropServices;

namespace LeeresExportScript
{
    /// <summary>
    /// Release script for data transfer to external data repositories
    /// </summary>
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("LeeresExportScript.Release")]
    public class KfxReleaseScript
    {
        #region Variables

        /// <summary>
        /// Reads the configuration information and saves the data to an external data repository
        /// </summary>
        public ReleaseData documentData;

        #endregion Variables

        #region Methods

        /// <summary>
        /// Gets called once on release start
        /// </summary>
        /// <returns>Returns a KfxReturnValue describing whether the action was successful</returns>
        public KfxReturnValue OpenScript()
        {
            try
            {
                return KfxReturnValue.KFX_REL_SUCCESS;
            }
            catch (Exception e)
            {
                documentData.LogError(e.ToString());
                return KfxReturnValue.KFX_REL_ERROR;
            }
        }

        /// <summary>
        /// Gets called once per each document to release
        /// </summary>
        /// <returns>Returns a KfxReturnValue describing whether the action was successful</returns>
        public KfxReturnValue ReleaseDoc()
        {
            try
            {
                return KfxReturnValue.KFX_REL_SUCCESS;
            }
            catch (Exception e)
            {
                documentData.LogError(e.ToString());
                return KfxReturnValue.KFX_REL_ERROR;
            }
        }

        /// <summary>
        /// Gets called once on release close
        /// </summary>
        /// <returns>Returns a KfxReturnValue describing whether the action was successful</returns>
        public KfxReturnValue CloseScript()
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                return KfxReturnValue.KFX_REL_SUCCESS;
            }
            catch (Exception e)
            {
                documentData.LogError(e.ToString());
                return KfxReturnValue.KFX_REL_ERROR;
            }
        }

        #endregion Methods
    }
}
```

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
