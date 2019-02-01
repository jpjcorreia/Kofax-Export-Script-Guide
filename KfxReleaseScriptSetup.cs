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
