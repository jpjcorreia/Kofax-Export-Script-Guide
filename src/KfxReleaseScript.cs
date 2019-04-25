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
                documentData.LogError(0, 0, 0, e.ToString(), "Kofax Capture Custom Export Connector", 0);
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
                foreach (IValue val in documentData.Values)
                {
                    if (val.TableName.IsEmpty())
                    {
                        string sourceName = val.SourceName;
                        string sourceValue = val.Value;

                        switch (val.SourceType)
                        {
                            case KfxLinkSourceType.KFX_REL_INDEXFIELD:
                                // sourceName is the field key
                                // sourceValue is the field value
                                break;

                            case KfxLinkSourceType.KFX_REL_VARIABLE:
                                // sourceName is the field key
                                // sourceValue is the field value
                                break;

                            case KfxLinkSourceType.KFX_REL_BATCHFIELD:
                                // sourceName is the field key
                                // sourceValue is the field value
                                break;
                        }
                    }
                }
                
                return KfxReturnValue.KFX_REL_SUCCESS;
            }
            catch (Exception e)
            {
                documentData.LogError(0, 0, 0, e.ToString(), "Kofax Capture Custom Export Connector", 0);
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
                documentData.LogError(0, 0, 0, e.ToString(), "Kofax Capture Custom Export Connector", 0);
                return KfxReturnValue.KFX_REL_ERROR;
            }
        }

        #endregion Methods
    }
}
