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
