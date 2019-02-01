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
            
            try
            {
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
            catch (Exception e)
            {
                // duplicate key exception
                throw;
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
