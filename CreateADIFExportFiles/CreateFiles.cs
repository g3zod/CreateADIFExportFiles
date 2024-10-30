using CreateADIFExportFiles.Properties;
using System;
using System.IO;
using System.Windows.Forms;
using AdifReleaseLib;
using AdifExportFilesCreator;
using System.Diagnostics.CodeAnalysis;

namespace CreateADIFExportFiles
{
    /**
     * <summary>
     *   Pprovides the user interface for the CreateADIFExportFiles application that
     *   exports tables from an ADIF Specification.<br/>
     *   <br/>
     *   The file formats are CSV (.csv), TSV (.tsv), XML (.xml). Excel (.xlsx) and OpenSource Calc (.ods).
     * </summary>
     */
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "This is the Windows UI")]
    public partial class CreateFiles : Form
    {
        private const string MessageBoxTitle = "Create ADIF Export Files";
        private readonly Settings settings = Settings.Default;
        private bool
            cancel = false,
            exporting = false;

        /**
         * <summary>
         *   Starts a log file in the Documents folder and upgrades the settings if required.<br/>
         *   If the program version is more recent than the one in the Settings.Default file, it upgrades the file.<br/>
         *   Initializes the form.
         * </summary>
         */
        public CreateFiles()
        {
            try
            {
                Logger.StartLog(Application.ProductName, Application.ProductVersion);
            }
            catch (Exception exc)
            {
                // No point in logging this if the Logger did not start up.
                _ = MessageBox.Show($"Unable to start log file.\r\n\r\n{exc.Message}", MessageBoxTitle);
                // Continue because failure to create a log is not disastrous.
            }

            try
            {
                string
                    settingsVersion = settings.Version,
                    thisVersion = Application.ProductVersion;

                decimal
                    settingsVersionAsDecimal = Common.VersionToDecimal(settingsVersion),
                    thisVersionAsDecimal = Common.VersionToDecimal(thisVersion);

                if (settingsVersionAsDecimal != thisVersionAsDecimal)
                {
                    if (settingsVersionAsDecimal < thisVersionAsDecimal)
                    {
                        Logger.Log($"Upgrading Settings from version {settingsVersion} to {thisVersion}");
                        settings.Upgrade();
                    }
                    else
                    {
                        Logger.Log($"Resetting the version in settings from {settingsVersion} to {thisVersion} because it is more recent");
                    }
                    settings.Version = thisVersion;
                    settings.Save();
                }
            }
            catch (Exception exc)
            {
                Logger.Log(exc);
                _ = MessageBox.Show($"Error while upgrading version.\r\n\r\n{exc.Message}", MessageBoxTitle);
                Logger.Close();
                Environment.Exit(1);
            }

            InitializeComponent();
        }

        /**
         * <summary>
         *   Enables or disables the Form's controls.
         * </summary>
         */
        private void CreateFiles_Load(object sender, EventArgs e)
        {
            try
            {
                EnableControls(true);
            }
            catch (Exception exc)
            {
                Logger.Log(exc);
                _ = MessageBox.Show($"Error while loading.\r\n\r\n{exc.Message}", MessageBoxTitle);
                Logger.Close();
                Environment.Exit(1);
            }
        }

        /**
         * <summary>
         *   Shows a File Open dialog in response to the Form's "..." button.
         * </summary>
         */
        private void BtnChooseFile_Click(object sender, EventArgs e)
        {
            try
            {
                /*
                 * Directory and file paths used in this application:
                 * 
                 *  <ijk>                                               (ADIF version number i.j.k)
                 *      ADIF_<ijk>.htm                                  (ADIF Specification version i.j.k)
                 *      ADIF_<ijk>_annotated.htm                        (Annotated ADIF Specification version i.j.k)
                 *      exports
                 *          csv                                         (Comma-separated values files)
                 *              datatypes.csv
                 *              enumerations.csv
                 *              enumerations_<name>.csv
                 *              ...
                 *              fields.csv
                 *          ods                                         (Apache(TM) OpenOffice(TM) Calc files)
                 *              datatypes.ods
                 *              enumerations.ods
                 *              enumerations_<name>.ods
                 *              ...
                 *              fields.ods
                 *          tsv                                         (Tab-separated values files)
                 *              datatypes.tsv
                 *              enumerations.tsv
                 *              enumerations_<name>.tsv
                 *              ...
                 *              fields.tsv
                 *          xlsx                                        (Microsoft(TM) Excel(TM) files)
                 *              datatypes.xlsx
                 *              enumerations.xlsx
                 *              enumerations_<name>.xlsx
                 *              ...
                 *              fields.xlsx
                 *          xml                                         (Extensible Markup Language files)
                 *              all.xml
                 *              datatypes.xml
                 *              enumerations.xml
                 *              enumerations_<name>.xml
                 *              ...
                 *              fields.xml
                 */

                {
                    string directoryPath;

                    try
                    {
                        directoryPath = Path.GetDirectoryName(TxtSpecificationFile.Text);
                    }
                    catch
                    {
                        directoryPath = string.Empty;
                    }

                    OfdSpecificationFile.InitialDirectory =
                        (!string.IsNullOrEmpty(directoryPath)) &&
                        Directory.Exists(directoryPath) ?
                            directoryPath :
                            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }
                OfdSpecificationFile.DefaultExt = "htm";
                OfdSpecificationFile.Filter = "XHTML files (*.htm)|*.htm|All files (*.*)|*.*";
                OfdSpecificationFile.FileName = File.Exists(TxtSpecificationFile.Text) ?
                    TxtSpecificationFile.Text :
                    string.Empty;

                if (OfdSpecificationFile.ShowDialog() == DialogResult.OK)
                {
                    TxtSpecificationFile.Text = OfdSpecificationFile.FileName;
                    settings.Save();
                }
            }
            catch (Exception exc)
            {
                Logger.Log(exc);
                ShowError($"Error choosing ADIF Specification file.\r\n\r\n{exc.Message}");
            }
        }

        /**
         * <summary>
         *   Exports the tables from the ADIF Specification in response to the Create Files button.
         * </summary>
         */
        private void BtnCreateFiles_Click(object sender, EventArgs e)
        {
            string specificationFullFilePath = string.Empty;

            try
            {
                exporting = true;
                EnableControls(false);

                specificationFullFilePath = Path.GetFullPath(TxtSpecificationFile.Text);

                Cursor = Cursors.WaitCursor;
                new Specification(
                    specificationFullFilePath,
                    Application.StartupPath,
                    new Specification.ProgressReporter(ShowProgress),
                    new Specification.UserPrompter(PromptUser)).Export();
            }
            catch (AdifAnnotatedSpecificationException)
            {
                Logger.Log("Export of annotated ADIF Specification cancelled by user");
                TsslProgress.Text = string.Empty;
            }
            catch (AdifDeletionException exc)
            {
                Logger.Log(exc);

                ShowError(
$@"Failed to delete all previous directories and files.

Please delete them manually and try again:
{Path.GetDirectoryName(specificationFullFilePath)}\exports

Error details: {(exc.InnerException != null ? exc.InnerException.Message : "None")}");

                TsslProgress.Text = string.Empty;
            }
            catch (Exception exc)
            {
                Logger.Log(exc);

                string innerExceptionMessage = exc.InnerException != null ? $"\r\n{exc.InnerException.Message}" : string.Empty;

                ShowError($"Error exporting ADIF Specification files.\r\n\r\n{exc.Message}{innerExceptionMessage}");
                TsslProgress.Text = string.Empty;
            }
            finally
            {
                try
                {
                    exporting = false;
                    Cursor = Cursors.Default;
                    EnableControls(true);
                }
                catch { }
            }
        }

        /**
         * <summary>
         *   Displays a progress message in the status bar and logs the message.&#160;
         *   Checks to see if the Close button has been clicked and if so, ends the program.
         * </summary>
         * 
         * <param name="message">The message to be be shown.</param>
         */
        private void ShowProgress(string message)
        {
            TsslProgress.Text = Common.ToSingleLine(message); ;
            Logger.Log(message);
            SsProgress.Refresh();
            Application.DoEvents();
            if (cancel)
            {
                Logger.Log("Program exit while exporting due to the Close button or window Close box/menu item being clicked.");
                Logger.Close();
                Environment.Exit(0);
            }
        }

        /**
         * <summary>
         *   Outputs a Yes/No MessageBox and returns the response to the caller.
         * </summary>
         * 
         * <param name="caption">The text to display.</param>
         * <param name="text">A caption to display.</param>
         * <value>The Yes or No response.</value>
         */
        private Specification.UserPrompterReply PromptUser(string text, string caption)
        {
            return MessageBox.Show(text, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes ?
                Specification.UserPrompterReply.Yes :
                Specification.UserPrompterReply.No;
        }

        /**
         * <summary>
         *   Enables or disables some of the user interface controls.
         * </summary>
         */
        private void EnableControls(bool enable)
        {
            LblSpecificationFile.Enabled =
            TxtSpecificationFile.Enabled =
            BtnChooseFile.Enabled = enable;

            if (enable)
            {
                BtnCreateFiles.Enabled = File.Exists(TxtSpecificationFile.Text);
            }
            else
            {
                BtnCreateFiles.Enabled = false;
            }
        }

        /**
         * <summary>
         *   Shows a Message Box contaiing details of an error.
         * </summary>
         */
        private void ShowError(string message)
        {
            _ = MessageBox.Show(
                message,
                "Error Exporting ADIF Specification Files",
                MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
        }

        /**
         * <summary>
         *   Closes the program in response to the Close button.&#160;
         *   Sets the 'cancel' flag so that if an export is in progress, the program can be ended at the next call to the <see cref="ShowProgress"/> method.
         * </summary>
         */
        private void BtnClose_Click(object sender, EventArgs e)
        {
            try
            {
                Close();
                cancel = true;
            }
            catch (Exception exc)
            {
                Logger.Log(exc);
                ShowError($"Error closing application.\r\n\r\n{exc.Message}");
            }
        }

        /**
         * <summary>
         *   Closes the log file if no export is in progress.
         * </summary>
         */
        private void CreateFiles_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (!exporting) Logger.Close();
        }

        /**
         * <summary>
         *   Sets the flag to close the program if it is currently  exporting.&#160;
         *   This is necessary when the windows close box is clicked.
         * </summary>
         */
        private void CreateFiles_FormClosing(object sender, FormClosingEventArgs e)
        {
            cancel = true;
        }

        /**
         * <summary>
         *   Enables or disables the <see cref="BtnCreateFiles"/> button depending on whether the file in <see cref="TxtSpecificationFile"/> exists.
         * </summary>
         */
        private void TxtSpecificationFile_TextChanged(object sender, EventArgs e)
        {
            try
            {
                bool fileExists = File.Exists(TxtSpecificationFile.Text);

                BtnCreateFiles.Enabled = fileExists;

                if (fileExists)
                {
                    settings.TxtSpecificationFile = TxtSpecificationFile.Text;
                    settings.Save();
                }
            }
            catch { }
        }
    }
}
