using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace AdifReleaseLib
{
    /**
     * <summary>
     *   This class is used to create an Excel (.xlsx) and OpenOffice Calc (.ods) file containing table(s) from the ADIF Specification.
     * </summary>
     */
    public class ExcelWriter : IDisposable, ITableWriter
    {
        private readonly string xlsxFileName;
        private readonly string odsFileName;
        private static Excel.Application application;
        private readonly Excel.Workbook workBook;
        private readonly Excel.Worksheet workSheet;
        private int rowNumber = 1;
        private readonly List<string> recordBuilder;

        static readonly object missing = Type.Missing;

        private ExcelWriter()
        {
            // The parameterless constructor is not available.
        }

        /**
         * <summary>
         *   This is used to create and write to Excel (.xlsx) and OpenSource Calc (.ods) files containing a table from the ADIF Specification.
         * </summary>
         * 
         * <remarks>
         *   <para>
         *      It works by creating an Excel workbook and saving it both as an Excel (.xls) file
         *      and as a Calc (.ods) file.
         *   </para>
         *   <para>
         *      A version of Excel must be installed on the computer.
         *   </para>
         * </remarks>
         * 
         * <param name="xlsxFileName">The Excel file path.</param>
         * <param name="odsFileName">The Calc file path.</param>
         * <param name="adifVersion">The ADIF Specification version number.</param>
         * <param name="adifStatus">The ADIF Specification status.</param>
         * <param name="title">A title to be included in the file.</param>
         * <param name="author">The author to be included in the file.</param>
         * <param name="sheetName">The name of the worksheet within the workbook.</param>
         */
        public ExcelWriter(
            string xlsxFileName,
            string odsFileName,
            string adifVersion,
            string adifStatus,
            string title,
            string author,
            string sheetName)
        {
            this.xlsxFileName = xlsxFileName;
            this.odsFileName  = odsFileName;
            if (application == null)
            {
                application = new Excel.Application();
            }
            workBook = application.Workbooks.Add(missing);
            workBook.EnableAutoRecover = false;
            workBook.Title = title;
            workBook.Author = author;
            WriteCustomProperty("ADIF Version", adifVersion);
            WriteCustomProperty("ADIF Status", adifStatus);
            workSheet = (Excel.Worksheet)application.Worksheets[1];
            workSheet.Name = sheetName.Length > 31 ?
                sheetName.Substring(0, 31) :
                sheetName;
            recordBuilder = new List<string>(50);
        }

        /**
         * <summary>
         *   Writes a custom document property.
         * </summary>
         * 
         * <remarks>
         *   Custom document properties are visible in Excel by clicking on File, Info, then Properties and choose Advanced Properties.&#160;
         *   Finally, click on the Custom tab.&#160; (Custom properties are likely not to of much value other than if the workbook is
         *   accessed programatically.&#160; Also, the ADIF Specification version and status are already included in the Title property.)
         * </remarks>
         * 
         * <param name="name">The custom document tag's name.</param>
         * <param name="value">The custom document tag's value.</param>
         */
        private void WriteCustomProperty(string name, string value)
        {
            object oDocCustomProps = workBook.CustomDocumentProperties;
            Type typeDocCustomProps = oDocCustomProps.GetType();

            object[] oArgs = {
                name,
                false,
                Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                value};

            typeDocCustomProps.InvokeMember(
                "Add",
                System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod,
                null,
                oDocCustomProps,
                oArgs);
        }

        /**
         * <inheritdoc cref="ITableWriter.WriteField(string)"/>
         */
        public void WriteField(string value)
        {
            recordBuilder.Add(value);
        }

        /**
         * <inheritdoc cref="ITableWriter.WriteFields(string[])"/>
         */
        public void WriteFields(string[] fields)
        {
            foreach (string field in fields)
            {
                WriteField(field);
            }
        }

        /**
         * <inheritdoc cref="ITableWriter.WriteFields(List{string})"/>
         */
        public void WriteFields(List<string> fields)
        {
            foreach (string field in fields)
            {
                WriteField(field);
            }
        }

        /**
         * <inheritdoc cref="ITableWriter.WriteRecord"/>
         */
        public void WriteRecord()
        {
            Excel.Range rng;

            rng = (Excel.Range)workSheet.Cells[rowNumber++, 1];
            rng = rng.get_Resize(1, recordBuilder.Count);
            rng.NumberFormat = "@";

            try
            {
                rng.Value2 = recordBuilder.ToArray();
            }
            catch
            {
                // There is a known limitation that is not cleared until Excel 2007 of 911 characters per cell.

                const int maxLength = 911;

                string item2 = recordBuilder[2];

                if (item2.Length > maxLength)
                {
                    const string suffix = " ...";

                    System.Diagnostics.Debug.WriteLine(string.Format(
                         "WriteRecord exception: length: {0}, changing to {1}",
                         item2.Length.ToString(),
                         maxLength.ToString()));

                    recordBuilder[2] = item2.Substring(0, maxLength - suffix.Length) + suffix;
                    rng.Value2 = recordBuilder.ToArray();
                }
                else
                {
                    throw;
                }
            }
            recordBuilder.Clear();
        }

        /**
         * <inheritdoc cref="ITableWriter.WriteHeader"/>
         */
        public void WriteHeader()
        {
            Excel.Range rng;

            rng = (Excel.Range)workSheet.Cells[rowNumber++, 1];
            rng = rng.get_Resize(1, recordBuilder.Count);
            rng.NumberFormat = "@";
            rng.Value2 = recordBuilder.ToArray();
            rng.Font.Bold = true;
            recordBuilder.Clear();
        }

        /**
         * <inheritdoc cref="ITableWriter.Close"/>
         */
        public void Close()
        {
            const int
                 xlOpenXMLWorkbook          = 51,  // xlsx
                 xlOpenDocumentSpreadsheet  = 60;  // ods

            int format;

            bool
                xlsxWritten,
                odsWritten  = false;

            // Cater for missing call to WriteRecord
            if (recordBuilder.Count > 0)
            {
                WriteRecord();
            }

            format = xlOpenXMLWorkbook;

            workBook.SaveAs(
                    xlsxFileName,                           //  Filename
                    format,                                 //  FileFormat
                    missing,                                //  Password
                    missing,                                //  WriteResPassword
                    missing,                                //  ReadOnlyRecommended
                    missing,                                //  CreateBackup
                    Excel.XlSaveAsAccessMode.xlExclusive,   //  AccessMode
                    missing,                                //  ConflictResolution
                    false,                                  //  AddToMru
                    missing,                                //  TextCodepage
                    missing,                                //  TextVisualLayout
                    missing);                               //  Local

            xlsxWritten = true;
            //another();

            try
            {
                // This will not work on older versions of Excel.

                format = xlOpenDocumentSpreadsheet;

                workBook.SaveAs(
                        odsFileName,                            //  Filename
                        format,                                 //  FileFormat
                        missing,                                //  Password
                        missing,                                //  WriteResPassword
                        missing,                                //  ReadOnlyRecommended
                        missing,                                //  CreateBackup
                        Excel.XlSaveAsAccessMode.xlExclusive,   //  AccessMode
                        missing,                                //  ConflictResolution
                        false,                                  //  AddToMru
                        missing,                                //  TextCodepage
                        missing,                                //  TextVisualLayout
                        missing);                               //  Local

                odsWritten = true;
            }
            catch
            {
            }

            workBook.Close(false, missing, missing);

            {
                // File.Delete does not delete files synchronously, so that typically the
                // file that has just been created is actually an old one that has been
                // overridden.  To compensate for the confusing "Date created" date,
                // set all the file's date manually.

                if (xlsxWritten)
                {
                    try
                    {
                        Common.SetFileTimesToNow(xlsxFileName);
                    }
                    catch
                    {
                        // This crashes with a permissions fail on Windows 10 now, so just ignore.
                    }
                }
                if (odsWritten)
                {
                    try
                    {
                        Common.SetFileTimesToNow(odsFileName);
                    }
                    catch
                    {
                        // This crashes with a permissions fail on Windows 10 now, so just ignore.
                    }
                }
            }

            //application.Quit();
        }

        /**
         * <summary>
         *   Terminates the Excel process.
         * </summary>
         * 
         * <remarks>
         *   This is for use only when the program has finished i.e. not each time the files are written.
         * </remarks>
         */
        public static void End()
        {
            if (application != null)
            {
                application.Quit();
                application = null;
            }
        }

        bool disposed = false;

        /**
         * <summary>
         *   Standard disposal pattern code from MSDN.
         * </summary>
         */
        public void Dispose()
        {
            Dispose(true);
            // This object will be cleaned up by the Dispose method.
            // Therefore, you should call GC.SupressFinalize to
            // take this object off the finalization queue 
            // and prevent finalization code for this object
            // from executing a second time.
            GC.SuppressFinalize(this);
        }

        // Dispose(bool disposing) executes in two distinct scenarios.
        // If disposing equals true, the method has been called directly
        // or indirectly by a user's code. Managed and unmanaged resources
        // can be disposed.
        // If disposing equals false, the method has been called by the 
        // runtime from inside the finalizer and you should not reference 
        // other objects. Only unmanaged resources can be disposed.

        /**
         * <summary>
         *   Standard disposal pattern code from MSDN.
         * </summary>
         */
        private void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if(!this.disposed)
            {
                // If disposing equals true, dispose all managed 
                // and unmanaged resources.
                if(disposing)
                {
                    // Dispose managed resources.
                    Close();
                }
    		 
                // Call the appropriate methods to clean up 
                // unmanaged resources here.
                // If disposing is false, 
                // only the following code is executed.
            }
            disposed = true;         
        }

        // Use C# destructor syntax for finalization code.
        // This destructor will run only if the Dispose method 
        // does not get called.
        // It gives your base class the opportunity to finalize.
        // Do not provide destructors in types derived from this class.

        /**
         * <summary>
         *   Standard disposal pattern code from MSDN.
         * </summary>
         */
        ~ExcelWriter()      
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of
            // readability and maintainability.
            Dispose(false);
        }
    }
}
