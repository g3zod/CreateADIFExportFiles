using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace AdifReleaseLib
{
    /**
     * <summary>
     *   This class is used to create an tab-separated values (.tsv) file containing table(s) from the ADIF Specification.
     * </summary>
     * 
     * <remarks>
     *   Any tabs that might be present in the ADIF Specification HTML file will be replaced with space characters.
     * </remarks>
     */
    public class TsvWriter : IDisposable, ITableWriter
    {
        private readonly string fileName;
        private readonly StreamWriter streamWriter;
        private readonly StringBuilder recordBuillder;

        private TsvWriter()
        {
            // The parameterless constructor is not available.
        }

        /**
         * <summary>
         *   Creates a TSB file and opens a StreamWriter object.
         * </summary>
         * 
         * <param name="filePath">The output TSV file's path.</param>
         */
        public TsvWriter(string filePath)
        {
            fileName = filePath;
            streamWriter = new StreamWriter(filePath, false, System.Text.Encoding.UTF8);
            recordBuillder = new StringBuilder(2048);
        }

        /**
         * <inheritdoc cref="ITableWriter.WriteField(string)"/>
         */
        public void WriteField(string value)
        {
            if (recordBuillder.Length != 0)
            {
                recordBuillder.Append('\t');
            }
            recordBuillder.Append(value);
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
            streamWriter.WriteLine(recordBuillder.ToString());
            recordBuillder.Length = 0;
        }

        /**
         * <inheritdoc cref="ITableWriter.WriteHeader"/>
         */
        public void WriteHeader()
        {
            WriteRecord();
        }

        /**
         * <inheritdoc cref="ITableWriter.Close"/>
         */
        public void Close()
        {
            // Cater for missing call to WriteRecord
            if (recordBuillder.Length > 0)
            {
                WriteRecord();
            }
            streamWriter.Close();
            streamWriter.Dispose();

            Common.SetFileTimesToNow(fileName);
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
        ~TsvWriter()      
        {
            // Do not re-create Dispose clean-up code here.
            // Calling Dispose(false) is optimal in terms of
            // readability and maintainability.
            Dispose(false);
        }
    }
}
