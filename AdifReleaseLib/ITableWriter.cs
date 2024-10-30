using System.Collections.Generic;

namespace AdifReleaseLib
{
    /**
     * <summary>
     *   This interface contains the methods used to write an ADIF Specification's table cells to CSV, TSV, Excel, and Calc files.
     * </summary>
     * 
     * <remarks>
     *   It is implemented by these classes:
     *   <list type="bullet">
     *      <item><description><see cref="CsvWriter"/></description></item>
     *      <item><description><see cref="ExcelWriter"/></description></item>
     *      <item><description><see cref="TsvWriter"/></description></item>
     *   </list>
     * </remarks>
     */
    public interface ITableWriter
    {
        /**
         * <summary>
         *   Writes a single value to the file.
         * </summary>
         * 
         * <param name="field">The value to be written to the file.</param>
         */
        void WriteField(string field);

        /**
         * <summary>
         *   Writes one or more values from a <see cref="string"/>[] to a file.
         * </summary>
         * 
         * <param name="fields">A <see cref="string"/>[] containing the values to be written to the file.</param>
         */
        void WriteFields(string[] fields);

        /**
         * <summary>
         *   Writes one or more values from a List&lt;<see cref="string"/>&gt; to a file.
         * </summary>
         * 
         * <param name="fields">A List&lt;<see cref="string"/>&gt; object containing the values to be written to the file.</param>
         */
        void WriteFields(List<string> fields);

        /**
         * <summary>
         *   Signals that a record is ready to be written to the file.
         * </summary>
         */
        void WriteRecord();

        /**
         * <summary>
         *   Signals that a header record is ready to be written to the file.
         * </summary>
         */
        void WriteHeader();

        /**
         * <summary>
         *   Writes any outstanding values to the file and then closes it.
         * </summary>
         */
        void Close();
    }
}
