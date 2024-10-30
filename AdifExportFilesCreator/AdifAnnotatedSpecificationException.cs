using System;

namespace AdifExportFilesCreator
{
    /**
     * <summary>
     *   Represents an error where the user has chosen not to export from an annotated ADIF Specification.
     * </summary>
     */
    public class AdifAnnotatedSpecificationException : AdifException
    {
        /**
         * <summary>
         *   Represents an error where the user has chosen not to export from an annotated ADIF Specification.<br/>
         *   <br/>
         *   This is normally unacceptable because an annotated version contains deleted text that can
         *   probably surface in the exported files.
         * </summary>
         */
        public AdifAnnotatedSpecificationException()
            : base("Export cancelled by user.")
        {
        }

        /**
         * <summary>
         *   Represents an error where the user has chosen not to export from an annotated ADIF Specification.
         * </summary>
         * 
         * <param name="message">The error message that explains the reason for the exception, or an empty string ("").</param>
         */
        public AdifAnnotatedSpecificationException(string message)
            : base(message)
        {
        }

        /**
         * <summary>
         *   Represents an error where the user has chosen not to export from an annotated ADIF Specification.
         * </summary>
         * 
         * <param name="message">The error message that explains the reason for the exception, or an empty string ("").</param>
         * <param name="innerException">The <see cref="Exception"/> instance that caused the current exception.</param>
         */
        public AdifAnnotatedSpecificationException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
