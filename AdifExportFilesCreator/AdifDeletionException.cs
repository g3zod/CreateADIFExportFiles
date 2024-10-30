using System;

namespace AdifExportFilesCreator
{
    /**
     * <summary>
     *   Represents an error where deleting the previous subdirectories and files under the "export" directory failed.
     * </summary>
     */
    public class AdifDeletionException : AdifException
    {
        /**
         * <value>
         *   The directory path that caused the deletion error.
         * </value>
         */
        public string Path { get; }

        /**
         * <summary>
         *   Represents an error where deleting the previous subdirectories and files under the "export" directory failed.
         * </summary>
         */
        public AdifDeletionException()
            : base()
        {
        }

        /**
         * <summary>
         *   Represents an error where deleting the previous subdirectories and files under the "export" directory failed.
         * </summary>
         * 
         * <param name="message">The error message that explains the reason for the exception, or an empty string ("").</param>
         */
        public AdifDeletionException(string message)
            : base(message)
        {
        }

        /**
         * <summary>
         *   Represents an error where deleting the previous subdirectories and files under the "export" directory failed.
         * </summary>
         * 
         * <param name="message">The error message that explains the reason for the exception, or an empty string ("").</param>
         * <param name="innerException">The <see cref="Exception"/> instance that caused the current exception.</param>
         */
        public AdifDeletionException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        /**
         * <summary>
         *   Represents an error where deleting the previous subdirectories and files under the "export" directory failed.
         * </summary>
         * 
         * <param name="message">The error message that explains the reason for the exception, or an empty string ("").</param>
         * <param name="innerException">The <see cref="Exception"/> instance that caused the current exception.</param>
         * <param name="path">The directory path that caused the error.</param>
         */
        public AdifDeletionException(string message, Exception innerException, string path)
            : base(message, innerException)
        {
            Path = path;
        }
    }
}
