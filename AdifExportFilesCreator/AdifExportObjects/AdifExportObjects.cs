using System;
using System.Collections.Generic;

namespace AdifExportFilesCreator
{
    /**
     * <summary>
     *   Contains objects that can be used to serialize and deserialize the ADIF Specifications tables as JSON.
     * </summary>
     */
    namespace AdifExportObjects
    {
        /**
         * <summary>
         *   This is the top-level class that allows an object to be exported with the name "Adif" rather than being anonymous.
         * </summary>
         */
        public class Export
        {
            public Adif Adif { get; set; }
        }

        /**
         * <summary>
         *   This contains properties representing basic information about the ADIF Specification
         *   and lists of ADIF datatypes, enumerations, and fields.
         * </summary>
         */
        public class Adif
        {
            public string Version { get; set; }
            public string Status { get; set; }
            public DateTime Date { get; set; }
            public DataTypes DataTypes { get; set; }
            public Enumerations Enumerations { get; set; }
            public Fields Fields { get; set; }
        }

        /**
         * <summary>
         *   This represents the ADIF datatypes table.
         * </summary>
         */
        public class DataTypes
        {
            public Header Header { get; set; }
            public Records Records { get; set; }
            public Records2 Records2 { get; set; }
        }

        /**
         * <summary>
         *   This represents the full set of ADIF enumeration tables.
         * </summary>
         */
        public class Enumerations : Dictionary<string, Enumeration> { }

        /**
         * <summary>
         *   This represents the ADIF field header and QSO tables.
         * </summary>
         */
        public class Fields
        {
            public Header Header { get; set; }
            public Records Records { get; set; }
            public Records2 Records2 { get; set; }
        }

        /**
         * <summary>
         *   This represents an individual ADIF enumeration table.
         * </summary>
         */
        public class Enumeration
        {
            public Header Header { get; set; }
            public Records Records { get; set; }
            public Records2 Records2 { get; set; }
        }

        /**
         * <summary>
         *   This represents a list of titles used in an individual table's header record.
         * </summary>
         */
        public class Header : List<string> { }

        /**
         * <summary>
         *   This represents an un-named list of records.
         * </summary>
         */
        public class Records : List<Record> { }

        /**
         * <summary>
         *   This represents a named list of records.
         * </summary>
         */
        public class Records2 : Dictionary<string, Record> { }

        /**
         * <summary>
         *   This represents an individual record as a key/value pair of strings in a dictionary.
         * </summary>
         */
        public class Record : Dictionary<string, string> { }
    }
}
