/*
 * This class exports the ADIF Specification XHTML file as associated files for including as part of the release of a
 * new version of the ADIF Specification.
 * 
 * The files comprise CSV (.csv), TSV (.tsv), XML (.xml), Excel (.xlsx), and OpenOffice Calc (.ods) files containing
 * data types, enumerations, and fields.
 * 
 * The files are stored in subdirectories within the directory that contains the source ADIF Specification XHTML file
 * (e.g. for ADIF version 3.0.6 the directory name is 306):
 * 
 *    exports                               - Directory for sub-directories based on file type (e.g. .csv).
 * 
 *    exports/csv                           - Directory for CSV (.csv) files.
 *    exports/csv/datatypes.csv             - CSV file containing the data types.
 *    exports/csv/enumerations.csv          - CSV file containing the enumerations.
 *    exports/csv/enumerations_{name}.csv   - CSV file containing the enumeration with the name {name}.
 *    exports/csv/fields.csv                - CSV file containing the fields.
 * 
 *    exports/tsv                           - Directory for TSV (.tsv) files.
 *    exports/tsv/datatypes.tsv             - TSV file containing the data types.
 *    exports/tsv/enumerations.tsv          - TSV file containing the enumerations.
 *    exports/tsv/enumerations_{name}.tsv   - TSV file containing the enumeration with the name {name}.
 *    exports/tsv/fields.tsv                - TSV file containing the fields.
 * 
 *    exports/xlsx                          - Directory for Excel (.xlsx) files.
 *    exports/xlsx/datatypes.xlsx           - Excel file containing the data types.
 *    exports/xlsx/enumerations.xlsx        - Excel file containing the enumerations.
 *    exports/xlsx/enumerations_{name}.xlsx - Excel file containing the enumeration with the name {name}.
 *    exports/xlsx/fields.xlsx              - Excel file containing the fields.
 * 
 *    exports/ods                           - Directory for OpenOffice Calc (.ods) files.
 *    exports/ods/datatypes.ods             - OpenOffice Calc file containing the data types.
 *    exports/ods/enumerations.ods          - OpenOffice Calc file containing the enumerations.
 *    exports/ods/enumerations_{name}.ods   - OpenOffice Calc file containing the enumeration with the name {name}.
 *    exports/ods/fields.ods                - OpenOffice Calc file containing the fields.
 * 
 *    exports/xml                           - Directory for XML (.xml) files.
 *    exports/xml/all.xml                   - XML file containing the data types, enumerations, and fields.
 *    exports/xml/datatypes.xml             - XML file containing the data types.
 *    exports/xml/enumerations.xml          - XML file containing the enumerations.
 *    exports/xml/enumerations_{name}.xml   - XML file containing the enumeration with the name {name}.
 *    exports/xml/fields.xml                - XML file containing the fields.
 *
 * Notes:
 * 
 *  * The CSV, TSV, and XML files are encoded as UTF-8 with a byte order mark (BOM) of 0xEF, 0xBB, 0xBF at the
 *    start of the file.  These 3 bytes can be ignored but are included for compatibility with Microsoft software.
 * 
 *  * CSV file values are enclosed by double quotes (") and any double quotes embedded within the value are
 *    encoded as a pair of double quotes.  E.g.
 *        "This value contains a double quotes "" character"
 * 
 *  * TSV file values are separated by tab characters.  A tab character will never occur within a value.
 *  
 *  * The enumeration file names for individual enumerations are lowercase and have spaces replaced by underscores.
 *    E.g. the ARRL Sections table is exported as the CSV file enumeration_arrl_sections.csv.
 * 
 *  * Within a file, enumeration names have spaces replaced by underscores but the original case of the names
 *    is preserved e.g. ARRL_Section
 * 
 *  * All the XML files have the same overall structure as the all.xml file, the differences being that:
 *    -  The datatypes.xml file omits the <enumerations> and <fields> elements. 
 *    -  The enumerations.xml and named enumeration XML files omit the <dataTypes> and <fields> elements. 
 *    -  The fields.xml file omits the <dataTypes> and <enumerations> elements. 
 * 
 *  * The Excel and OpenOffice Calc files have the font in header records set to 'bold'.
 *    The work sheet names are set as appropriate to:
 *    - "Data Types"
 *    - "Enumerations"
 *    - "{name} Enumeration" truncated to the Excel work sheet name limit of 31 characters.
 *    - "Fields"
 *    where {name} is the enumeration name e.g. "ARRL_Section Enumeration".
 * 
 *    The files include some document and custom properties:
 *      Title:          "ADIF Specification Export of {item}, Version {version}, Status {status}"
 *      Author:         "ADIF Development Group"
 *      ADIF Version:   "{version}"
 *      ADIF Status:    "{status}"
 *    where
 *      {version}   is the ADIF Specification version, e.g. "3.0.6".
 *      {status}    is the ADIF Specification status of "Draft", "Proposed", or "Released".
 *      {item}      is "Data Types", "Enumerations", "Enumeration {name}", or "Fields".
 *      {name}      is an enumeration name, e.g. "ARRL_Section".
 *    E.g.
 *      Title:          "ADIF Specification Export of Enumeration ARRL_Section, Version 3.0.6, Status Proposed"
 *      Author:         "ADIF Development Group"
 *      ADIF Version:   "3.0.6"
 *      ADIF Status:    "Proposed"
 *  
 *  * Do not use the CSV and TSV files with spreadsheet software such as Excel and OpenOffice Calc.
 *    Instead use the Excel (.xlsx) or OpenOffice Calc (.ods) files.
 *    This is because spreadsheet software will damage CSV and TSV enumeration values that look like numbers and
 *    contain leading zeros and values that look superficially like dates and / or times.  E.g. 00123 in a CSV or TSV file
 *    will end up as 123 in a spreadsheet, whereas the Excel and OpenOffice Calc files are created with all
 *    cells set to be text, which stops the spreadsheet software guessing the data types of cells.
 * 
 *  * The tables in the ADIF Specification are exported with all their columns.  Any sequences of whitespaces
 *    are replaced by a single space.  Formatting is stripped out.
 * 
 *  * In CSV and TSV files:
 *    - The first record contains a header and this is followed by values records.
 *
 *  * In the enumeration CSV and TSV files:
 *    - The first column in values records contains the name of the enumeration, e.g. "ARRL_Section"
 *    - The enumerations.csv and enumerations.tsv files contain multiple header records, with one per enumeration.
 *      When reading the file, the header records can be identified by looking for the text:
 *          Eumeration Name
 *      within a record's first value.
 * 
 *  * There are some generated columns that don't exist in the ADIF Specification as such:
 *    - "Import-only":
 *          Values will be blank or contain the value "Import-only" if the specification indicates somewhere
 *          within a table row that the item the row refers to is import-only (deprecated).
 *    - "Comments":
 *          Sometimes the table cells with names in (e.g. data type names) contain additional information along the
 *          lines of "xxx (import-only; use yyy instead)".  In these cases the text within the parentheses is moved
 *          into the "Comments" field.
 *    - "ADIF Version" and "ADIF Status":
 *          All files except XML files include these columns, which contain the ADIF Specification
 *          version (e.g. 3.0.6) and status (Draft, Proposed, or Released).
 *    - "Minimum Value" and "Maximum Value":
 *          The data types and fields files include these columns.  They contain the minimum and maximum allowed
 *          numeric values for the data type or field.
 *          Note that this does not include all numeric fields because ADIF does not specify the minimum or maximum
 *          allowed values for number types as imposed by data types within programming languages.
 *    - "Header Field":
 *          The ADIF fields files contain this, which is "Y" for ADIF header fields and blank for ADIF record fields.
 *    - "DXCC Entity Code":
 *          All the Primary Administrative Subdivision tables in the specification are combined and this column is 
 *          exported to differentiate between them.
 *    - "Contained Within":
 *          Some of the Primary Administrative Subdivision tables tables include rows that span all columns in the
 *          table and contain the name/details of a locality that encloses the Primary Administrative Subdivisions
 *          defined within the following records.  These enclosing names/details are exported in this column.
 * 
 *  * "Deleted" columns will either be blank or contain the value "Deleted".
 * 
 *  * Future versions of the ADIF Specification may include changes in the structure of the tables such as:
 *    - a change in a column's title.
 *    - a change in the order of columns.
 *    - additional columns.
 *    - removal of columns.
 *    As far as possible, these types of changes will be avoided, but if they do occur, the files' contents will
 *    reflect them.
 *    To cater for this when accessing the files with software, it's strongly recommended that the titles in the
 *    header records are used to determine which column is which rather than relying on a column being the
 *    "nth" column between successive releases.
 *    Failing this, software that assumes that the "nth" column has a particular data item in should at least check
 *    that the header record contains the expected title for that column.
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using AdifReleaseLib;

namespace AdifExportFilesCreator
{
    /**
     * <summary>
     *   This class exports the ADIF Specification XHTML file as associated files for including as part of the release of a
     *   new version of the ADIF Specification.<br/>
     *   <br/>
     *   The files comprise CSV (.csv), TSV (.tsv), XML (.xml), Microsoft Excel (.xlsx), and OpenOffice Calc (.ods) files
     *   containing data types, enumerations, and fields.
     * </summary>
     */
    public class Specification
    {
        private const string
            Header_ADIF_Version = "ADIF Version",
            Header_ADIF_Status = "ADIF Status";

        /**
         * <summary>
         *   A delegate type for reporting progress back to calling code that can be logged or displalyed.
         * </summary>
         * 
         * <param name="message">The progress message.</param>
         */
        public delegate void ProgressReporter(string message);

        /**
         * <value>
         *   A <see cref="ProgressReporter"/> delegate object.
         * </value>
         */
        private ProgressReporter ReportProgress { get; set; }

        /**
         * <value>Represents a Yes or No answer from a user.</value>
         */
        public enum UserPrompterReply { Yes, No };

        /**
         * <summary>
         *   A delegate type for prompting a user with a message requiring a "Yes" or "No" answer.
         * </summary>
         * 
         * <param name="text">The text to be displayed.</param>
         * <param name="caption">A caption to be displayed.</param>
         * 
         * <value>The Yes or No answer from the user.</value>
         */
        public delegate UserPrompterReply UserPrompter(string text, string caption);

        /**
         * <value>
         *   A <see cref="UserPrompterReply"/> delegate object.
         * </value>
         */
        private UserPrompter PromptUser { get; set; }

        private readonly XmlDocument xmlDocIn;

        private readonly string
            ExportsPath,
            StartupPath;

        internal readonly string
            ExportsCsvPath,
            ExportsTsvPath,
            ExportsXlsxPath,
            ExportsOdsPath,
            ExportsXmlPath;

        private readonly SortedDictionary<string, Enumeration> enumerations;
        private readonly FieldList fields;
        private readonly DataTypeList dataTypes;

        internal string AdifVersion { get; }

        internal string AdifStatus { get; }

        internal string GetTitle(string name) => $"{name} exported from {AdifStatus} ADIF Specification {AdifVersion}";

        internal string Author => "ADIF Development Group";

        /**
         * <summary>
         *   An array of ADIF version numbers supported by this version of the program.
         * </summary>
         * 
         * <remarks>
         *   As code changes to this program may be required to support future versions of the
         *   ADIF Specification, only the most recently tested ADIF version is allowed along with
         *   any previous versions that have also been tested for backwards compatibility.
         * </remarks>
         * 
         * <value>
         *   The versions of ADIF supported by this version of the program.
         * </value>
         **/
        private readonly string[] SupportedVersions = new string[] {
            "3.1.4",
            "3.1.5",
        };

        private readonly string[] SupportedStatuses = new string[] {
            "Draft",
            "Proposed",
            "Released",
        };

        /**
         * <summary>
         *   Creates an object for exporting an ADIF Specification XHTML file as associated files for
         *   including as part of the release of a new version of the ADIF Specification.
         * </summary>
         * 
         * <param name="htmlDocumentPath">The path of the ADIF Specification XHTML file.</param>
         * <param name="startupPath">The path to the directory that contains the excutable and DLL files.</param>
         * <param name="reportProgress">A delegate that will be called back to report progress.</param>
         * <param name="promptUser">A delegate that will be called back to ask the user a question with a Yes or No answer.</param>
         */
        public Specification(
            string htmlDocumentPath,
            string startupPath,
            ProgressReporter reportProgress,
            UserPrompter promptUser)
        {
            ReportProgress = reportProgress;
            PromptUser = promptUser;

            ReportProgress?.Invoke($"Reading specification {htmlDocumentPath} ...");
#pragma warning disable format 
            enumerations        = new SortedDictionary<string, Enumeration>();
            fields              = new FieldList("fields", this);
            dataTypes           = new DataTypeList("dataTypes", this);
            ExportsPath         = Path.Combine(Path.GetDirectoryName(htmlDocumentPath), "exports");
            ExportsCsvPath      = Path.Combine(ExportsPath, "csv");
            ExportsTsvPath      = Path.Combine(ExportsPath, "tsv");
            ExportsXlsxPath     = Path.Combine(ExportsPath, "xlsx");
            ExportsOdsPath      = Path.Combine(ExportsPath, "ods");
            ExportsXmlPath      = Path.Combine(ExportsPath, "xml");
            StartupPath         = startupPath;
#pragma warning restore format
            xmlDocIn = LoadAdifSpecification(htmlDocumentPath);
            {
                {
                    // See if the ADIF Specification version is in the list of supported ADIF versions for this version of the program.
                    //
                    // The reason is that versions of the ADIF Specification that are more recent than this version of the program
                    // may include tables that require additional code to export correctly or possibly export at all.
                    //
                    // Additionally, if the ADIF Specification is earlier than the versions supported by this version of the program,
                    // the specification may not contain the exact data expected by the program.
                    // Also, this version of the progrma has only been tested with the versions specified in the list.

                    const string adifVersionMetaName = "adifversion";

                    XmlElement adifVersionMeta = 
                        (XmlElement)xmlDocIn.DocumentElement.SelectSingleNode($"head/meta[@name='{adifVersionMetaName}']") ?? throw new AdifException($"No meta tag found with the name \"{adifVersionMetaName}\"");

                    AdifVersion = adifVersionMeta.GetAttribute("content");

                    if (AdifVersion == null || (Array.IndexOf(SupportedVersions, AdifVersion) < 0))
                    {
                        StringBuilder temp = new StringBuilder(SupportedVersions.Length * 7);  // 5 characters + 2 more to allow for ", ".

                        foreach (string supportedVersion in SupportedVersions)
                        {
                            if (temp.Length > 0)
                            {
                                _ = temp.Append(", ");
                            }
                            _ = temp.Append(supportedVersion);
                        }

                        throw new AdifException(
                            $"ADIF Version {(AdifVersion ?? "null")} is not supported by this program.\r\n\r\nSupported versions are: {temp}");
                    }
                }
                {
                    const string adifStatusMetaName = "adifstatus";

                    XmlElement adifStatusMeta = 
                        (XmlElement)xmlDocIn.DocumentElement.SelectSingleNode($"head/meta[@name='{adifStatusMetaName}']") ?? throw new AdifException($"No meta tag found with the name \"{adifStatusMetaName}\"");

                    AdifStatus = adifStatusMeta.GetAttribute("content");

                    if (AdifStatus == null || !(Array.IndexOf(SupportedStatuses, AdifStatus) >= 0))
                    {
                        StringBuilder temp = new StringBuilder(128);

                        foreach (string supportedStatus in SupportedStatuses)
                        {
                            if (temp.Length > 0)
                            {
                                _ = temp.Append(", ");
                            }
                            _ = temp.Append(supportedStatus);
                        }

                        throw new AdifException(
                            $"ADIF Status is invalid: \"{(AdifStatus ?? "null")}\".\r\n\r\nSupported statuses are: {temp}");
                    }
                }

                ReportProgress?.Invoke($"Specification Version is {AdifVersion}, Status is {AdifStatus}");
            }
            {
                XmlNodeList deletions = xmlDocIn.SelectNodes("//*[@class='deletion'] | //*[@class='deletionstrike']");

                if (deletions.Count > 0)
                {
                    string fileName = Path.GetFileName(htmlDocumentPath);

                    if (fileName.ToLower().Contains("annotated.htm"))
                    {
                        if (PromptUser == null ||
                            PromptUser(
                                $"Warning: '{fileName}' is an annotated ADIF specification.\r\n" +
                                    "\r\n" +
                                    "The program will remove deleted text from the specification, but it is safer to use an un-annotated specification.\r\n" +
                                    "\r\n" +
                                    "Continue?",
                                "Warning: Annotated ADIF Specification") == UserPrompterReply.No)
                        {
                            throw new AdifAnnotatedSpecificationException($"The specification '{fileName}' is annotated");
                        }
                    }
                    foreach (XmlElement deletion in deletions)
                    {
                        deletion.ParentNode.RemoveChild(deletion);
                    }
                }
            }

            XmlNodeList tables = xmlDocIn.DocumentElement.GetElementsByTagName("table");

            foreach (XmlElement table in tables)
            {
                XmlAttribute idAttr = table.Attributes["id"];

                if (idAttr != null)
                {
                    string id = idAttr.Value;

                    if (id.StartsWith("Enumeration_"))
                    {
                        LoadEnumerationTable(id, table);
                    }
                    else if (id.StartsWith("Field_"))
                    {
                        fields.LoadTable(id, table);
                    }
                    else if (id == "_Data_Types")
                    {
                        dataTypes.LoadTable(id, table);
                    }
                }
            }
        }

        /**
         * <summary>
         *   Reads an enumeration table element from the ADIF Specification's HTML and loads the contents
         *   into Lists.
         * </summary>
         * 
         * <param name="id">The enumeration table's ID attribute.</param>
         * <param name="table">The enumeration table's XML element.</param>
         */
        internal void LoadEnumerationTable(string id, XmlElement table)
        {
            Enumeration enumeration;

            Enumeration.GetEnumerationName(id, out string enumerationName, out int dxccEntity);

            if (enumerations.TryGetValue(enumerationName, out Enumeration value))
            {
                enumeration = value;
            }
            else
            {
                enumeration = new Enumeration(enumerationName, this);
                enumerations.Add(enumerationName, enumeration);
            }
            enumeration.LoadTable(enumerationName, dxccEntity, table);
        }

        /**
         * <summary>
         *   Exports an ADIF Specification's data types, enumerations and fields as files.<br/>
         * </summary>
         */
        public void Export()
        {
            ReportProgress?.Invoke($"Exporting ...");

            try
            {
                CleanFiles();
                ExportSchema();
                ExportDataTypes();
                ExportEnumerations();
                ExportFields();
                ExportAll();

                ReportProgress?.Invoke("Completed");
            }
            finally
            {
                AdifReleaseLib.ExcelWriter.End();
            }
        }

        /**
         * <summary>
         *   Removes any files and sub-directories under "\export" leaving behind empty csv, ods, tsv, xlsx and xml sub-directories.
         * </summary>
         */
        private void CleanFiles()
        {
            if (!Directory.Exists(ExportsPath))
            {
                _ = Directory.CreateDirectory(ExportsPath);
            }
            else
            {
                foreach (string filePath in Directory.GetFiles(ExportsPath))
                {
                    File.Delete(filePath);  // Delete files in the exports directory.
                }

                // For use with Directory.Delete but not available in .NET Framework
                //EnumerationOptions enumOpts = new EnumerationOptions();

                //enumOpts.RecurseSubdirectories = true;
                //enumOpts.MaxRecursionDepth = 10;

                foreach (string fileTypeDirectoryPath in Directory.GetDirectories(ExportsPath))
                {
                    try
                    {
                        Directory.Delete(fileTypeDirectoryPath, true /*, enumOpts */);  // Delete the subdirectories and and their contents.
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        throw new AdifDeletionException("Failed to delete old directories and files", ex, fileTypeDirectoryPath);
                    }
                }
            }
            foreach (string fileTypeDirectoryName in new string[]
                {
                    ExportsCsvPath,
                    ExportsTsvPath,
                    ExportsXlsxPath,
                    ExportsOdsPath,
                    ExportsXmlPath
                })
            {
                _ = Directory.CreateDirectory(fileTypeDirectoryName);
            }
        }

        /**
         * <summary>
         *   Copies the adifexport.xsd file from the application's directory into the <![CDATA[<version>\exports\xml]]> directory.
         * </summary>
         **/
        private void ExportSchema()
        {
            const string schemaFileName = "adifexport.xsd";

            File.Copy(
                Path.Combine(StartupPath, schemaFileName),
                Path.Combine(ExportsXmlPath, schemaFileName));
        }

        /**
         * <summary>
         *   Exports the Data Taypes table from the ADIF Specification.
         * </summary>
         **/
        private void ExportDataTypes()
        {
            ReportProgress?.Invoke("Exporting data types ...");
            DataTypeList.BeginExport();
            dataTypes.Export();
            DataTypeList.EndExport();
        }

        /**
         * <summary>
         *   Exports the Enumerations tables from the ADIF Specification.
         * </summary>
         **/
        private void ExportEnumerations()
        {
            Enumeration.BeginExport();
            foreach (string enumerationName in enumerations.Keys)
            {
                ReportProgress?.Invoke(string.Format("Exporting enumeration {0} ...", enumerationName));
                enumerations[enumerationName].Export();
            }
            Enumeration.EndExport();
        }

        /**
         * <summary>
         *   Exports the Fields table from the ADIF Specification.
         * </summary>
         **/
        private void ExportFields()
        {
            ReportProgress?.Invoke("Exporting fields ...");
            FieldList.BeginExport();
            fields.Export();
            FieldList.EndExport();
        }

        /**
         * <summary>
         *   Merges the datatypes.xml, enumerations.xml and fields.xml files into the all.xml file.
         * </summary>
         */
        private void ExportAll()
        {
            string
                fieldsXml,
                enumerationsXml;

            XmlDocument
                rootAllDoc      = new XmlDocument(),
                doc             = new XmlDocument();

            ReportProgress?.Invoke("Exporting all ...");

            rootAllDoc.Load(Path.Combine(ExportsXmlPath, "datatypes.xml"));

            doc.Load(Path.Combine(ExportsXmlPath, "enumerations.xml"));
            enumerationsXml = doc.DocumentElement.InnerXml;

            doc.Load(Path.Combine(ExportsXmlPath, "fields.xml"));
            fieldsXml = doc.DocumentElement.InnerXml;

            rootAllDoc.DocumentElement.InnerXml = rootAllDoc.DocumentElement.InnerXml + enumerationsXml + fieldsXml;

            string fileName = Path.Combine(ExportsXmlPath, "all.xml");

            rootAllDoc.Save(fileName);
            AdifReleaseLib.Common.SetFileTimesToNow(fileName);
        }

        /**
         * <summary>
         *   Load the XHTML Windows-1252 encoded ADIF Specification file into an <see cref="XmlDocument"/> object.<br/>
         *   <br/>
         *   This is done by copying the file's contents into a temporary XML file,
         *   replacing the XMLDOC and html tags with the xml tag, a DTD, and an html tag with no attributes.&#160;
         *   The temporary file is then loaded into an <see cref="XmlDocument"/> object.
         * </summary>
         * 
         * <remarks>
         *   Previous attempts to load the XHTML file "as-is" into an <see cref="XmlDocument"/> object failed,
         *   hence the somewhat clumsy approach here of changing the XMLDOC declaration etc.
         * </remarks>
         * 
         * <param name="adifDocPath">The ADIF Specification XHTML file's path</param>
         * 
         * <returns>The <see cref="XmlDocument"/> object containing the XML version of the XHTML ADIF Specification.</returns>
         */
        private static XmlDocument LoadAdifSpecification(string adifDocPath)
        {            
            string tempDocPath = Path.GetTempFileName();
            XmlDocument xmlDocIn = new XmlDocument();

            try
            {
                xmlDocIn.PreserveWhitespace = true;

                using (StreamReader htmlDocInStream = new StreamReader(adifDocPath, Common.Windows1252Encoding))
                {
                    using (StreamWriter tempDocStream = new StreamWriter(tempDocPath, false, Common.Windows1252Encoding))
                    {
                        //while (!htmlDocInStream.ReadLine().Contains("<html"));      // Skip <html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">

                        //tempDocStream.WriteLine("<?xml version=\"1.0\" encoding=\"windows-1252\" ?>");
                        //tempDocStream.WriteLine("<html>");

                        //while (!htmlDocInStream.EndOfStream)
                        //{
                        //    tempDocStream.WriteLine(htmlDocInStream.ReadLine().Replace("&nbsp;", " "));  // The character entity &nbsp; is not declared by default in XML.                            
                        //}

                        // The following is a revised version of the above that uses an internal DTD to replace &nbsp; with
                        // a space instead of replacing them in the code here.  Both methods work, but this new one feels
                        // less of a "hack".

                        while (!htmlDocInStream.ReadLine().Contains("<html")) { };      // Skip <html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">

                        tempDocStream.WriteLine("<?xml version=\"1.0\" encoding=\"windows-1252\" ?>");
                        tempDocStream.WriteLine("<!DOCTYPE html [ <!ENTITY nbsp \" \"> ]>");
                        tempDocStream.WriteLine("<html>");

                        while (!htmlDocInStream.EndOfStream)
                        {
                            tempDocStream.WriteLine(htmlDocInStream.ReadLine());
                        }
                    }
                }

                // When using .NET 8, passing xmlDocIn.Load() a file name will fail because it cannot
                // deal with the Windows-1252 encoding.  To prevent this, it's necessary to set up a StreamReader with
                // a Windows-1252 encoding object and use that with xmlDocIn.Load

                using (StreamReader sw = new StreamReader(tempDocPath, Common.Windows1252Encoding))
                {
                    xmlDocIn.Load(sw);
                }
            }
            finally
            {
                try
                {
                    File.Delete(tempDocPath);
                }
                catch { }
            }
            return xmlDocIn;
        }

        /**
         * <summary>
         *   Writes out the header and fields to a <see cref="CsvWriter"/>, <see cref="TsvWriter"/> or
         *   <see cref="ExcelWriter"/> object.
         * </summary>
         * 
         * <param name="writer">The writer object to be used.</param>
         * <param name="headerRecord">A header record.</param>
         * <param name="valueRecords">An array of value records.</param>
         */
        internal void ExportToTableWriter(
            AdifReleaseLib.ITableWriter writer,
            List<string> headerRecord,
            List<string[]> valueRecords)
        {
            writer.WriteFields(headerRecord);
            writer.WriteField(Specification.Header_ADIF_Version);
            writer.WriteField(Specification.Header_ADIF_Status);
            writer.WriteHeader();

            foreach (string[] valueRecord in valueRecords)
            {
                writer.WriteFields(valueRecord);
                writer.WriteField(AdifVersion);
                writer.WriteField(AdifStatus);
                writer.WriteRecord();
            }
        }
    }
}
