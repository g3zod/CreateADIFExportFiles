/*
 * This class exports the ADIF Specification XHTML file as associated files for including as part of the release of a
 * new version of the ADIF Specification.
 * 
 * The files comprise CSV (.csv), TSV (.tsv), XML (.xml), Microsoft Excel (.xlsx), and Apache OpenOffice Calc (.ods) files containing
 * data types, enumerations, and fields.
 * 
 * Also JSON (.json) is now created but for the time being is not officially included in ADIF releases.
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
 *    exports/json                           - Directory for JSON (.json) files.
 *    exports/json/all.json                   - JSON file containing the data types, enumerations, and fields.
 *    exports/json/datatypes.json             - JSON file containing the data types.
 *    exports/json/enumerations.json          - JSON file containing the enumerations.
 *    exports/json/enumerations_{name}.json   - JSON file containing the enumeration with the name {name}.
 *    exports/json/fields.json                - JSON file containing the fields.
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
 *  * All the JSON files have the same overall structure as the all.json file, the differences being that:
 *    -  The datatypes.json file has a null value for the Enumerations and Fields objects. 
 *    -  The enumerations.json and named enumeration JSON files have null values for the DataTypes and Fields objects. 
 *    -  The fields.json file has a null value for the DataTypes and Enumerations objects. 
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

using AdifReleaseLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Text.Unicode;
using System.Xml;

namespace AdifExportFilesCreator
{
    /**
     * <summary>
     *   This class exports the ADIF Specification XHTML file as associated files for including as part of the release of a
     *   new version of the ADIF Specification.<br/>
     *   <br/>
     *   The files comprise CSV (.csv), TSV (.tsv), XML (.xml), Microsoft Excel (.xlsx), and Apache OpenOffice Calc (.ods) files
     *   containing data types, enumerations, and fields.
     * </summary>
     */
    public partial class Specification
    {
#pragma warning disable IDE0079
#pragma warning disable layout, IDE0055
        public const bool
            ExportJsonRecordsAlt =  true,
            ExportJsonRecords =     true;

        private const string

            // These are used when writing out header records to the exported files.

            Header_ADIF_Version =   "ADIF Version",
            Header_ADIF_Status =    "ADIF Status",

            // These are the names used in meta tags in the ADIF Specification XHTML file.

            adifVersionMetaName =   "adifversion",
            adifStatusMetaName =    "adifstatus",
            adifDateMetaName =      "adifdate";  // Introduced at the proposed version of ADIF 3.1.5 dated 2024/11/06.
#pragma warning restore layout, IDE0055, IDE0079

        /**
         * <summary>
         *   Creates a UTF-8 encoding object that does not emit a byte order mark (BOM) and throws an exception for invalid characters.
         *   
         *   This is for use with JSON because the JSON standard RFC 8259 specifies UTF-8 without a BOM:
         *   https://datatracker.ietf.org/doc/html/rfc8259
         * </summary>
         */
        internal static readonly Encoding JSONFileEncoding = new UTF8Encoding(false, true);

        /**
         * <summary>
         *   Returns a string containing a bool true value in JSON format.
         * </summary>
         */
        public static readonly string JsonTrueAsString = JsonSerializer.Serialize(true);

        /**
         * <summary>
         *   Converts a date from the ADIF Specification (e.g. "2024-11-13") to a UTC <see cref="DateTime"/>.
         *   <see cref=""/>
         * </summary>
         * 
         * <remarks>See <see href="https://stackoverflow.com/questions/3556144/how-to-create-a-net-datetime-from-iso-8601-format"/></remarks>
         */
        private static DateTime ConvertDateToUtc(string value) => DateTime.Parse(
            value.Contains('Z', StringComparison.OrdinalIgnoreCase) ? value : value + 'Z',
            null,
            System.Globalization.DateTimeStyles.RoundtripKind);

        /**
         * <summary>
         *   Converts a date <see cref="string"/> from the ADIF Specification (e.g. "2024-11-13") to a <see cref="string"/> containing
         *   a universal form of the date (e.g. "2024-11-13T00:00:00Z").
         * </summary>
         */
        public static string ConvertDateToJsonUtcString(string value) => JsonSerializer.Serialize(
            ConvertDateToUtc(value)).Replace(
                "\"",
                string.Empty);

        /**
         * <returns>
         *   Returns a <see cref="JsonSerializerOptions"/> object that reduces the need for escape sequences
         *   and provides formatted output.
         * </returns>
         */
        public static readonly JsonSerializerOptions JsonSerializerOptions = new()
        {
            // "UnicodeRanges.All" is misleading; perhaps because the encoder is aimed at JavaScript rather than
            // specifically JSON, it still encodes some characters that are legal in JSON string literals,
            // noticeably (so far):
            //
            //      26  &  ampersand
            //      27  '  apostrophe
            //      3C  <  less than sign
            //      3E  >  greater than sign

            Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
            WriteIndented = true
        };

        /**
         * <summary>
         *   An object that is used to serialize JSON for the all.json file.
         * </summary>
         */
        public AdifExportObjects.Export AllJsonExport;

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
            ExportsXmlPath,
            ExportsJsonPath;

        private readonly SortedDictionary<string, Enumeration> enumerations;
        private readonly FieldList fields;
        private readonly DataTypeList dataTypes;

        internal string AdifVersion { get; }

        internal string AdifStatus { get; }

        internal DateTime AdifDate { get; } = DateTime.MinValue;

        internal string GetTitle(string name) => $"{name} exported from {AdifStatus} ADIF Specification {AdifVersion}";

        internal static string Author => "ADIF Development Group";

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
        private readonly string[] SupportedVersions = [
            "3.1.4",
            "3.1.5",
            "3.1.6",
        ];

        private readonly string[] SupportedStatuses = [
            "Draft",
            "Proposed",
            "Released",
        ];

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
#if INCLUDE_TEST_CODE
            Extensions.TestExtensionMethods();
#endif            
            ReportProgress = reportProgress;
            PromptUser = promptUser;

            ReportProgress?.Invoke($"Reading specification {htmlDocumentPath} ...");
#pragma warning disable format
            enumerations        = [];
            fields              = new FieldList("fields", this);
            dataTypes           = new DataTypeList("dataTypes", this);
            ExportsPath         = Path.Combine(Path.GetDirectoryName(htmlDocumentPath), "exports");
            ExportsCsvPath      = Path.Combine(ExportsPath, "csv");
            ExportsTsvPath      = Path.Combine(ExportsPath, "tsv");
            ExportsXlsxPath     = Path.Combine(ExportsPath, "xlsx");
            ExportsOdsPath      = Path.Combine(ExportsPath, "ods");
            ExportsXmlPath      = Path.Combine(ExportsPath, "xml");
            ExportsJsonPath     = Path.Combine(ExportsPath, "json");
            StartupPath         = startupPath;
#pragma warning restore format
            //xmlDocIn = LoadAdifSpecification(htmlDocumentPath);
            xmlDocIn = XhtmlFileLoader.Load(htmlDocumentPath);
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

                    XmlElement adifVersionMeta =
                        (XmlElement)xmlDocIn.DocumentElement.SelectSingleNode($"head/meta[@name='{adifVersionMetaName}']") ?? throw new AdifException($"No meta tag found with the name \"{adifVersionMetaName}\"");

                    AdifVersion = adifVersionMeta.GetAttribute("content");

                    if (AdifVersion == null || (Array.IndexOf(SupportedVersions, AdifVersion) < 0))
                    {
                        StringBuilder temp = new(SupportedVersions.Length * 7);  // 5 characters + 2 more to allow for ", ".

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
                    XmlElement adifStatusMeta =
                        (XmlElement)xmlDocIn.DocumentElement.SelectSingleNode($"head/meta[@name='{adifStatusMetaName}']") ?? throw new AdifException($"No meta tag found with the name \"{adifStatusMetaName}\"");

                    AdifStatus = adifStatusMeta.GetAttribute("content");

                    if (AdifStatus == null || !(Array.IndexOf(SupportedStatuses, AdifStatus) >= 0))
                    {
                        StringBuilder temp = new(128);

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
                {
                    XmlElement adifDateMeta =
                        (XmlElement)xmlDocIn.DocumentElement.SelectSingleNode($"head/meta[@name='{adifDateMetaName}']");

                    string contents = adifDateMeta?.GetAttribute("content");

                    if (contents != null)
                    {
                        AdifDate = ConvertDateToUtc(contents);
                    }
                    else
                    {
                        Logger.Log("The ADIF Specification does not contain a valid meta tag named \"adifDate\"");
                    }
                }

                ReportProgress?.Invoke($"Specification Version is {AdifVersion}, Status is {AdifStatus}");
            }
            {
                XmlNodeList deletions = xmlDocIn.SelectNodes("//*[@class='deletion'] | //*[@class='deletionstrike']");

                if (deletions.Count > 0)
                {
                    string fileName = Path.GetFileName(htmlDocumentPath);

                    if (fileName.Contains("annotated.htm", StringComparison.OrdinalIgnoreCase))
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
         * 
         * <remarks>
         *   <para>
         *      The JSON files are created by populating the <see cref="AdifExportObjects"/> classes and then
         *      using C# <see cref="JsonSerializer"/> classes to create a JSON string.
         *    </para>
         *   <para>
         *      This not the only, or most direct, method of doing this.  Alternatives would be to use
         *      <see cref="Utf8JsonWriter"/> or <see cref="JsonNode"/>.  However, using the C# classes
         *      ensures that the they can be used to both serialize and de-serialize the JSON.
         *   </para>
         *   <para>
         *     Note: The <see cref="ValidateAllJson(string)"/> method has examples of using the <see cref="AdifExportObjects"/>,
         *     <see cref="JsonDocument"/>, and <see cref="JsonNode"/> classes to extra data from the JSON all.json file.
         *   </para>
         * </remarks>
         */
        public void Export()
        {
            ReportProgress?.Invoke($"Exporting ...");

            try
            {
                AllJsonExport = CreateExportObject(AdifVersion, AdifStatus, AdifDate);

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
                    ExportsXmlPath,
                    ExportsJsonPath,
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
                rootAllDoc = new(),
                doc = new();

            ReportProgress?.Invoke("Exporting all ...");

            rootAllDoc.Load(Path.Combine(ExportsXmlPath, "datatypes.xml"));

            doc.Load(Path.Combine(ExportsXmlPath, "enumerations.xml"));
            enumerationsXml = doc.DocumentElement.InnerXml;

            doc.Load(Path.Combine(ExportsXmlPath, "fields.xml"));
            fieldsXml = doc.DocumentElement.InnerXml;

            rootAllDoc.DocumentElement.InnerXml = rootAllDoc.DocumentElement.InnerXml + enumerationsXml + fieldsXml;

            {
                string allXmlFileName = Path.Combine(ExportsXmlPath, "all.xml");

                rootAllDoc.Save(allXmlFileName);
                AdifReleaseLib.Common.SetFileTimesToNow(allXmlFileName);
            }
            {
                string allJsonFileName = Path.Combine(ExportsJsonPath, "all.json");
                string json = JsonSerializer.Serialize(AllJsonExport, JsonSerializerOptions);

                File.WriteAllText(allJsonFileName, json, JSONFileEncoding);
                Common.SetFileTimesToNow(allJsonFileName);

                ValidateAllJson(allJsonFileName);
            }
        }
    
        private const string ExpectedResults_314 =
@"Data Type Name, Data Type Indicator, Description, Minimum Value, Maximum Value, Import-only, Comments
Enumeration Name, Abbreviation, Section Name, DXCC Entity Code, From Date, Deleted Date, Import-only, Comments
Field Name, Data Type, Enumeration, Description, Header Field, Minimum Value, Maximum Value, Import-only, Comments
Alaska
Bistrita-Nasaud
ADIF_VER
CREATED_TIMESTAMP
PROGRAMID
PROGRAMVERSION
USERDEFn
Boolean B
Number N
Date D
Time T
String S
IntlString I
MultilineString M
IntlMultilineString G
Enumeration E
Location L
";

        private const string ExpectedResults_315 =
@"Data Type Name, Data Type Indicator, Description, Minimum Value, Maximum Value, Import-only, Comments
Enumeration Name, Abbreviation, Section Name, DXCC Entity Code, From Date, Deleted Date, Import-only, Comments
Field Name, Data Type, Enumeration, Description, Header Field, Minimum Value, Maximum Value, Import-only, Comments
Alaska
Bistrița-Năsăud
ADIF_VER
CREATED_TIMESTAMP
PROGRAMID
PROGRAMVERSION
USERDEFn
Boolean B
Number N
Date D
Time T
String S
IntlString I
MultilineString M
IntlMultilineString G
Enumeration E
Location L
";

        private const string ExpectedResults_316 = ExpectedResults_315;

        /**
         * <summary>
         *   This performs some tests on the all.json file by de-serializing its contents into
         *   <see cref="AdifExportObjects"/>, <see cref="JsonDocument"/>, and <see cref="JsonNode"/> objects
         *   and extracting various data from them and saving it in a <see cref="StringBuilder"/> object.
         *   That is then compared again the expected strings for each version of ADIF.
         * </summary>
         * 
         * <remarks>
         *   It is almost certainly overkill using all three types of object, but the source code does also 
         *   serve as guide to extracting information from them (which is not always self-evident).
         * </remarks>
         * 
         * <param name="path">The path to the all.json file.</param>
         */
        private void ValidateAllJson(string path)
        {
            ReportProgress?.Invoke($"Validating all.json ...");

            string json = File.ReadAllText(path, Encoding.UTF8);
            StringBuilder results = new(32768);

            string expectedResults = AdifVersion switch
            {
                "3.1.4" => ExpectedResults_314,
                "3.1.5" => ExpectedResults_315,
                "3.1.6" => ExpectedResults_316,
                _ => string.Empty,
            };

            {
                ///////////////////////////////////////////
                // Validate using AdifExportObjects.Export.
                ///////////////////////////////////////////

                AdifExportObjects.Export export = JsonSerializer.Deserialize<AdifExportObjects.Export>(json);

                results.Clear();

                {
                    // Structure / Ensure Headers are correct in DataTypes, one Enumeration, and Fields.

                    // Func<AdifExportObjects.Header, string> HeaderToString = list =>
                    //    list == null ? "null" : string.Join(", ", list);

                    static string HeaderToString(AdifExportObjects.Header header) =>
                        header == null ? "Header == null" : string.Join(", ", header);

                    _ = results
                        .AppendLine(HeaderToString(export.Adif.DataTypes.Header))
                        .AppendLine(HeaderToString(export.Adif.Enumerations["ARRL_Section"].Header))
                        .AppendLine(HeaderToString(export.Adif.Fields.Header));
                }
                {
                    // Enumerations / ARRL_Sect

                    string arrlAkSectionName = export.Adif.Enumerations["ARRL_Section"].Records["AK"]["Section Name"];

                    _ = results.AppendLine(arrlAkSectionName);
                }
                {
                    // Primary Administrative Subdivision / Romania / BN
                    // The object names for Primary Administrative Subdivision are in the form code.dxcc, so the name is BN.275

                    string romaniaPasBn = export.Adif.Enumerations["Primary_Administrative_Subdivision"].Records["BN.275"]["Primary Administrative Subdivision"];

                    _ = results.AppendLine(romaniaPasBn);
                }
                {
                    // Header fields.

                    AdifExportObjects.Records records = export.Adif.Fields.Records;

                    foreach (string fieldName in records.Keys)
                    {
                        AdifExportObjects.Record value = records[fieldName];

                        // If it is not a header field, the item with the name (key) "Header Field" is omitted,
                        // so TryGetValue is needed rather than fetching the field and checking if its value is "Y".

                        if (value.TryGetValue("Header Field", out _))
                        {
                            _ = results.AppendLine(fieldName);
                        }
                    }
                }
                {
                    // Datatypes that have Data Type Indicators.

                    AdifExportObjects.DataTypes dataTypes = export.Adif.DataTypes;

                    foreach (string dataTypeName in dataTypes.Records.Keys)
                    {
                        AdifExportObjects.Record dataType = dataTypes.Records[dataTypeName];

                        // If there is no Data Type Indicator, the item with the name (key) "Data Type Indicator" is omitted,
                        // so TryGetValue is needed before fetching the value for the name (key) "Data Type Indicator".

                        if (dataType.TryGetValue("Data Type Indicator", out string dataTypeIndicator))
                        {
                            _ = results.AppendLine($"{dataTypeName} {dataTypeIndicator}");
                        }
                    }
                }

                string compareLinesResult = results.ToString().CompareLines(expectedResults);

                if (compareLinesResult.Length != 0)
                {
                    throw new AdifException($"Validation using AdifExportObjects.Export failed: {compareLinesResult}");
                }
            }

            {
                ///////////////////////////////
                // Validate using JsonDocument.
                ///////////////////////////////

                JsonDocument jsonDocument = JsonSerializer.Deserialize<JsonDocument>(json);

                results.Clear();

                {
                    // Structure / Ensure Headers are correct in DataTypes, one Enumeration, and Fields.

                    static string HeaderToString(JsonElement header)
                    {
                        StringBuilder listString = new(512);

                        foreach (JsonElement value in header.EnumerateArray())
                        {
                            if (listString.Length != 0)
                            {
                                _ = listString.Append(", "); 
                            }
                            _ = listString.Append(value.GetString());
                        }
                        return listString.ToString();
                    }

                    _ = results
                        .AppendLine(HeaderToString(jsonDocument.RootElement.GetProperty("Adif").GetProperty("DataTypes").GetProperty("Header")))
                        .AppendLine(HeaderToString(jsonDocument.RootElement.GetProperty("Adif").GetProperty("Enumerations").GetProperty("ARRL_Section").GetProperty("Header")))
                        .AppendLine(HeaderToString(jsonDocument.RootElement.GetProperty("Adif").GetProperty("Fields").GetProperty("Header")));
                }
                {
                    // Enumerations / ARRL_Sect

                    string arrlAkSectionName = jsonDocument.RootElement.GetProperty("Adif").GetProperty("Enumerations").GetProperty("ARRL_Section").GetProperty("Records").GetProperty("AK").GetProperty("Section Name").GetString();

                    _ = results.AppendLine(arrlAkSectionName);
                }
                {
                    // Primary Administrative Subdivision / Romania / BN
                    // The object names for Primary Administrative Subdivision are in the form code.dxcc, so the name is BN.275

                    string romaniaPasBn = jsonDocument.RootElement.GetProperty("Adif").GetProperty("Enumerations").GetProperty("Primary_Administrative_Subdivision").GetProperty("Records").GetProperty("BN.275").GetProperty("Primary Administrative Subdivision").GetString();

                    _ = results.AppendLine(romaniaPasBn);
                }
                {
                    // Header fields.

                    JsonElement records = jsonDocument.RootElement.GetProperty("Adif").GetProperty("Fields").GetProperty("Records");

                    foreach (JsonProperty field in records.EnumerateObject())
                    {
                        string fieldName = field.Name;

                        // If it is not a header field, the item with the name (key) "Header Field" is omitted,
                        // so TryGetProperty is needed rather than fetching the field and checking if its value is "Y".

                        if (field.Value.TryGetProperty("Header Field", out _))
                        {
                            _ = results.AppendLine(fieldName);
                        }
                    }
                }
                {
                    // Datatypes that have Data Type Indicators.

                    JsonElement records = jsonDocument.RootElement.GetProperty("Adif").GetProperty("DataTypes").GetProperty("Records");

                    foreach (JsonProperty dataType in records.EnumerateObject())
                    {
                        string dataTypeName = dataType.Name;

                        // If there is no Data Type Indicator, the item with the name (key) "Data Type Indicator" is omitted,
                        // so TryGetProperty is needed before fetching the value for the name (key) "Data Type Indicator".

                        if (dataType.Value.TryGetProperty("Data Type Indicator", out JsonElement dataTypeIndicator))
                        {
                            _ = results.AppendLine($"{dataTypeName} {dataTypeIndicator.GetString()}");
                        }
                    }
                }

                string compareLinesResult = results.ToString().CompareLines(expectedResults);

                if (compareLinesResult.Length != 0)
                {
                    throw new AdifException($"Validation using JsonDocument failed: {compareLinesResult}");
                }
            }

            {
                ///////////////////////////
                // Validate using JsonNode.
                ///////////////////////////

                JsonNode jsonNode = JsonSerializer.Deserialize<JsonNode>(json);

                results.Clear();

                {
                    // Structure / Ensure Headers are correct in DataTypes, one Enumeration, and Fields.

                    static string HeaderToString(JsonArray header)
                    {
                        StringBuilder listString = new(512);

                        foreach (string value in header.Select(v => (string)v))
                        {
                            if (listString.Length != 0)
                            {
                                _ = listString.Append(", ");
                            }
                            _ = listString.Append(value);
                        }
                        return listString.ToString();
                    }

                    _ = results
                        .AppendLine(HeaderToString(jsonNode.Root["Adif"]["DataTypes"]["Header"].AsArray()))
                        .AppendLine(HeaderToString(jsonNode.Root["Adif"]["Enumerations"]["ARRL_Section"]["Header"].AsArray()))
                        .AppendLine(HeaderToString(jsonNode.Root["Adif"]["Fields"]["Header"].AsArray()));
                }
                {
                    // Enumerations / ARRL_Sect

                    string arrlAkSectionName = jsonNode.Root["Adif"]["Enumerations"]["ARRL_Section"]["Records"]["AK"]["Section Name"].ToString();

                    _ = results.AppendLine(arrlAkSectionName);
                }
                {
                    // Primary Administrative Subdivision / Romania / BN
                    // The object names for Primary Administrative Subdivision are in the form code.dxcc, so the name is BN.275

                    string romaniaPasBn = jsonNode.Root["Adif"]["Enumerations"]["Primary_Administrative_Subdivision"]["Records"]["BN.275"]["Primary Administrative Subdivision"].ToString();

                    _ = results.AppendLine(romaniaPasBn);
                }
                {
                    // Header fields.

                    JsonObject records = (JsonObject)jsonNode.Root["Adif"]["Fields"]["Records"];

                    foreach (KeyValuePair<string, JsonNode> field in records)
                    {
                        string fieldName = field.Key;

                        // If it is not a header field, the item with the name (key) "Header Field" is omitted,
                        // so a null check is needed rather than fetching the field and checking if its value is "Y".

                        if (field.Value["Header Field"] != null)
                        {
                            _ = results.AppendLine(fieldName);
                        }
                    }
                }
                {
                    // Datatypes that have Data Type Indicators.

                    JsonObject records = (JsonObject)jsonNode.Root["Adif"]["DataTypes"]["Records"];

                    foreach (KeyValuePair<string, JsonNode> dataType in records)
                    {
                        string dataTypeName = dataType.Key;

                        // If there is no Data Type Indicator, the item with the name (key) "Data Type Indicator" is omitted,
                        // so a null check is needed before fetching the value for the name (key) "Data Type Indicator".

                        string dataTypeIndicator = (string)dataType.Value["Data Type Indicator"];

                        if (dataTypeIndicator != null)
                        {
                            _ = results.AppendLine($"{dataTypeName} {dataTypeIndicator}");
                        }
                    }
                }

                string compareLinesResult = results.ToString().CompareLines(expectedResults);

                if (compareLinesResult.Length != 0)
                {
                    throw new AdifException($"Validation using JsonNode failed: {compareLinesResult}");
                }
            }
        }

        private static readonly byte[] Utf8Bom = [0xEF, 0xBB, 0xBF];

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
        //private static XmlDocument LoadAdifSpecification(string adifDocPath)
        //{            
        //    string tempDocPath = Path.GetTempFileName();
        //    XmlDocument xmlDocIn = new();

        //    try
        //    {
        //        xmlDocIn.PreserveWhitespace = true;
        //        Encoding adifDocEncoding;

        //        // Determine whether the file is Windows-1252 (versions up to 3.1.5) or UTF-8 (versions 3.1.6 onwards).

        //        using (FileStream s = new(adifDocPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        //        {
        //            byte[] buffer = new byte[Utf8Bom.Length];

        //            int bytesRead = s.Read(buffer, 0, buffer.Length);

        //            if (bytesRead != Utf8Bom.Length)
        //            {
        //                throw new AdifException($"Failed to read first 3 bytes from {adifDocPath} to determine encoding");
        //            }
        //            adifDocEncoding = buffer.SequenceEqual(Utf8Bom) ?
        //                Encoding.UTF8 :
        //                Common.Windows1252Encoding;
        //        }

        //        using (StreamReader htmlDocInStream = new(adifDocPath, adifDocEncoding))
        //        {
        //            // Re-write the document as a true XML file and add an internal DTD to replace &nbsp; character
        //            // entities with a space.  (To be pedantic, it should replace them with &#160; but there is
        //            // no point because this program replaces all white-space sequences with a single space anyway.)

        //            using StreamWriter tempDocStream = new(tempDocPath, false, adifDocEncoding);

        //            while (!htmlDocInStream.ReadLine().Contains("<html")) { };      // Skip <html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">

        //            tempDocStream.WriteLine($"<?xml version=\"1.0\" encoding=\"{adifDocEncoding.WebName}\" ?>");
        //            tempDocStream.WriteLine("<!DOCTYPE html [ <!ENTITY nbsp \" \"> ]>");
        //            tempDocStream.WriteLine("<html>");

        //            while (!htmlDocInStream.EndOfStream)
        //            {
        //                tempDocStream.WriteLine(htmlDocInStream.ReadLine());
        //            }
        //        }

        //        // For ADIF 3.1.4 and 3.1.5, passing xmlDocIn.Load() a file name will fail because it cannot deal with
        //        // Windows-1252 encoding.  To overcome this, it is necessary to set up a StreamReader with a Windows-1252
        //        // encoding object and use that with xmlDocIn.Load()
        //        //
        //        // While that's not a problem for ADIF 3.1.6 and later because they will be UTF-8 encoded, the same code
        //        // using a StreamReader will of course work fine.

        //        using StreamReader sw = new(tempDocPath, adifDocEncoding);
        //        xmlDocIn.Load(sw);

        //        XmlElement meta = (XmlElement)xmlDocIn.DocumentElement.SelectSingleNode("head/meta[@http-equiv='Content-Type']") ??
        //            throw new AdifException($"meta http-equiv element not found in the ADIF Specification copy {tempDocPath}");
        //        string contentAttribute = meta.Attributes["content"].Value;

        //        // Ensure the encoding declared in the ADIF Specification matches whether or not the file has a UTF-8 byte order mark.
        //        // This could happen if the ADIF Specification declares itself as UTF-8 but has been saved without the byte order mark.

        //        if (contentAttribute.Contains("1252", StringComparison.OrdinalIgnoreCase))
        //        {
        //            if (adifDocEncoding != Common.Windows1252Encoding)
        //            {
        //                throw new AdifException($"The encoding in the ADIF Specification is \"{contentAttribute}\" but the file is UTF-8 encoded");
        //            }
        //        }
        //        else if (contentAttribute.Contains("utf-8", StringComparison.OrdinalIgnoreCase))
        //        {
        //            if (adifDocEncoding != Encoding.UTF8)
        //            {
        //                throw new AdifException($"The encoding in the ADIF Specification is \"{contentAttribute}\" but the file is not UTF-8 encoded");
        //            }
        //        }
        //        else
        //        {
        //            throw new AdifException($"The encoding in the ADIF Specification {contentAttribute} is not windows-1252 or utf-8");
        //        }
        //    }
        //    finally
        //    {
        //        try
        //        {
        //            File.Delete(tempDocPath);
        //        }
        //        catch (Exception ex)
        //        {
        //            Logger.Log($"Error: Failed to delete {tempDocPath}");
        //            Logger.Log(ex);
        //        }
        //    }
        //    return xmlDocIn;
        //}

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

        /**
         * <summary>
         *   Creates a new <see cref="AdifExportObjects.Export"/> object and initializes its <see cref="AdifExportObjects.Export.Adif"/> property
         *   to an Adif object, which in turn has its <see cref="AdifExportObjects.Adif.Version"/>, <see cref="AdifExportObjects.Adif.Status"/>
         *   and <see cref="AdifExportObjects.Adif.Date"/>properties initialized.
         * </summary>
         * 
         * <param name="adifVersion">The ADIF Specification's version in the form i.j.k</param>
         * <param name="adifStatus">The ADIF Specification's status of "Draft", "Proposed", or "Released".</param>
         * <param name="adifDate">The ADIF Specification's release date.</param>
         */
        public static AdifExportObjects.Export CreateExportObject(
            string adifVersion,
            string adifStatus,
            DateTime adifDate)
        {
            return new()
            {
                Adif = new()
                {
                    Status = adifStatus,
                    Version = adifVersion,
                    Date = adifDate,
                    Created = DateTime.UtcNow.Truncate(TimeSpan.TicksPerSecond)
                }
            };
        }

        [GeneratedRegex(@"^\d{4}-\d{2}-\d{2}$")]
        private static partial Regex RawExportDate();

        /**
         * <summary>
         *   When exporting JSON, this determines if a string has inadvertently not been converted from the form YYYY-MM-DD
         *   to YYYY-MM-DDT00:00:00Z
         * </summary>
         * 
         * <param name="text">The text to be tested.</param>
         * 
         * <returns>true if the date is in the form yyyy-mm-dd and otherwise, false</returns>
         */
        internal static bool IsRawExportDate(string text)
        {
            return RawExportDate().IsMatch(text);
        }
    }
}
