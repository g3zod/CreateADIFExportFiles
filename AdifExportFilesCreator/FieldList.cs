using AdifReleaseLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Text.RegularExpressions;

namespace AdifExportFilesCreator
{
    /**
     * <summary>
     *   xports the fields in the ADIF Specification XHTML file as associated files for
     *   including in the released files.<br/>
     *   <br/>
     *   The files comprise CSV (.csv), TSV (.tsv), XML (.xml), Microsoft Excel (.xlsx), and OpenSource Calc (.ods) files.
     * </summary>
     */
    internal class FieldList
    {
        private const int MaxFields = 200;
#pragma warning disable format
        private readonly List<string>               fieldNames      = new List<string>              (MaxFields);
        private readonly List<string[]>             fields          = new List<string[]>            (MaxFields);
        private readonly Dictionary<string, int>    fieldMap        = new Dictionary<string, int>   (MaxFields);
        private string[]                            htmlFieldNames  = null;
#pragma warning restore format
        private int AddField(string fieldName)
        {
            int fieldNumber;

            if (fieldMap.TryGetValue(fieldName, out int value))
            {
                fieldNumber = value;
            }
            else
            {
                fieldNames.Add(fieldName);
                fieldNumber = fieldNames.Count - 1;
                fieldMap.Add(fieldName, fieldNumber);
            }
            return fieldNumber;
        }

        private readonly string name;
        internal string Name => name;

        private FieldList()
        {
            // Not used.
        }

        /**
         * <summary>
         *   Initializes the object.
         * </summary>
         * 
         * <param name="name">The name of the type of ADIF data being exported.</param>
         * <param name="specification">The <see cref="Specification"/> object, used for obtaining information such as directory paths.</param>
         */
        private static Specification specification;
        internal FieldList(string name, Specification specification)
        {
            this.name = name;
            FieldList.specification  = specification;
        }

        internal void LoadTable(string id, XmlElement table)
        {
            bool isHeaderField = id == "Field_Header";
            bool readHeaderRow = false;
#pragma warning disable format
            int htmlFieldCount,

                fieldNameColumn,
                dataTypeColumn          = -1,
                enumerationColumn       = -1,
                descriptionColumn       = -1,

                importOnlyColumn        = -1,
                commentsColumn          = -1,

                headerFieldColumn       = -1,
                minimumValueColumn      = -1,
                maximumValueColumn      = -1;
#pragma warning restore format
            foreach (XmlElement row in table.GetElementsByTagName("tr"))
            {
                string
                    minimumValue = string.Empty,
                    maximumValue = string.Empty;

                if (!readHeaderRow)
                {
                    // Title row

                    XmlNodeList nodes = row.SelectNodes("th|td");

                    htmlFieldCount = nodes.Count;
                    htmlFieldNames = new string[htmlFieldCount];

                    if (htmlFieldCount == 0)
                    {
                        throw new AdifException("No columns found in header row");
                    }

                    for (int cellIndex = 0; cellIndex < htmlFieldCount; cellIndex++)
                    {
                        string fieldName = AdifReleaseLib.Common.ReplaceWhiteSpace(nodes[cellIndex].InnerText).Trim();

                        htmlFieldNames[cellIndex] = fieldName;
                        int fieldIndex = AddField(fieldName);

                        switch (fieldName)
                        {
                            case "Field Name":
                                fieldNameColumn = fieldIndex;
                                break;

                            case "Data Type":
                                dataTypeColumn = fieldIndex;
                                break;

                            case "Enumeration":
                                enumerationColumn = fieldIndex;
                                break;

                            case "Description":
                                descriptionColumn = fieldIndex;
                                break;
                        }
                    }

#pragma warning disable format
                    headerFieldColumn   = AddField("Header Field");
                    minimumValueColumn  = AddField("Minimum Value");
                    maximumValueColumn  = AddField("Maximum Value");

                    importOnlyColumn    = AddField("Import-only");
                    commentsColumn      = AddField("Comments");
#pragma warning restore format

                    readHeaderRow = true;
                }
                else
                {
                    StringBuilder comments = new StringBuilder(256);
                    string[] values = new string[fieldNames.Count];
                    bool importOnly = false;
                    XmlNodeList nodes = row.GetElementsByTagName("td");

                    for (int cellIndex = 0; cellIndex < htmlFieldNames.Length; cellIndex++)
                    {
                        string value = cellIndex < nodes.Count ?
                            AdifReleaseLib.Common.ReplaceWhiteSpace(nodes[cellIndex].InnerText).Trim() :
                            string.Empty;

                        if (cellIndex == descriptionColumn)
                        {
                            // The description may contain a minimum and/or maximum value surrounded by either
                            //     <span title="GreaterThan"></span> (only allowed for integers)
                            //     <span title="Minimum"></span>
                            //     <span title="Maximum"></span>
                            //
                            // The .//span is required because these can be embedded within other tags
                            // such as <a> and <span>.

                            XmlNode minMaxNode;

                            minMaxNode = nodes[cellIndex].SelectSingleNode(".//span[@title='GreaterThan']");
                            if (minMaxNode != null)
                            {
                                string sValue = minMaxNode.InnerText.Trim();

                                if (sValue.IndexOf('.') >= 0 ||
                                    !int.TryParse(sValue, out int iValue))
                                {
                                    throw new Exception(string.Format(
                                        "The <span title=\"GreaterThan\"> tag contains {0} but only integer values are allowed",
                                        sValue));
                                }
                                else
                                {
                                    minimumValue = (iValue + 1).ToString(AdifReleaseLib.Common.AdifNumberFormatInfo);
                                }
                            }

                            minMaxNode = nodes[cellIndex].SelectSingleNode(".//span[@title='Minimum']");
                            if (minMaxNode != null)
                            {
                                minimumValue = minMaxNode.InnerText.Trim();
                            }

                            minMaxNode = nodes[cellIndex].SelectSingleNode(".//span[@title='Maximum']");
                            if (minMaxNode != null)
                            {
                                maximumValue = minMaxNode.InnerText.Trim();
                            }

                            if (value.ToLower().StartsWith("import-only"))
                            {
                                importOnly = true;
                            }
                        }
                        else if (cellIndex == enumerationColumn &&
                                 value.Length > 0)
                        {
                            if (values[0] == "CREDIT_SUBMITTED" || values[0] == "CREDIT_GRANTED")
                            {
                                // These two fields have a complex enumeration column value due to a change of enumeration.
                                //
                                // To make this useful for machine use, output a comma-separated list of the data types,
                                // where the first item in any such data type list is always the current data type and
                                // subsequent data type items in the list are import-only (deprecated).

                                value = "Credit,Award";
                            }
                            else if (value[0] == '(')
                            {
                                value = ParseComplexEnumerationName(value);
                            }
                            else if (value.StartsWith("Submode"))
                            {
                                value = ParseComplexEnumerationName("(Submode, function of MODE field's value)");
                            }
                            else if (value.IndexOf(' ') >= 0)
                            {
                                value = ReplaceSpacesInEnumerationName(value);
                            }
                        }
                        else if (cellIndex == dataTypeColumn)
                        {
                            if (values[0] == "CREDIT_SUBMITTED" || values[0] == "CREDIT_GRANTED")
                            {
                                // These two fields have a complex data type column value due to a change of enumeration.
                                //
                                // To make this useful for machine use, output a comma-separated list of the data types,
                                // where the first item in any such data type list is always the current data type and
                                // subsequent data type items in the list are import-only (deprecated).

                                value = "CreditList,AwardList";
                            }
                        }
                        values[fieldMap[htmlFieldNames[cellIndex]]] = value;
                    }
#pragma warning disable format
                    values[importOnlyColumn]    = importOnly ? "Import-only" : string.Empty;
                    values[commentsColumn]      = comments.ToString().ToLower() == "import-only" ? string.Empty : comments.ToString();

                    values[headerFieldColumn]   = isHeaderField ? "Y" : string.Empty;

                    values[minimumValueColumn]  = minimumValue;
                    values[maximumValueColumn]  = maximumValue;
#pragma warning restore format

                    fields.Add(values);
                }
            }
        }

        internal void Export()
        {
            OrderColumnsForExport(
                GetColumnNamesForFile(),
                out List<string> orderedHeaderRecord,
                out List<string[]> orderedValueRecords);

            ExportToCsvTsvExcel(name, orderedHeaderRecord, orderedValueRecords);
            ExportToXml(name, orderedHeaderRecord, orderedValueRecords, specification.AdifVersion, specification.AdifStatus, true, true);
        }

        private static void ExportToCsvTsvExcel(
#pragma warning disable IDE0060
            string name,
#pragma warning restore IDE0060
            List<string> headerRecord, 
            List<string[]> valueRecords)
        {
            const string baseFileName = "fields";

            using (AdifReleaseLib.CsvWriter csvWriter =
                new AdifReleaseLib.CsvWriter(Path.Combine(specification.ExportsCsvPath, $"{baseFileName}.csv")))
            {
                specification.ExportToTableWriter(csvWriter, headerRecord, valueRecords);
            }

            using (AdifReleaseLib.TsvWriter tsvWriter =
                new AdifReleaseLib.TsvWriter(Path.Combine(specification.ExportsTsvPath, $"{baseFileName}.tsv")))
            {
                specification.ExportToTableWriter(tsvWriter, headerRecord, valueRecords);
            }

            string content = "Fields";

            using (ExcelWriter excelWriter =
                new ExcelWriter(
                    Path.Combine(specification.ExportsXlsxPath, $"{baseFileName}.xlsx"),
                    Path.Combine(specification.ExportsOdsPath, $"{baseFileName}.ods"),
                    specification.AdifVersion,
                    specification.AdifStatus,
                    specification.GetTitle(content),
                    specification.Author,
                    content))
            {
                specification.ExportToTableWriter(excelWriter, headerRecord, valueRecords);
            }
        }

        private static void ExportToXml(
#pragma warning disable IDE0060
            string name,
#pragma warning restore IDE0060
            List<string> headerRecord, 
            List<string[]> valueRecords, 
            string adifVersion, 
            string adifStatus, 
            bool addHeaderNamesToRecords, 
            bool addEmptyValues)
        {
            XmlDocument xmlDoc = Common.CreateAdifExportXmlDocument(adifVersion, adifStatus, out XmlElement adifEl);

            XmlElement fieldsEl = (XmlElement)adifEl.AppendChild(xmlDoc.CreateElement("fields"));
            XmlElement headerEl = (XmlElement)fieldsEl.AppendChild(xmlDoc.CreateElement("header"));

            foreach (string header in headerRecord)
            {
                XmlElement valueEl = (XmlElement)headerEl.AppendChild(xmlDoc.CreateElement("value"));

                valueEl.InnerText = header;
            }

            foreach (string[] valueRecord in valueRecords)
            {
                XmlElement valuesEl = (XmlElement)fieldsEl.AppendChild(xmlDoc.CreateElement("record"));
                int i = 0;

                foreach (string value in valueRecord)
                {
                    if ((!addEmptyValues) || value.Length > 0)
                    {
                        XmlElement valueEl = (XmlElement)valuesEl.AppendChild(xmlDoc.CreateElement("value"));

                        if (addHeaderNamesToRecords)
                        {
                            valueEl.SetAttribute("name", headerRecord[i]);
                        }
                        if (value.Length > 0)
                        {
                            switch (headerRecord[i])
                            {
                                case "Deleted":
                                case "Import-only":
                                case "Header Field":
                                    valueEl.InnerText = XmlConvert.ToString(true);
                                    break;

                                default:
                                    valueEl.InnerText = value;
                                    break;
                            }
                        }
                    }
                    i++;
                }
            }
            string fileName = Path.Combine(specification.ExportsXmlPath, "fields.xml");

            xmlDoc.Save(fileName);
            AdifReleaseLib.Common.SetFileTimesToNow(fileName);
        }

        private void OrderColumnsForExport(
            string[] columnNamesForFile,
            out List<string> newHeaderRecord,
            out List<string[]> newValueRecords)
        {
            const int MaxColumns = 20;

            int[] columnMapIndexes = new int[MaxColumns];
            int nextIndex = 0;

            for (int i = 0; i < MaxColumns; i++)
            {
                columnMapIndexes[i] = -1;
            }

            foreach (string columnMapEntry in columnNamesForFile)
            {
                bool found = false;

                for (int index = 0; index < fieldNames.Count; index++)
                {
                    if (fieldNames[index] == columnMapEntry)
                    {
                        columnMapIndexes[nextIndex++] = index;
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    throw new AdifException(
                        $"Table '{name}': Unable to find field '{columnMapEntry}' in columnMap");
                }
            }

            List<string> orderedHeaderRecord = new List<string>(20);

            for (int index = 0; index < columnMapIndexes.Length && columnMapIndexes[index] >= 0; index++)
            {
                orderedHeaderRecord.Add(fieldNames[columnMapIndexes[index]]);
            }

            List<string[]> orderedValueRecords = new List<string[]>(fields.Count);

            foreach (string[] valueRecord in fields)
            {
                string[] orderedValueRecord = new string[orderedHeaderRecord.Count];

                for (int index = 0; index < columnMapIndexes.Length && columnMapIndexes[index] >= 0; index++)
                {
                    orderedValueRecord[index] = valueRecord[columnMapIndexes[index]];
                }
                orderedValueRecords.Add(orderedValueRecord);
            }

            newHeaderRecord = orderedHeaderRecord;
            newValueRecords = orderedValueRecords;
        }

        private string[] GetColumnNamesForFile()
        {
            string defaultTitles = string.Join(",", htmlFieldNames);

            return string.Format(
                "{0},{1},{2},{3},{4},{5}",
                defaultTitles,
                "Header Field",
                "Minimum Value",
                "Maximum Value",
                "Import-only",
                "Comments").Split(new char[] { ',' });
        }

        /**
         * <summary>
         *   This method replaces the spaces in an enumeration name with underscores.
         * </summary>
         *
         * <param name="name">The enumeration's name.</param>
         * 
         * <returns>The enumeration name with spaces replaced by underscores.</returns>
         */
        private string ReplaceSpacesInEnumerationName(string name) => name.Replace(' ', '_');

        private static readonly Regex regex = new Regex(
                @"\(([^,]*), function of ([\w ]*) field's value\)");

        private string ParseComplexEnumerationName(string complexEnumerationName)
        {
            // Some enumerations form a set dependent on the value of another field and have the form:
            //      (x x x, function of y y y field's value)
            // E.g. (Primary Administrative Subdivision, function of MY_DXCC field's value)

            string enumerationName = string.Empty;
            Match match = regex.Match(complexEnumerationName);
            
            if (match.Success)
            {
                GroupCollection groups = match.Groups;

                //System.Diagnostics.Debug.WriteLine("-----------------------------");
                //foreach (System.Text.RegularExpressions.Group group in groups)
                //{
                //    System.Diagnostics.Debug.WriteLine(group.Value);
                //}

                if (groups.Count == 3)
                {
                    // [0] is entire expression, [1] is enumeration name, and [2] is the field name that the enumeration is a function of.

                    enumerationName = $"{ReplaceSpacesInEnumerationName(groups[1].Value)}[{groups[2].Value}]";
                }
            }
            return enumerationName;
        }

        /**
         * <summary>
         *   Performs any pre-export actions required.
         * </summary>
         */
        internal static void BeginExport()
        {
            // No pre-export actions required.
        }

        /**
         * <summary>
         *   Performs any post-export actions required.
         * </summary>
         */
        internal static void EndExport()
        {
            // No post-export actions required.
        }
    }
}
