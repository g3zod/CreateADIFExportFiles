using AdifReleaseLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Text.Json;
using AdifExportFilesCreator.AdifExportObjects;

namespace AdifExportFilesCreator
{
    /**
     * <summary>
     *   Exports the data types in the ADIF Specification XHTML file as associated files for
     *   including in the ADIF released files.<br/>
     *   <br/>
     *   The files comprise CSV (.csv), TSV (.tsv), XML (.xml), Microsoft Excel (.xlsx), Apache OpenSource Calc (.ods), and JSON files.
     * </summary>
     */
    internal class DataTypeList
    {
        private const int MaxFields = 50;

        // Don't know why 'disable layout' is causing IDE0079 - elsewhere it doesn't - so suppress it.
#pragma warning disable IDE0079         // Remove unnecessary suppression
#pragma warning disable layout, IDE0055 // Layout, Fix formatting
        private readonly List<string>               fieldNames      = new(MaxFields);
        private readonly List<string[]>             fields          = new(MaxFields);
        private readonly Dictionary<string, int>    fieldMap        = new(MaxFields);
        private string[]                            htmlFieldNames  = null;
#pragma warning restore layout, IDE0055 // Layout, Fix formatting

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

        private DataTypeList()
        {
            // Not used.
        }

        private static Specification specification;

        /**
         * <summary>
         *   Initializes the object.
         * </summary>
         * 
         * <param name="name">The name of the type of ADIF data being exported.</param>
         * <param name="specification">The <see cref="Specification"/> object, used for obtaining information such as directory paths.</param>
         */
        internal DataTypeList(string name, Specification specification)
        {
            this.name = name;
            DataTypeList.specification = specification;
        }

        internal void LoadTable(
#pragma warning disable IDE0060
            string id,
#pragma warning restore IDE0060
            XmlElement table)
        {
            bool readHeaderRow = false;
#pragma warning disable format, IDE0055
            int htmlFieldCount,

                dataTypeNameColumn      = -1,
                dataTypeIndicatorColumn,
                descriptionColumn       = -1,

                importOnlyColumn        = -1,
                commentsColumn          = -1,

                minimumValueColumn      = -1,
                maximumValueColumn      = -1;
#pragma warning restore format, IDE0055
            foreach (XmlElement row in table.SelectNodes("tbody/tr"))
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
                        int bracketPosn = fieldName.IndexOf('(');

                        if (bracketPosn >= 0)
                        {
                            fieldName = fieldName[..(bracketPosn - 1)];
                        }

                        htmlFieldNames[cellIndex] = fieldName;
                        int fieldIndex = AddField(fieldName);

                        switch (fieldName)
                        {
                            case "Data Type Name":
                                dataTypeNameColumn = fieldIndex;
                                break;

                            case "Data Type Indicator":
                                dataTypeIndicatorColumn = fieldIndex;
                                break;

                            case "Description":
                                descriptionColumn = fieldIndex;
                                break;
                        }
                    }

                    minimumValueColumn = AddField("Minimum Value");
                    maximumValueColumn = AddField("Maximum Value");

                    importOnlyColumn = AddField("Import-only");
                    commentsColumn = AddField("Comments");

                    readHeaderRow = true;
                }
                else
                {
                    StringBuilder comments = new(256);
                    string[] values = new string[fieldNames.Count];
                    bool importOnly = false;
                    XmlNodeList nodes = row.GetElementsByTagName("td");

                    for (int cellIndex = 0; cellIndex < htmlFieldNames.Length; cellIndex++)
                    {
                        string value = cellIndex < nodes.Count ?
                            AdifReleaseLib.Common.ReplaceWhiteSpace(nodes[cellIndex].InnerText).Trim() :
                            string.Empty;

                        if (cellIndex == dataTypeNameColumn)
                        {
                            const string importOnlyText = " import-only";

                            if (value.Contains(importOnlyText, StringComparison.OrdinalIgnoreCase))
                            {
                                importOnly = true;
                                value = value.Replace(importOnlyText, string.Empty);
                            }
                        }
                        else if (cellIndex == descriptionColumn)
                        {
                            // The description may contain a minimum and/or maximum value surrounded by one of:
                            //     <span title="GreaterThan"></span> (only allowed for integers)
                            //     <span title="Minimum"></span>
                            //     <span title="Maximum"></span>
                            // The .//span is required because these can be embedded within other tags
                            // such as <a> and <span>.

                            XmlNode minMaxNode;

                            minMaxNode = nodes[cellIndex].SelectSingleNode(".//span[@title='GreaterThan']");
                            if (minMaxNode != null)
                            {
                                string sValue = minMaxNode.InnerText.Trim();

                                if (sValue.Contains('.') ||
                                    !int.TryParse(sValue, out int iValue))
                                {
                                    throw new Exception($"The <span title=\"GreaterThan\"> tag contains {sValue} but only integer values are allowed");
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
                        }
                        values[fieldMap[htmlFieldNames[cellIndex]]] = value;
                    }
                    values[importOnlyColumn] = importOnly ? "Import-only" : string.Empty;
                    values[commentsColumn] = comments.ToString().Equals("import-only", StringComparison.OrdinalIgnoreCase) ? string.Empty : comments.ToString();

                    values[minimumValueColumn] = minimumValue;
                    values[maximumValueColumn] = maximumValue;

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

            ExportToCsvTsvExcel(
                name,
                orderedHeaderRecord,
                orderedValueRecords);
            ExportToXml(
                name,
                orderedHeaderRecord,
                orderedValueRecords,
                specification.AdifVersion,
                specification.AdifStatus,
                specification.AdifDate,
                true,
                true);
            ExportToJson(
                name,
                orderedHeaderRecord,
                orderedValueRecords,
                specification.AdifVersion,
                specification.AdifStatus,
                specification.AdifDate,
                true,
                true);
        }

        private static void ExportToCsvTsvExcel(
#pragma warning disable IDE0060
            string enumerationName,
#pragma warning restore IDE0060
            List<string> headerRecord, 
            List<string[]> valueRecords)
        {
            const string baseFileName = "datatypes";

            using (AdifReleaseLib.CsvWriter csvWriter =
                new(Path.Combine(specification.ExportsCsvPath, $"{baseFileName}.csv")))
            {
                specification.ExportToTableWriter(csvWriter, headerRecord, valueRecords);
            }

            using (AdifReleaseLib.TsvWriter tsvWriter =
                new(Path.Combine(specification.ExportsTsvPath, $"{baseFileName}.tsv")))
            {
                specification.ExportToTableWriter(tsvWriter, headerRecord, valueRecords);
            }

            string content = "Data Types";

            using ExcelWriter excelWriter =
                new(
                    Path.Combine(specification.ExportsXlsxPath, $"{baseFileName}.xlsx"),
                    Path.Combine(specification.ExportsOdsPath, $"{baseFileName}.ods"),
                    specification.AdifVersion,
                    specification.AdifStatus,
                    specification.GetTitle(content),
                    Specification.Author,
                    content);
            specification.ExportToTableWriter(excelWriter, headerRecord, valueRecords);
        }

        private static void ExportToXml(
#pragma warning disable IDE0060
            string enumerationName,
#pragma warning restore IDE0060
            List<string> headerRecord, 
            List<string[]> valueRecords, 
            string adifVersion, 
            string adifStatus, 
            DateTime adifDate,
            bool addHeaderNamesToRecords, 
            bool addEmptyValues)
        {
            XmlDocument xmlDoc = Common.CreateAdifExportXmlDocument(adifVersion, adifStatus, adifDate, out XmlElement adifEl);

            XmlElement enumerationsEl = (XmlElement)adifEl.AppendChild(xmlDoc.CreateElement("dataTypes"));
            XmlElement headerEl = (XmlElement)enumerationsEl.AppendChild(xmlDoc.CreateElement("header"));

            foreach (string header in headerRecord)
            {
                XmlElement valueEl = (XmlElement)headerEl.AppendChild(xmlDoc.CreateElement("value"));

                valueEl.InnerText = header;
            }

            foreach (string[] valueRecord in valueRecords)
            {
                XmlElement valuesEl = (XmlElement)enumerationsEl.AppendChild(xmlDoc.CreateElement("record"));
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
                            if (headerRecord[i] == "Deleted" ||
                                headerRecord[i] == "Import-only")
                            {
                                valueEl.InnerText = XmlConvert.ToString(true);
                            }
                            else
                            {
                                valueEl.InnerText = value;
                            }
                        }
                    }
                    i++;
                }
            }
            string filePath = Path.Combine(specification.ExportsXmlPath, "datatypes.xml");

            xmlDoc.Save(filePath);
            AdifReleaseLib.Common.SetFileTimesToNow(filePath);
        }

        private static void ExportToJson(
#pragma warning disable IDE0060
            string resourceName,
#pragma warning restore IDE0060
            List<string> headerRecord,
            List<string[]> valueRecords,
            string adifVersion,
            string adifStatus,
            DateTime adifDate,
#pragma warning disable IDE0060 // Remove unused parameter
            bool addHeaderNamesToRecords,
#pragma warning restore IDE0060 // Remove unused parameter
            bool addEmptyValues)
        {
            Export export = Specification.CreateExportObject(
                adifVersion,
                adifStatus,
                adifDate);

            Adif adif = export.Adif;

            {
                DataTypes dataTypes = new()
                {
                    Header = [.. headerRecord]
                };
                adif.DataTypes = dataTypes;

                //if (Specification.ExportJsonRecordsAlt)
                //{
                //    dataTypes.RecordsAlt = [];

                //    foreach (string[] valueRecord in valueRecords)
                //    {
                //        int i = 0;

                //        Record record = [];

                //        foreach (string value in valueRecord)
                //        {
                //            if ((!addEmptyValues) || value.Length > 0)
                //            {
                //                switch (headerRecord[i].ToUpper())
                //                {
                //                    case "DELETED":
                //                    case "IMPORT-ONLY":
                //                        record.Add(headerRecord[i], Specification.JsonTrueAsString);
                //                        break;

                //                    default:
                //                        record.Add(headerRecord[i], value);
                //                        break;
                //                }
                //            }
                //            i++;
                //        }
                //        dataTypes.RecordsAlt.Add(record);
                //    }
                //}

                if (Specification.ExportJsonRecords)
                {
                    dataTypes.Records = [];

                    string dataTypeName = string.Empty;

                    foreach (string[] valueRecord in valueRecords)
                    {
                        int i = 0;

                        Record record = [];

                        foreach (string value in valueRecord)
                        {
                            if (i == 0)
                            {
                                dataTypeName = value;
                            }

                            if ((!addEmptyValues) || value.Length > 0)
                            {
                                switch (headerRecord[i].ToUpper())
                                {
                                    case "DELETED":
                                    case "IMPORT-ONLY":
                                        record.Add(headerRecord[i], Specification.JsonTrueAsString);
                                        break;

                                    default:
                                        if (Specification.IsRawExportDate(value))
                                        {
                                            // Dates should always be exported in the YYYY-DD-MMT00:00:00Z format.

                                            throw new AdifException($"Unexpected date while exporting JSON for '{headerRecord[i]}' in Data Type: value='{value}'");
                                        }
                                        record.Add(headerRecord[i], value);
                                        break;
                                }
                            }
                            i++;
                        }
                        dataTypes.Records.Add(dataTypeName, record);
                    }
                }
                specification.AllJsonExport.Adif.DataTypes = dataTypes;
            }

            string filePath = Path.Combine(specification.ExportsJsonPath, "datatypes.json");
            string json = JsonSerializer.Serialize(export, typeof(Export), Specification.JsonSerializerOptions);

            File.WriteAllText(filePath, json, Specification.JSONFileEncoding);
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

            List<string> orderedHeaderRecord = new(20);

            for (int index = 0; index < columnMapIndexes.Length && columnMapIndexes[index] >= 0; index++)
            {
                orderedHeaderRecord.Add(fieldNames[columnMapIndexes[index]]);
            }

            List<string[]> orderedValueRecords = new(fields.Count);

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
                "{0},{1},{2},{3},{4}",
                defaultTitles,
                "Minimum Value",
                "Maximum Value",
                "Import-only",
                "Comments").Split([',']);
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
