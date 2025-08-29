using AdifReleaseLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Text.Json;
using AdifExportFilesCreator.AdifExportObjects;
using System.Text.RegularExpressions;

namespace AdifExportFilesCreator
{
    /**
     * <summary>
     *   This class exports the enumerations in the ADIF Specification XHTML file as associated files for
     *   including in the released files.<br/>
     *   <br/>
     *   The files comprise CSV (.csv), TSV (.tsv), XML (.xml), Microsoft Excel (.xlsx), and Apache OpenSource Calc (.ods) files.
     * </summary>
     */
    internal partial class Enumeration
    {
        private const int MaxFields = 20;
        private static readonly char[] commaSplitChar = [','];

        private const string ImportOnlyLc = "import-only";

#pragma warning disable format, IDE0055
        private const string
            PrimaryAdministrativeSubdivisionName        = "Primary_Administrative_Subdivision",
            SecondaryAdministrativeSubdivisionName      = "Secondary_Administrative_Subdivision",
            SecondaryAdministrativeSubdivisionAltName   = "Secondary_Administrative_Subdivision_Alt",
            RegionName                                  = "Region";

        private readonly List<string>               fieldNames      = new(MaxFields);
        private readonly List<string[]>             fields          = new(MaxFields);
        private readonly Dictionary<string, int>    fieldMap        = new(MaxFields);
        private string[]                            htmlFieldNames;
#pragma warning restore format, IDE0055

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

        internal string Name { get; }

        private Enumeration()
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
        internal Enumeration(string name, Specification specification)
        {
            this.Name = name;
            Enumeration.specification  = specification;
        }

        internal void LoadTable(string enumerationName, int dxccEntity, XmlElement table)
        {
#pragma warning disable format, IDE0055
            bool readHeaderRow = false;

            int htmlFieldCount,

                importOnlyColumn                            = -1,
                commentsColumn                              = -1,
                enumerationNameColumn                       = -1,

                dxccEntityColumn                            = -1,
                containedWithinColumn                       = -1,
                oblastNumberColumn                          = -1,
                cqZoneColumn                                = -1,
                ituZoneColumn                               = -1,
                prefixColumn                                = -1,

                alaskaJudicialDistrictColumn                = -1,

                nzRegionsRegionColumn                       = -1,
                nzRegionsDistrictColumn                     = -1,

                arrlSectionSectionNameColumn                = -1,
                arrlSectionDxccEntityCodeColumn             = -1,
                lowerFreqMHzColumn                          = -1,
                upperFreqMHzColumn                          = -1,

                contestIDColumn                             = -1,

                submodesColumn                              = -1,

                creditForColumn                             = -1,

                qslViaDescriptionColumn                     = -1,

                deletedColumn                               = -1,

                primaryAdministrativeSubdivisionColumn      = -1,
                secondaryAdministrativeSubdivisionColumn    = -1,

                htmlDeletedColumn                           = -1;

            bool
                isPrimaryAdministrativeSubdivision      = dxccEntity >= 0 && Name.StartsWith("Primary"),
                isSecondaryAdministrativeSubdivision    = dxccEntity >= 0 && Name.StartsWith("Secondary") && !Name.EndsWith("_Alt"),
                isSecondaryAdministrativeSubdivisionAlt = dxccEntity >= 0 && Name.StartsWith("Secondary") && Name.EndsWith("_Alt");

            //if (dxccEntity >= 0)
            //{
            //    System.Diagnostics.Debug.Assert(isPrimaryAdministrativeSubdivision      == (name == PrimaryAdministrativeSubdivisionName));
            //    System.Diagnostics.Debug.Assert(isSecondaryAdministrativeSubdivision    == (name == SecondaryAdministrativeSubdivisionName));
            //    System.Diagnostics.Debug.Assert(isSecondaryAdministrativeSubdivisionAlt == (name == SecondaryAdministrativeSubdivisionAltName));
            //}

            string dxccEntityValue =
                isPrimaryAdministrativeSubdivision ||
                isSecondaryAdministrativeSubdivision ||
                isSecondaryAdministrativeSubdivisionAlt ?
                    dxccEntity.ToString() :
                    null;

            string containedWithinValue = null;
#pragma warning restore format, IDE0055

            foreach (XmlElement row in table.GetElementsByTagName("tr"))
            {
                string primaryAdministrativeSubdivisionDeleted = string.Empty;

                int colspan = -1;
                {
                    XmlElement th = (XmlElement)row.SelectSingleNode("th[@colspan]");

                    if (th != null)
                    {
                        colspan = int.Parse(th.GetAttribute("colspan"));
                    }
                }

                if (colspan > 0)
                {
                    if (readHeaderRow)
                    {
                        containedWithinValue = AdifReleaseLib.Common.ReplaceWhiteSpace(row.InnerText).Trim();
                    }
                    else
                    {
                        // This the row with the entity name and number in - ignore.

                        if (!row.InnerText.Contains(" " + dxccEntityValue + " "))
                        {
                            throw new AdifException($"Primary Administrative Subdivision DXCC entity code mismatch between the HTML file's table ID (\"{dxccEntityValue}\") and the table header row (\"{Common.ReplaceWhiteSpace(row.InnerText.Trim())}\")");
                        }
                    }
                }
                else if (!readHeaderRow)
                {
                    // Title row

                    XmlNodeList nodes = row.SelectNodes("th|td");

                    htmlFieldCount = nodes.Count;
                    htmlFieldNames = new string[htmlFieldCount];

                    if (htmlFieldCount == 0)
                    {
                        throw new AdifException($"No columns found in header row for enumeration '{enumerationName}'");
                    }

                    for (int cellIndex = 0; cellIndex < htmlFieldCount; cellIndex++)
                    {
                        string fieldName = AdifReleaseLib.Common.ReplaceWhiteSpace(nodes[cellIndex].InnerText).Trim();

                        if ((enumerationName == "Mode" || enumerationName == "Submode") &&
                            fieldName.Contains(" (to be supplied in a future specification)"))
                        {
                            fieldName = fieldName.Replace(" (to be supplied in a future specification)", string.Empty);
                        }
                        else if (enumerationName == "DXCC_Entity_Code" &&
                                 fieldName == "Country Code")
                        {
                            fieldName = "Entity Code";
                        }
                        else if (enumerationName == "QSO_Upload_Status" &&
                                 fieldName == "Via")
                        {
                            fieldName = "Status";
                        }
                        else if (enumerationName == "Award_Sponsor" &&
                                 fieldName == "Sponsor>")
                        {
                            // Workaround for error in ADIF 3.0.6

                            fieldName = "Sponsor";
                        }

                        htmlFieldNames[cellIndex] = fieldName;
                        int fieldIndex = AddField(fieldName);

                        if (isPrimaryAdministrativeSubdivision && fieldName == "Primary Administrative Subdivision")
                        {
                            primaryAdministrativeSubdivisionColumn = fieldIndex;  // For later checking of contents for comments.
                        }
                        else if (isSecondaryAdministrativeSubdivision && fieldName == "Secondary Administrative Subdivision")
                        {
                            secondaryAdministrativeSubdivisionColumn = fieldIndex;
                        }

                        if (enumerationName == "ARRL_Section")
                        {
                            if (fieldName == "Section Name")
                            {
                                arrlSectionSectionNameColumn = fieldIndex;
                            }
                            else if (fieldName == "DXCC Entity Code")
                            {
                                arrlSectionDxccEntityCodeColumn = fieldIndex;
                            }
                        }
                        else if (fieldName == "Lower Freq (MHz)")
                        {
                            lowerFreqMHzColumn = fieldIndex;
                        }
                        else if (fieldName == "Upper Freq (MHz)")
                        {
                            upperFreqMHzColumn = fieldIndex;
                        }
                        else if (fieldName == "Contest-ID")
                        {
                            contestIDColumn = fieldIndex;
                        }
                        else if (fieldName == "Submodes")
                        {
                            submodesColumn = fieldIndex;
                        }
                        else if (fieldName == "Deleted")
                        {
                            deletedColumn = fieldIndex;
                            htmlDeletedColumn = cellIndex;  // deletedColumn is the OUTPUT column and htmlDeletedColumn is the INPUT column.
                        }
                        else if (fieldName == "Credit For")
                        {
                            creditForColumn = fieldIndex;
                        }
                        else if (enumerationName == "QSL_Via" &&
                                 fieldName == "Description")
                        {
                            qslViaDescriptionColumn = fieldIndex;
                        }
                    }

                    importOnlyColumn        = AddField("Import-only");
                    commentsColumn          = AddField("Comments");
                    enumerationNameColumn   = AddField("Enumeration Name");

                    if (isPrimaryAdministrativeSubdivision)
                    {
                        dxccEntityValue         = dxccEntity.ToString();

                        dxccEntityColumn        = AddField("DXCC Entity Code");
                        containedWithinColumn   = AddField("Contained Within");
                        oblastNumberColumn      = AddField("Oblast #");
                        cqZoneColumn            = AddField("CQ Zone");
                        ituZoneColumn           = AddField("ITU Zone");
                        prefixColumn            = AddField("Prefix");
                        deletedColumn           = AddField("Deleted");
                    }
                    else if (isSecondaryAdministrativeSubdivision)
                    {
                        dxccEntityValue = dxccEntity.ToString();

                        dxccEntityColumn                = AddField("DXCC Entity Code");
                        alaskaJudicialDistrictColumn    = AddField("Alaska Judicial District");
                        deletedColumn                   = AddField("Deleted");
                    }
                    else if (isSecondaryAdministrativeSubdivisionAlt)
                    {
                        dxccEntityValue = dxccEntity.ToString();

                        dxccEntityColumn        = AddField("DXCC Entity Code");
                        nzRegionsRegionColumn   = AddField("Region");
                        nzRegionsDistrictColumn = AddField("District");
                        deletedColumn           = AddField("Deleted");
                    }

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
                            Common.ReplaceWhiteSpace(nodes[cellIndex].InnerText).Trim() :
                            string.Empty;

                        if (value.Contains(ImportOnlyLc, StringComparison.OrdinalIgnoreCase))
                        {
                            importOnly = true;
                        }

                        if (cellIndex == 0 ||
                            cellIndex == arrlSectionSectionNameColumn ||
                            cellIndex == qslViaDescriptionColumn ||
                            cellIndex == secondaryAdministrativeSubdivisionColumn)
                        {
                            value = ExtractValuesFromCell(value, out _, out _, out comments);
                        }

                        // Frequencies > 999 MHz in the Band enumeration contain thousands separators (','); remove these.

                        if ((cellIndex == lowerFreqMHzColumn || cellIndex == upperFreqMHzColumn) &&
                            value.IndexOf(',') > 0)
                        {
                            value = value.Replace(",", string.Empty);
                        }

                        // Some contest IDs are not in uppercase - make them uppercase.

                        if (cellIndex == contestIDColumn)
                        {
                            value = value.ToUpper();
                        }

                        // Remove spaces from the comma-delimited list in the Submodes field.

                        if (cellIndex == submodesColumn &&
                            value.IndexOf(',') > 0)
                        {
                            StringBuilder newSubmodes = new(1024);

                            foreach (string submode in value.Split(
                                commaSplitChar,
                                StringSplitOptions.RemoveEmptyEntries))
                            {
                                _ = newSubmodes.Append(submode.Trim()).Append(',');
                            }
                            value = newSubmodes.ToString()[..(newSubmodes.Length - 1)];
                        }

                        if ((cellIndex == creditForColumn || cellIndex == arrlSectionDxccEntityCodeColumn) &&
                            value.Contains(' '))
                        {
                            value = value.Replace(" ", string.Empty);
                        }

                        // In "Deleted" column, change "Y" to "Deleted"

                        if (cellIndex == deletedColumn)
                        {
                            if (value.Length > 0)
                            {
                                value = value.ToLower() switch
                                {
                                    "y" or "deleted" => "Deleted",
                                    _ => throw new AdifException($"Unexpected value for 'Deleted' in enumeration '{enumerationName}': value='{value}'"),
                                };
                            }
                        }

                        if (isPrimaryAdministrativeSubdivision &&
                            cellIndex == htmlDeletedColumn)
                        {
                            if (value == "Y")
                            {
                                primaryAdministrativeSubdivisionDeleted = "Deleted";
                            }
                        }

                        if (isPrimaryAdministrativeSubdivision &&
                            cellIndex == primaryAdministrativeSubdivisionColumn)
                        {
                            //if (value == "Kamchatka (Kamchatskaya oblast)" ||
                            //    value == "Krasnoyarsk (Krasnoyarsk Kraj)")
                            //{
                            //    primaryAdministrativeSubdivisionDeleted = "Deleted";
                            //}

                            value = ExtractValuesFromCell(value, out bool deleted, out importOnly, out comments);

                            if (deleted)
                            {
                                primaryAdministrativeSubdivisionDeleted = "Deleted";
                            }
                        }
                        values[fieldMap[htmlFieldNames[cellIndex]]] = value;
                        System.Diagnostics.Debug.Assert(value != null);
                    }

                    if (comments.ToString().Contains(ImportOnlyLc, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new AdifException($"Comments contains \"import-only\": \"{comments}\"");
                    }

                    values[enumerationNameColumn] = enumerationName;
                    values[importOnlyColumn] =      importOnly ? "Import-only" : string.Empty;
                    values[commentsColumn] =        comments.ToString();

                    if (isPrimaryAdministrativeSubdivision)
                    {
                        values[dxccEntityColumn] = dxccEntityValue;

                        values[containedWithinColumn] = containedWithinValue ?? string.Empty;

                        values[deletedColumn] = primaryAdministrativeSubdivisionDeleted;

                        if (values[oblastNumberColumn] == null)
                        {
                            values[oblastNumberColumn] = string.Empty;
                        }
                        if (values[cqZoneColumn] == null)
                        {
                            values[cqZoneColumn] = string.Empty;
                        }
                        if (values[ituZoneColumn] == null)
                        {
                            values[ituZoneColumn] = string.Empty;
                        }
                        if (values[prefixColumn] == null)
                        {
                            values[prefixColumn] = string.Empty;
                        }
                        if (values[deletedColumn] == null)
                        {
                            values[deletedColumn] = string.Empty;
                        }
                    }
                    else if (isSecondaryAdministrativeSubdivision)
                    {
                        values[dxccEntityColumn] = dxccEntityValue;

                        if (values[alaskaJudicialDistrictColumn] == null)
                        {
                            values[alaskaJudicialDistrictColumn] = string.Empty;
                        }
                        if (values[deletedColumn] == null)
                        {
                            values[deletedColumn] = string.Empty;
                        }
                    }
                    else if (isSecondaryAdministrativeSubdivisionAlt)
                    {
                        values[dxccEntityColumn] = dxccEntityValue;

                        if (values[nzRegionsRegionColumn] == null)
                        {
                            values[nzRegionsRegionColumn] = string.Empty;
                        }
                        if (values[nzRegionsDistrictColumn] == null)
                        {
                            values[nzRegionsDistrictColumn] = string.Empty;
                        }
                        if (values[deletedColumn] == null)
                        {
                            values[deletedColumn] = string.Empty;
                        }
                    }

                    foreach (string value in values)
                    {
                        if (value == null)
                        {
                            const string nullDisplay = "null";
                            StringBuilder valuesList = new(values.Length);

                            foreach (string temp in values)
                            {
                                _ = valuesList.Append($"\"{temp ?? nullDisplay}\", ");
                            }

                            throw new AdifException($"null value in enumeration '{enumerationName}': values={valuesList}");
                        }
                    }
                    fields.Add(values);
                }
            }
        }

        /**
         * <summary>
         *   This takes the value of a Primary Adminstrative Subdivision cell and separates out these items:<br />
         *   - The actual name part.<br />
         *   - Whether or not it is Deleted.<br />
         *   - Whether or not it is Import-only.<br />
         *   - Comments.<br />
         *   <br />
         *   If there is more than one comment, they will be separated with "; ".
         * </summary>
         * 
         * <remarks>
         *   The code is relatively complex due to the inconsistent ways the items are specified in the Primary Administrative Subdivision cells.
         * </remarks>
         * 
         * <param name="value">The value of the Primary Administrative Subdivision cell in the ADIF Specification.</param>
         * <param name="deleted">Whether or not it is Deleted.</param>
         * <param name="importOnly">Whether or not it is Import-only.</param>
         * <param name="comments">A list of comments.</param>
         * 
         * <returns>The actual name part.</returns>
         */
        private static string ExtractValuesFromCell(
            string value,
            out bool deleted,
            out bool importOnly,
            out StringBuilder comments)
        {
            bool adjustedZabaykalskyKrajValue = false;

            string
                temp,  // Preserve the value argument's value in case it is needed for an error message.
                newValue = string.Empty;

            // Adjust some values found in ADIF 3.1.5 and earlier.

            switch (value)
            {
                case "Distrito Federal (import-only replaced by CMX)":
                    value = "Distrito Federal (import-only - replaced by CMX)";
                    break;

                case "Sakha (Yakut) Republic":
                    value = "Republic of Sakha";
                    break;

                case "Zabaykalsky Kraj - referred to as Chita (Chitinskaya oblast) before 2008-03-01":
                    value = "Zabaykalsky Kraj - referred to as Chita [Chitinskaya oblast] before 2008-03-01";
                    adjustedZabaykalskyKrajValue = true;
                    break;

                case "Carbonia-Iglesias (import-only replaced by SU)":
                    value = "Carbonia-Iglesias (import-only - replaced by SU)";
                    break;

                default:
                    break;
            }

            temp = value;

            List<string> commentList = new(8);

            deleted = false;
            importOnly = false;

            while (true)
            {
                int
                    startPosn = temp.IndexOf('('),
                    endPosn;

                if (startPosn >= 0)
                {
                    endPosn = temp.IndexOf(')');

                    if (endPosn < 0)
                    {
                        throw new AdifException($"Enumeration value has mismatched parentheses: \"{value}\"");
                    }

                    {
                        string commentValue = temp.Substring(startPosn + 1, endPosn - startPosn - 1).Trim();

                        if (commentValue.StartsWith(ImportOnlyLc, StringComparison.OrdinalIgnoreCase))
                        {
                            importOnly = true;
                            commentValue = commentValue[ImportOnlyLc.Length..].Trim();

                            if (commentValue.StartsWith('-'))
                            {
                                commentValue = commentValue[1..].Trim();

                                if (!value.Contains("replaced by", StringComparison.OrdinalIgnoreCase))
                                {
                                    throw new AdifException($"Unexpected comment string found: value='{value}'");
                                }
                            }
                        }
                        if (commentValue.Length > 0)
                        {
                            commentList.Add(commentValue);
                        }
                    }
                    if (newValue.Length == 0)
                    {
                        newValue = temp[..startPosn].TrimEnd();
                    }
                    temp = temp[(endPosn + 1)..].Trim();
                }
                else
                {
                    break;
                }
            }

            if (newValue.Length == 0)
            {
                newValue = value;
            }

            {
                // This block deals with comments embedded in values that are not surrounded by parentheses.
                //
                // The aim is that future Specifications will enclose them in parentheses and this section.
                // However, the block will need to remain in order to allow use with earlier Specifications.

                int dashPosn = temp.IndexOf("- ");

                if (dashPosn >= 0)
                {
                    string[] parts = [
                        temp[..dashPosn].Trim(),
                        temp[(dashPosn + 1)..].Trim() ];

                    if (parts[1].Contains("for contacts made before", StringComparison.OrdinalIgnoreCase))
                    {
                        deleted = true;
                        if (parts[0].Length > 0)
                        {
                            newValue = parts[0];
                        }
                    }
                    else if (parts[1].Contains("for contacts made on or after", StringComparison.OrdinalIgnoreCase))
                    {
                        if (parts[0].Length > 0)
                        {
                            newValue = parts[0];
                        }
                    }
                    else if (parts[1].Contains("referred to", StringComparison.OrdinalIgnoreCase))
                    {
                        if (parts[0].Length > 0)
                        {
                            newValue = parts[0];
                        }
                    }
                    else
                    {
                        throw new AdifException($"Unexpected comment string found: value='{value}'");
                    }
                    commentList.Insert(0, parts[1]);
                }
            }

            {
                comments = new(256);

                for (int i = 0; i < commentList.Count; i++)
                {
                    if (string.IsNullOrEmpty(commentList[i]))
                    {
                        throw new AdifException($"Empty or null comment found for value: \"{value}\"");
                    }

                    if (i > 0)
                    {
                        _ = comments.Append("; ");
                    }
                    if (adjustedZabaykalskyKrajValue)
                    {
                        commentList[i] = commentList[i].Replace("[", "(").Replace("]", ")");
                    }
                    _ = comments.Append(commentList[i]);
                }
            }
            return newValue;
        }
        internal void Export()
        {
            OrderColumnsForExport(
                GetColumnNamesForFile(),
                out List<string> orderedHeaderRecord,
                out List<string[]> orderedValueRecords);

            string baseFileName = $"enumerations_{Name.ToLower()}";
#pragma warning disable format
            ExportToCsvTsvExcel(
                Name,
                baseFileName,
                orderedHeaderRecord,
                orderedValueRecords);
            ExportToXml(
                Name,
                baseFileName,
                orderedHeaderRecord,
                orderedValueRecords,
                specification.AdifVersion,
                specification.AdifStatus,
                specification.AdifDate,
                true,
                true);
            ExportToJson(
                Name,
                baseFileName,
                orderedHeaderRecord,
                orderedValueRecords,
                specification.AdifVersion,
                specification.AdifStatus,
                specification.AdifDate,
                true,
                true);
#pragma warning restore format
        }

        private static void ExportToCsvTsvExcel(
            string          enumerationName,
            string          baseFileName,
            List<string>    headerRecord, 
            List<string[]>  valueRecords)
        {
            using (CsvWriter csvWriter =
                new(Path.Combine(specification.ExportsCsvPath, $"{baseFileName}.csv")))
            {
                specification.ExportToTableWriter(csvWriter, headerRecord, valueRecords);
                specification.ExportToTableWriter(AllCsvWriter, headerRecord, valueRecords);
            }

            using (TsvWriter tsvWriter =
                new(Path.Combine(specification.ExportsTsvPath, $"{baseFileName}.tsv")))
            {
                specification.ExportToTableWriter(tsvWriter, headerRecord, valueRecords);
                specification.ExportToTableWriter(AllTsvWriter, headerRecord, valueRecords);
            }

            string content = enumerationName + " Enumeration";

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
            specification.ExportToTableWriter(AllExcelWriter, headerRecord, valueRecords);
        }

        private static void ExportToXml(
#pragma warning disable IDE0079
#pragma warning disable layout, IDE0055
            string          enumerationName,
            string          baseFileName,
            List<string>    headerRecord, 
            List<string[]>  valueRecords, 
            string          adifVersion, 
            string          adifStatus, 
            DateTime        adifDate,
            bool            addHeaderNamesToRecords, 
            bool            addEmptyValues)
#pragma warning restore IDE0079, layout, IDE0055
        {
            XmlDocument xmlDoc = Common.CreateAdifExportXmlDocument(adifVersion, adifStatus, adifDate, out XmlElement adifEl);

            XmlElement enumerationsEl = (XmlElement)adifEl.AppendChild(xmlDoc.CreateElement("enumerations"));
            XmlElement enumerationEl = (XmlElement)enumerationsEl.AppendChild(xmlDoc.CreateElement("enumeration"));
            XmlElement headerEl = (XmlElement)enumerationEl.AppendChild(xmlDoc.CreateElement("header"));

            enumerationEl.SetAttribute("name", enumerationName);

            foreach (string header in headerRecord)
            {
                XmlElement valueEl = (XmlElement)headerEl.AppendChild(xmlDoc.CreateElement("value"));

                valueEl.InnerText = header;
            }

            foreach (string[] valueRecord in valueRecords)
            {
                XmlElement valuesEl = (XmlElement)enumerationEl.AppendChild(xmlDoc.CreateElement("record"));
                int i = 0;

                foreach (string value in valueRecord)
                {
                    if ((!addEmptyValues) || value.Length > 0)
                    {
                        XmlElement valueEl = (XmlElement)valuesEl.AppendChild(xmlDoc.CreateElement("value"));
                        string title = headerRecord[i];

                        if (addHeaderNamesToRecords)
                        {
                            valueEl.SetAttribute("name", title);
                        }
                        if (value.Length > 0)
                        {
                            valueEl.InnerText = title.ToUpper() switch
                            {
                                "DELETED" or "IMPORT-ONLY" => XmlConvert.ToString(true),
                                "DELETED DATE" or "FROM DATE" => AdifReleaseLib.Common.GetXmlDate(DateTime.Parse(value)),
                                _ => value,
                            };
                        }
                    }
                    i++;
                }
            }
            AllEnumerationsEl.InnerXml += enumerationEl.OuterXml;

            string fileName = Path.Combine(specification.ExportsXmlPath, $"{baseFileName}.xml");

            xmlDoc.Save(fileName);
            AdifReleaseLib.Common.SetFileTimesToNow(fileName);
        }

        private static void ExportToJson(
#pragma warning disable IDE0079
#pragma warning disable layout, IDE0055
            string          enumerationName,
            string          baseFileName,
            List<string>    headerRecord, 
            List<string[]>  valueRecords, 
            string          adifVersion, 
            string          adifStatus, 
            DateTime        adifDate,
#pragma warning disable IDE0060 // Remove unused parameter
            bool            addHeaderNamesToRecords,
            bool            addEmptyValues)
#pragma warning restore IDE0060 // Remove unused parameter
#pragma warning restore IDE0079, layout, IDE0055
        {
            Export export = Specification.CreateExportObject(
                adifVersion,
                adifStatus,
                adifDate);

            Adif adif = export.Adif;

            {
                bool
                    isPrimaryAdministrativeSubdivision = enumerationName == PrimaryAdministrativeSubdivisionName,
                    isRegionEnumeration = enumerationName == RegionName;

                int dxccEntityCodeColumnIndex = -1,
                    deletedColumnIndex = -1;

                for (int i = 0; i < headerRecord.Count; i++)
                {
                    string headerName = headerRecord[i];

                    switch (headerName.ToUpper())
                    {
                        case "DXCC ENTITY CODE":
                            dxccEntityCodeColumnIndex = i;
                            break;

                        case "DELETED":
                            deletedColumnIndex = i;
                            break;

                        default:
                            break;
                    }
                }

                AdifExportObjects.Enumeration enumeration = new()
                {
                    Header = [.. headerRecord]
                };
                adif.Enumerations = [];

                //if (Specification.ExportJsonRecordsAlt)
                //{
                //    enumeration.RecordsAlt = [];

                //    foreach (string[] valueRecord in valueRecords)
                //    {
                //        int i = 0;

                //        Record record = [];

                //        foreach (string value in valueRecord)
                //        {
                //            if ((!addEmptyValues) || value.Length > 0)
                //            {
                //                string title = headerRecord[i];

                //                switch (title.ToUpper())
                //                {
                //                    case "DELETED":
                //                    case "IMPORT-ONLY":
                //                        record.Add(title, Specification.JsonTrueAsString);
                //                        break;

                //                    case "DELETED DATE":
                //                    case "FROM DATE":
                //                        record.Add(title, Specification.ConvertDateToJsonUtcString(value));
                //                        break;

                //                    default:
                //                        record.Add(title, value);
                //                        break;
                //                }
                //            }
                //            i++;
                //        }
                //        enumeration.RecordsAlt.Add(record);
                //    }
                //}
                
                if (Specification.ExportJsonRecords)
                {
                    enumeration.Records = [];

                    string abbreviation = string.Empty;

                    foreach (string[] valueRecord in valueRecords)
                    {
                        int i = 0;

                        Record record = [];

                        foreach (string value in valueRecord)
                        {
                            if (i == 1)
                            {
                                abbreviation = value;
                                if (isPrimaryAdministrativeSubdivision)
                                {
                                    // Primary Administrative Subdivision abbreviations are not unique, so the JSON
                                    // Property name is made up of the cod and DXCC entity number, e.g.
                                    //   MEX.50
                                    // where the code is MEX and the DXCC entity number is 50.
                                    //
                                    // Also, Property names are duplicated when a code has two records for a single
                                    // DXCC entity where one code is marked as "Deleted".
                                    //
                                    // To cater for this, the Property name has ".Deleted" appended, for example:
                                    //   "051.5.Deleted"
                                    //
                                    // There is a potential flaw in this if there is the same "Code" has more than one
                                    // "Deleted" record within a DXCC.  This has not happened so far.
                                    // Possibly such future "Deleted" records could be merged in the ADIF Specification
                                    // and there would be no issue.
                                    //
                                    // An alternative approach could be to include a serial number in the Property name,
                                    // e.g.
                                    //   "051.5.Deleted.0"
                                    //   "051.5.Deleted.1"
                                    //   ...
                                    //
                                    // However, the issue with that is allocating the serial number such that does not
                                    // change between releases of the ADIF Specification ... although that could be
                                    // avoided by being careful with the order of records in the ADIF Specification.
                                    //
                                    // While having more complex property names than code.dxcc makes them difficult to find
                                    // using indexing, it is still possible to find them by iterating through the array of
                                    // properties and inspecting the "Code" key/value pair, in which the "Code" value is
                                    // unadorned.  Here is an example - note that "Code" just contains "051":
                                    /*
                                          "051.5.Deleted": {
                                            "Enumeration Name": "Primary_Administrative_Subdivision",
                                            "Code": "051",
                                            "Primary Administrative Subdivision": "Märket",
                                            "DXCC Entity Code": "5",
                                            "Deleted": "true"
                                          },
                                     */

#pragma warning disable format
                                    string
                                        dxccEntityCode =    valueRecord[dxccEntityCodeColumnIndex],
                                        deleted =           valueRecord[deletedColumnIndex];
#pragma warning restore
                                    abbreviation = $"{abbreviation}.{dxccEntityCode}";
                                    if (deleted.Length > 0)
                                    {
                                        abbreviation = $"{abbreviation}.{deleted}.0";
                                    }
                                }
                                else if (isRegionEnumeration &&
                                         abbreviation != "NONE")
                                {
                                    // The Region table has three rows with Region Entity Code set to "KO",
                                    // so the key is made unique by appending a dot and the DXCC Entity Code,
                                    // e.g. KO.296

                                    abbreviation = $"{value}.{valueRecord[dxccEntityCodeColumnIndex]}";
                                }
                            }

                            if ((!addEmptyValues) || value.Length > 0)
                            {
                                string title = headerRecord[i];

                                switch (title.ToUpper())
                                {
                                    case "DELETED":
                                    case "IMPORT-ONLY":
                                        record.Add(title, Specification.JsonTrueAsString);
                                        break;

                                    case "DELETED DATE":
                                    case "FROM DATE":
                                        record.Add(title, Specification.ConvertDateToJsonUtcString(value));
                                        break;

                                    default:
                                        if (isRegionEnumeration &&
                                            (title.Equals("START DATE", StringComparison.OrdinalIgnoreCase) ||
                                             title.Equals("END DATE", StringComparison.OrdinalIgnoreCase)))
                                        {
                                            record.Add(title, Specification.ConvertDateToJsonUtcString(value));
                                        }
                                        else
                                        {
                                            if (IsRawExportDate(value))
                                            {
                                                // Dates should always be exported in the YYYY-DD-MMT00:00:00Z format.

                                                throw new AdifException($"Unexpected date while exporting JSON for '{title}' in enumeration '{enumerationName}': value='{value}'");
                                            }
                                            record.Add(title, value);
                                        }
                                        break;
                                }
                            }
                            i++;
                        }
                        try
                        {
                            enumeration.Records.Add(abbreviation, record);
                        }
                        catch (Exception ex)
                        {
                            throw new AdifException(
                                $"Exception adding {abbreviation} to {enumerationName}",
                                ex);
                        }
                    }
                }
                adif.Enumerations.Add(enumerationName, enumeration);
                AllEnumerationsJsonExport.Adif.Enumerations.Add(enumerationName, enumeration);
                specification.AllJsonExport.Adif.Enumerations.Add(enumerationName, enumeration);
            }

            string filePath = Path.Combine(specification.ExportsJsonPath, $"{baseFileName}.json");

            string json = JsonSerializer.Serialize(export, typeof(Export), Specification.JsonSerializerOptions);

            File.WriteAllText(filePath, json, Specification.JSONFileEncoding);

            AdifReleaseLib.Common.SetFileTimesToNow(filePath);
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
                    StringBuilder temp = new(1024);

                    foreach (string fieldName in fieldNames)
                    {
                        _ = temp.Append($"\"{fieldName}\", ");
                    }

                    throw new AdifException(
                        $"Table '{Name}': Unable to find field '{columnMapEntry}' in columnMap. fieldName={temp}");
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

        /**
         * <summary>
         *   Determines the list of column names to be output.<br />
         *   <br />
         *   Normally these are determined by the table column header values in the ADIF Specification,
         *   but for some tables, the column names need to be overridden,
         * </summary>
         */
        private string[] GetColumnNamesForFile()
        {
            string defaultTitles;

            if (Name.StartsWith(PrimaryAdministrativeSubdivisionName))
            {
                defaultTitles = "Code,Primary Administrative Subdivision,DXCC Entity Code,Contained Within,Oblast #,CQ Zone,ITU Zone,Prefix,Deleted";
            }
            else if (Name.StartsWith(SecondaryAdministrativeSubdivisionAltName))  // Must check this before the non-_Alt version to avoid mis-matching.
            {
                defaultTitles = "Code,DXCC Entity Code,Region,District,Deleted";
            }
            else if (Name.StartsWith(SecondaryAdministrativeSubdivisionName))
            {
                defaultTitles = "Code,Secondary Administrative Subdivision,DXCC Entity Code,Alaska Judicial District,Deleted";
            }
            else
            {
                defaultTitles = string.Join(",", htmlFieldNames);
            }

            return $"{"Enumeration Name"},{defaultTitles},{"Import-only"},{"Comments"}".Split([',']);
        }
#pragma warning disable format
        private static AdifReleaseLib.CsvWriter     AllCsvWriter =              null;
        private static AdifReleaseLib.TsvWriter     AllTsvWriter =              null;
        private static AdifReleaseLib.ExcelWriter   AllExcelWriter =            null;
        private static XmlDocument                  AllXmlDoc =                 null;
        private static XmlElement                   AllEnumerationsEl =         null;
        private static AdifExportObjects.Export     AllEnumerationsJsonExport = null;
#pragma warning restore format

        /**
         * <summary>
         *   Performs any pre-export actions required, which are
         *   creating writer objects for exporting files containing all enumerations and creating the XML document object to
         *   be saved later as all.xml
         * </summary>
         */
        internal static void BeginExport()
        {
            try
            {
                const string
                    content = "Enumerations",
                    baseFileName = "enumerations";

                AllCsvWriter = new AdifReleaseLib.CsvWriter(Path.Combine(specification.ExportsCsvPath, $"{baseFileName}.csv"));
                AllTsvWriter = new AdifReleaseLib.TsvWriter(Path.Combine(specification.ExportsTsvPath, $"{baseFileName}.tsv"));
                AllExcelWriter = new AdifReleaseLib.ExcelWriter(
                    Path.Combine(Path.Combine(specification.ExportsXlsxPath, $"{baseFileName}.xlsx")),
                    Path.Combine(Path.Combine(specification.ExportsOdsPath, $"{baseFileName}.ods")),
                    specification.AdifVersion,
                    specification.AdifStatus,
                    specification.GetTitle(content),
                    Specification.Author,
                    content);

                AllXmlDoc = Common.CreateAdifExportXmlDocument(
                    specification.AdifVersion,
                    specification.AdifStatus,
                    specification.AdifDate,
                    out XmlElement adifEl);

                AllEnumerationsEl = (XmlElement)adifEl.AppendChild(AllXmlDoc.CreateElement("enumerations"));

                AllEnumerationsJsonExport = Specification.CreateExportObject(
                    specification.AdifVersion,
                    specification.AdifStatus,
                    specification.AdifDate);

                AllEnumerationsJsonExport.Adif.Enumerations = [];

                specification.AllJsonExport.Adif.Enumerations = [];
            }
            catch
            {
#pragma warning disable format
                AllCsvWriter        = null;
                AllExcelWriter      = null;
                AllTsvWriter        = null;
                AllXmlDoc           = null;
                AllEnumerationsEl   = null;

                AllEnumerationsJsonExport = null;
#pragma warning disable format

                throw;
            }
        }

        /**
         * <summary>
         *   Performs any post-export actions required, which are closing the writer objects and
         *   setting the file dates to the current date and time.
         * </summary>
         */
        internal static void EndExport()
        {
            if (AllCsvWriter == null)
            {
                throw new AdifException("Enumeration.EndExport: Called without a successful prior call to BeginExport");
            }
            else
            {
                AllCsvWriter.Close();
                AllCsvWriter = null;

                AllExcelWriter.Close();
                AllExcelWriter = null;

                AllTsvWriter.Close();
                AllTsvWriter = null;

                string fileName = Path.Combine(specification.ExportsXmlPath, "enumerations.xml");

                AllXmlDoc.Save(fileName);
                AllXmlDoc = null;

                AdifReleaseLib.Common.SetFileTimesToNow(fileName);

                {
                    string allJsonFileName = Path.Combine(specification.ExportsJsonPath, "enumerations.json");
                    string json = JsonSerializer.Serialize(AllEnumerationsJsonExport, Specification.JsonSerializerOptions);

                    AllEnumerationsJsonExport = null;

                    File.WriteAllText(allJsonFileName, json, Specification.JSONFileEncoding);

                    AdifReleaseLib.Common.SetFileTimesToNow(allJsonFileName);
                }
            }
        }

        /**
         * <summary>
         *   Converts a table id attribute value in the ADIF Specification into
         *   an enumeration name and, where appropriate, a DXCC entity code. 
         *   The id attributes are one of two forms:
         *        Enumeration_{name}
         *        Enumeration_{name}_{dxcc}
         *   where {name} is the name of the field and {dxcc} is the DXCC entity code.
         * </summary>
         * 
         * <param name="id">The table element's id attribute.</param>
         * <param name="enumerationName">The name of the enumeration for including in the export files.</param>
         * <param name="dxccEntity">The DXCC entity of the enumeration for including in the export files.</param>
         */
        internal static void GetEnumerationName(string id, out string enumerationName, out int dxccEntity)
        {
            enumerationName = id["Enumeration_".Length..];

            if (enumerationName.StartsWith(PrimaryAdministrativeSubdivisionName))
            {
                string dxccEntityValue = enumerationName[(PrimaryAdministrativeSubdivisionName.Length + 1)..];

                dxccEntity = int.Parse(dxccEntityValue);
                enumerationName = PrimaryAdministrativeSubdivisionName;
            }
            else if (enumerationName.StartsWith(SecondaryAdministrativeSubdivisionAltName))
            {
                // Must check the _Alt name first because the non-_Alt check will match both cases.

                string dxccEntityValue = enumerationName[(SecondaryAdministrativeSubdivisionAltName.Length + 1)..];

                dxccEntity = int.Parse(dxccEntityValue);
                enumerationName = SecondaryAdministrativeSubdivisionAltName;
            }
            else if (enumerationName.StartsWith(SecondaryAdministrativeSubdivisionName))
            {
                string dxccEntityValue = enumerationName[(SecondaryAdministrativeSubdivisionName.Length + 1)..];

                dxccEntity = int.Parse(dxccEntityValue);
                enumerationName = SecondaryAdministrativeSubdivisionName;
            }
            else
            {
                dxccEntity = -1;
            }
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
         * <returns>true if the date is in the form YYYY-MM-DD</returns>
         */
        private static bool IsRawExportDate(string text)
        {
            return RawExportDate().IsMatch(text);
        }
    }
}
