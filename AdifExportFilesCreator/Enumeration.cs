using AdifReleaseLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace AdifExportFilesCreator
{
    /**
     * <summary>
     *   This class exports the enumerations in the ADIF Specification XHTML file as associated files for
     *   including in the released files.<br/>
     *   <br/>
     *   The files comprise CSV (.csv), TSV (.tsv), XML (.xml), Microsoft Excel (.xlsx), and OpenSource Calc (.ods) files.
     * </summary>
     */
    internal class Enumeration
    {
        private const int MaxFields = 20;
        private static readonly char[] commaSplitChar = new char[] { ',' };

#pragma warning disable format, IDE0055
        private const string
            PrimaryAdministrativeSubdivisionName        = "Primary_Administrative_Subdivision",
            SecondaryAdministrativeSubdivisionName      = "Secondary_Administrative_Subdivision",
            SecondaryAdministrativeSubdivisionAltName   = "Secondary_Administrative_Subdivision_Alt";

        private readonly List<string>               fieldNames      = new List<string>              (MaxFields);
        private readonly List<string[]>             fields          = new List<string[]>            (MaxFields);
        private readonly Dictionary<string, int>    fieldMap        = new Dictionary<string, int>   (MaxFields);
        private string[]                            htmlFieldNames  = null;
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
                    StringBuilder comments = new StringBuilder(256);
                    string[] values = new string[fieldNames.Count];
                    bool importOnly = false;
                    XmlNodeList nodes = row.GetElementsByTagName("td");

                    for (int cellIndex = 0; cellIndex < htmlFieldNames.Length; cellIndex++)
                    {
                        string value = cellIndex < nodes.Count ?
                            Common.ReplaceWhiteSpace(nodes[cellIndex].InnerText).Trim() :
                            string.Empty;

                        if (value.ToLower().Contains("import-only"))
                        {
                            importOnly = true;
                        }

                        if (cellIndex == 0 ||
                            cellIndex == arrlSectionSectionNameColumn ||
                            cellIndex == qslViaDescriptionColumn ||
                            cellIndex == secondaryAdministrativeSubdivisionColumn)
                        {
                            int bracketPosition = value.IndexOfAny(new char[] { '(', '{', '[' });

                            if (bracketPosition >= 0)
                            {
                                string newComment = value.Substring(bracketPosition + 1, value.Length - bracketPosition - 2);

                                if (!comments.ToString().ToLower().Contains(newComment.ToLower()))
                                {
                                    if (comments.Length > 0)
                                    {
                                        comments.Append("; ");
                                    }
                                    comments.Append(newComment);
                                }
                                value = value.Substring(0, bracketPosition).Trim();
                            }
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
                            StringBuilder newSubmodes = new StringBuilder(1024);

                            foreach (string submode in value.Split(
                                commaSplitChar,
                                StringSplitOptions.RemoveEmptyEntries))
                            {
                                newSubmodes.Append(submode.Trim()).Append(',');
                            }
                            value = newSubmodes.ToString().Substring(0, newSubmodes.Length-1);
                        }

                        if ((cellIndex == creditForColumn || cellIndex == arrlSectionDxccEntityCodeColumn) &&
                            value.IndexOf(' ') >= 0)
                        {
                            value = value.Replace(" ", string.Empty);
                        }

                        // In "Deleted" column, change "Y" to "Deleted"

                        if (cellIndex == deletedColumn)
                        {
                            if (value.Length >0)
                            {
                                switch (value.ToLower())
                                {
                                    //case "n":
                                    //    break;

                                    case "y":
                                    case "deleted":
                                        value = "Deleted";
                                        break;

                                    default:
                                        throw new AdifException($"Unexpected value for 'Deleted' in enumeration '{enumerationName}': value='{value}'");
                                }
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

                            int posn = value.IndexOf(" - ");

                            if (posn > 0)
                            {
                                if (value.Contains(" - for contacts made before"))
                                {
                                    primaryAdministrativeSubdivisionDeleted = "Deleted";
                                }
                                else if (value.Contains(" - for contacts made on or after"))
                                {
                                    // Okay
                                }
                                else if (value.Contains(" - referred to"))
                                {
                                    // Okay
                                }
                                else
                                {
                                    throw new Exception($"Unexpected comment string found in Primary Administrative Subdivision: value='{value}'");
                                }

                                if (comments.Length > 0)
                                {
                                    comments.Append("; ");
                                }
                                comments.Append(value.Substring(posn + 3));
                                value = value.Substring(0, posn);
                            }

                            posn = value.IndexOf(" (");
                            if (posn > 0)
                            {
                                if (comments.Length > 0)
                                {
                                    comments.Append("; ");
                                }
                                comments.Append(value.Substring(posn + 2, value.Length - posn - 3));
                                value = value.Substring(0, posn);
                            }                            
                        }
                        values[fieldMap[htmlFieldNames[cellIndex]]] = value;
                        System.Diagnostics.Debug.Assert(value != null);
                    }
                    values[enumerationNameColumn]   = enumerationName;
                    values[importOnlyColumn]        = importOnly ? "Import-only" : string.Empty;
                    values[commentsColumn]          = comments.ToString().ToLower() == "import-only" ? string.Empty : comments.ToString();

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
                            StringBuilder valuesList = new StringBuilder(values.Length);

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

        internal void Export()
        {
            OrderColumnsForExport(
                GetColumnNamesForFile(),
                out List<string> orderedHeaderRecord,
                out List<string[]> orderedValueRecords);

            string baseFileName = $"enumerations_{Name.ToLower()}";
#pragma warning disable format
            ExportToCsvTsvExcel (
                Name,
                baseFileName,
                orderedHeaderRecord,
                orderedValueRecords);
            ExportToXml         (
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
            using (AdifReleaseLib.CsvWriter csvWriter =
                new AdifReleaseLib.CsvWriter(Path.Combine(specification.ExportsCsvPath, $"{baseFileName}.csv")))
            {
                specification.ExportToTableWriter(csvWriter, headerRecord, valueRecords);
                specification.ExportToTableWriter(AllCsvWriter, headerRecord, valueRecords);
            }

            using (AdifReleaseLib.TsvWriter tsvWriter =
                new AdifReleaseLib.TsvWriter(Path.Combine(specification.ExportsTsvPath, $"{baseFileName}.tsv")))
            {
                specification.ExportToTableWriter(tsvWriter, headerRecord, valueRecords);
                specification.ExportToTableWriter(AllTsvWriter, headerRecord, valueRecords);
            }

            string content = enumerationName + " Enumeration";

            using (AdifReleaseLib.ExcelWriter excelWriter =
                new AdifReleaseLib.ExcelWriter(
                    Path.Combine(specification.ExportsXlsxPath, $"{baseFileName}.xlsx"),
                    Path.Combine(specification.ExportsOdsPath, $"{baseFileName}.ods"),
                    specification.AdifVersion,
                    specification.AdifStatus,
                    specification.GetTitle(content),
                    specification.Author,
                    content))
            {
                specification.ExportToTableWriter(excelWriter,    headerRecord, valueRecords);
                specification.ExportToTableWriter(AllExcelWriter, headerRecord, valueRecords);
            }
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
                            switch (title.ToUpper())
                            {
                                case "DELETED":
                                case "IMPORT-ONLY":
                                    valueEl.InnerText = XmlConvert.ToString(true);
                                    break;

                                case "DELETED DATE":
                                case "FROM DATE":
                                    valueEl.InnerText = AdifReleaseLib.Common.GetXmlDate(DateTime.Parse(value));
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
            AllEnumerationsEl.InnerXml += enumerationEl.OuterXml;

            string fileName = Path.Combine(specification.ExportsXmlPath, $"{baseFileName}.xml");

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
                    StringBuilder temp = new StringBuilder(1024);

                    foreach (string fieldName in fieldNames)
                    {
                        _ = temp.Append($"\"{fieldName}\", ");
                    }

                    throw new AdifException(
                        $"Table '{Name}': Unable to find field '{columnMapEntry}' in columnMap. fieldName={temp}");
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

            return string.Format(
                "{0},{1},{2},{3}",
                "Enumeration Name",
                defaultTitles,
                "Import-only",
                "Comments").Split(new char[] { ',' });
        }
#pragma warning disable format
        private static AdifReleaseLib.CsvWriter     AllCsvWriter        = null;
        private static AdifReleaseLib.TsvWriter     AllTsvWriter        = null;
        private static AdifReleaseLib.ExcelWriter   AllExcelWriter      = null;
        private static XmlDocument                  AllXmlDoc           = null;
        private static XmlElement                   AllEnumerationsEl   = null;
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
                    specification.Author,
                    content);

                AllXmlDoc = Common.CreateAdifExportXmlDocument(
                    specification.AdifVersion,
                    specification.AdifStatus,
                    specification.AdifDate,
                    out XmlElement adifEl);

                AllEnumerationsEl = (XmlElement)adifEl.AppendChild(AllXmlDoc.CreateElement("enumerations"));
            }
            catch
            {
#pragma warning disable format
                AllCsvWriter        = null;
                AllExcelWriter      = null;
                AllTsvWriter        = null;
                AllXmlDoc           = null;
                AllEnumerationsEl   = null;
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
            enumerationName = id.Substring("Enumeration_".Length);

            if (enumerationName.StartsWith(PrimaryAdministrativeSubdivisionName))
            {
                string dxccEntityValue = enumerationName.Substring(PrimaryAdministrativeSubdivisionName.Length + 1);

                dxccEntity = int.Parse(dxccEntityValue);
                enumerationName = PrimaryAdministrativeSubdivisionName;
            }
            else if (enumerationName.StartsWith(SecondaryAdministrativeSubdivisionAltName))
            {
                // Must check the _Alt name first because the non-_Alt check will match both cases.

                string dxccEntityValue = enumerationName.Substring(SecondaryAdministrativeSubdivisionAltName.Length + 1);

                dxccEntity = int.Parse(dxccEntityValue);
                enumerationName = SecondaryAdministrativeSubdivisionAltName;
            }
            else if (enumerationName.StartsWith(SecondaryAdministrativeSubdivisionName))
            {
                string dxccEntityValue = enumerationName.Substring(SecondaryAdministrativeSubdivisionName.Length + 1);

                dxccEntity = int.Parse(dxccEntityValue);
                enumerationName = SecondaryAdministrativeSubdivisionName;
            }
            else
            {
                dxccEntity = -1;
            }
        }
    }
}
