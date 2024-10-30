# Create ADIF Export Files
## Description
This Windows GUI application will take an [ADIF Specification](https://adif.org.uk) XHTML file and export its tables
in these formats:
- XML (.xml)
- CSV (.csv)
- TSV (.tsv)
- Microsoft Excel (.xlsx)
- OpenOffice Calc (.ods)

The generated files for each ADIF version are then provided on the website in a ZIP file (created outside this applicaiton).
 
## Projects
| Name  | Purpose |
| ----- | ------- |
| AdifExportFilesCreator | Exports tables from the ADIF Specification |
| AdifReleaseLib  | Contains classes to write different file types and implements a log file |
| CreateADIFExportFiles  | GUI and related code |

## Software Requirements
- Microsoft .NET
- Microsoft Excel (needed for generating .xlsx and .ods files)
- Microsoft Visual Studio 2022

## Limitations
- The application version is kept in step with the ADIF Specification versions.  This is because of the potential need to add code to support new features in the ADIF Specification.

## See Also
[Create ADIF Test Files](https://github.com/g3zod/CreateADIFExportFiles)
