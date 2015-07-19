# ExcelToJsonConverter
A Unity plugin to convert Excel files to Json

This simple plugin allows you to convert Excel files to Json within Unity.

How to use:

I. Copy the contents of the ExcelToJsonConverter folder into your project's Assets/Editor folder.
II. Open Unity project and select Tools -> Excel to Json Converter. 
    - Select input folder where Excel files are located.
    - Select output folder to save Json files.
    - Hit "Convert Excel Files" button.
    - Check console window for conversion info.
    
How to call from code:
Note: This is not designed for runtime use (although it should work in runtime, but will be slow).

ExcelToJsonConverter excelProcessor = new ExcelToJsonConverter();
excelProcessor.ConversionToJsonSuccessfull += ConversionToJsonSuccessfullCallback;
excelProcessor.ConversionToJsonFailed += ConversionToJsonFailedCallback;
excelProcessor.ConvertExcelFilesToJson(inputPath, outputPath, false);

This can be useful to integrate with your build scripts. See test project for example of this.
    
Notes:
- Supports .xls (1997 - 2004) and .xlsx (2007) excel file formats.
- Supports multiple sheets per file. Each sheet is saved separately to a Json file with the same name. (e.g. Sheet1 saved to Sheet1.json)
- Assumes that the first row of a sheet are column headers.
- If you want to ignore a column and not have it saved in the Json file, prefix the column header with '~'. E.G. ~Notes
- If you want to ignore a sheet and not have it converted to Json, prefix the sheet name with '~'. E.G. ~Temp


This plugin uses (and would not be possible without) ExcelDataReader & Json.Net.
