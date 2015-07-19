using UnityEngine;
using System;
using System.IO;
using System.Data;
using System.Collections;
using System.Collections.Generic;
using Excel;
using Newtonsoft.Json;
using System.Text.RegularExpressions;

public class ExcelToJsonConverter 
{	
	public delegate void ConversionToJsonSuccessfullHandler();
	public event ConversionToJsonSuccessfullHandler ConversionToJsonSuccessfull = delegate {};
	
	public delegate void ConversionToJsonFailedHandler();
	public event ConversionToJsonFailedHandler ConversionToJsonFailed = delegate {};
	
	/// <summary>
	/// Converts all excel files in the input folder to json and saves them in the output folder.
	/// Each sheet within an excel file is saved to a separate json file with the same name as the sheet name.
	/// Files, sheets and columns whose name begin with '~' are ignored.
	/// </summary>
	/// <param name="inputPath">Input path.</param>
	/// <param name="outputPath">Output path.</param>
	/// <param name="recentlyModifiedOnly">If set to <c>true</c>, will only process recently modified files only.</param>
	public void ConvertExcelFilesToJson(string inputPath, string outputPath, bool recentlyModifiedOnly = false)
	{
		List<string> excelFiles = GetExcelFileNamesInDirectory(inputPath);
		Debug.Log("Excel To Json Converter: " + excelFiles.Count.ToString() + " excel files found.");
		
		if (recentlyModifiedOnly)
		{
			excelFiles = RemoveUnmodifiedFilesFromProcessList(excelFiles, outputPath);
			
			if (excelFiles.Count == 0)
			{
				Debug.Log("Excel To Json Converter: No updates to excel files since last conversion.");
			}
			else
			{
				Debug.Log("Excel To Json Converter: " + excelFiles.Count.ToString() + " excel files updated/added since last conversion.");
			}
		}
		
		bool succeeded = true;
		
		for (int i = 0 ; i < excelFiles.Count; i++)
		{
			if (!ConvertExcelFileToJson(excelFiles[i], outputPath))
			{
				succeeded = false;
				break;
			}
		}
		
		if (succeeded)
		{
			ConversionToJsonSuccessfull();
		}
		else
		{
			ConversionToJsonFailed();
		}
	}
	
	/// <summary>
	/// Gets all the file names in the specified directory
	/// </summary>
	/// <returns>The excel file names in directory.</returns>
	/// <param name="directory">Directory.</param>
	private List<string> GetExcelFileNamesInDirectory(string directory)
	{
		string[] directoryFiles = Directory.GetFiles(directory);
		List<string> excelFiles = new List<string>();
		
		// Regular expression to match against 2 excel file types (xls & xlsx), ignoring
		// files with extension .meta and starting with ~$ (temp file created by excel when fie
		Regex excelRegex = new Regex(@"^((?!(~\$)).*\.(xlsx|xls$))$");
		
		for (int i = 0; i < directoryFiles.Length; i++)
		{
			string fileName = directoryFiles[i].Substring(directoryFiles[i].LastIndexOf('/') + 1);
			
			if (excelRegex.IsMatch(fileName))
			{
				excelFiles.Add(directoryFiles[i]);
			}
		}
		
		return excelFiles;
	}
	
	/// <summary>
	/// Converts each sheet in the specified excel file to json and saves them in the output folder.
	/// The name of the processed json file will match the name of the excel sheet. Ignores
	/// sheets whose name begin with '~'. Also ignores columns whose names begin with '~'.
	/// </summary>
	/// <returns><c>true</c>, if excel file was successfully converted to json, <c>false</c> otherwise.</returns>
	/// <param name="filePath">File path.</param>
	/// <param name="outputPath">Output path.</param>
	public bool ConvertExcelFileToJson(string filePath, string outputPath)
	{
		Debug.Log("Excel To Json Converter: Processing: " + filePath);
		DataSet excelData = GetExcelDataSet(filePath);
		
		if (excelData == null)
		{
			Debug.LogError("Excel To Json Converter: Failed to process file: " + filePath);
			return false;
		}
		
		string spreadSheetJson = "";
		
		// Process Each SpreadSheet in the excel file
		for (int i = 0; i < excelData.Tables.Count; i++)
		{
			spreadSheetJson = GetSpreadSheetJson(excelData, excelData.Tables[i].TableName);
			if (String.IsNullOrEmpty(spreadSheetJson))
			{
				Debug.LogError("Excel To Json Converter: Failed to covert Spreadsheet '" + excelData.Tables[i].TableName + "' to json.");
				return false;
			}
			else
			{
				// The file name is the sheet name with spaces removed
				string fileName = excelData.Tables[i].TableName.Replace(" ", string.Empty);
				WriteTextToFile(spreadSheetJson, outputPath + "/" + fileName + ".json");
				Debug.Log("Excel To Json Converter: " + excelData.Tables[i].TableName + " successfully written to file.");
			}
		}
		
		return true;
	}
	
	/// <summary>
	/// Gets the excel data reader for the specified file.
	/// </summary>
	/// <returns>The excel data reader for file or null if file type is invalid.</returns>
	/// <param name="filePath">File path.</param>
	private IExcelDataReader GetExcelDataReaderForFile(string filePath)
	{
		FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
		
		// Create the excel data reader
		IExcelDataReader excelReader;
		
		// Create regular expressions to detect the type of excel file
		Regex xlsRegex = new Regex(@"^(.*\.(xls$))");
		Regex xlsxRegex = new Regex(@"^(.*\.(xlsx$))");
		
		// Read the excel file depending on it's type
		if (xlsRegex.IsMatch(filePath))
		{
			// Reading from a binary Excel file ('97-2003 format; *.xls)
			excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
		}
		else if (xlsxRegex.IsMatch(filePath))
		{
			// Reading from a OpenXml Excel file (2007 format; *.xlsx)
			excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
		}
		else
		{
			Debug.LogError("Excel To Json Converter: Unexpected files type: " + filePath);
			stream.Close();
			return null;
		}
		
		// Close the stream
		stream.Close();
		
		// First row are columns names
		excelReader.IsFirstRowAsColumnNames = true;
		
		return excelReader;
	}
	
	/// <summary>
	/// Gets the Excel data from the specified file
	/// </summary>
	/// <returns>The excel data set or null if file is invalid.</returns>
	/// <param name="filePath">File path.</param>
	private DataSet GetExcelDataSet(string filePath)
	{
		// Get the excel data reader with the excel data
		IExcelDataReader excelReader = GetExcelDataReaderForFile(filePath);
		
		if (excelReader == null)
		{
			return null;
		}
		
		// Get the data from the excel file
		DataSet data = new DataSet();
		
		do
		{
			// Get the DataTable from the current spreadsheet
			DataTable table = GetExcelSheetData(excelReader);
			
			if (table != null)
			{
				// Add the table to the data set
				data.Tables.Add(table);
			}
		}
		while(excelReader.NextResult()); // Read the next sheet
		
		return data;
	}
	
	/// <summary>
	/// Gets the Excel data from current spreadsheet
	/// </summary>
	/// <returns>The spreadsheet data table.</returns>
	/// <param name="excelReader">Excel Reader.</param>
	private DataTable GetExcelSheetData(IExcelDataReader excelReader)
	{
		if (excelReader == null)
		{
			Debug.LogError("Excel To Json Converter: Excel Reader is null. Cannot read data");
			return null;
		}
		
		// Ignore sheets which start with ~
		Regex sheetNameRegex = new Regex(@"^~.*$");
		if (sheetNameRegex.IsMatch(excelReader.Name))
		{
			return null;
		}
		
		// Create the table with the spreadsheet name
		DataTable table = new DataTable(excelReader.Name);
		table.Clear();
		
		string value = "";
		bool rowIsEmpty;
		
		// Read the rows and columns
		while (excelReader.Read())
		{
			DataRow row = table.NewRow();
			rowIsEmpty = true;
			
			for (int i = 0; i < excelReader.FieldCount; i++) 
			{
				// If the column is null and this is the first row, skip
				// to next iteration (do not want to include empty columns)
				if (excelReader.IsDBNull(i) && 
				    (excelReader.Depth == 1 || i > table.Columns.Count - 1))
				{
					continue;
				}
				
				value = excelReader.IsDBNull(i) ? "" : excelReader.GetString(i);
				
				// If this is the first row, add the values as columns
				if (excelReader.Depth == 1)
				{
					table.Columns.Add(value);
				}
				else // Otherwise, add as rows
				{
					row[table.Columns[i]] = value;
				}
				
				if (!string.IsNullOrEmpty(value))
				{
					rowIsEmpty = false;
				}
			}
			
			// Add the row to the table if it was not column headers and 
			// the row was not empty
			if (excelReader.Depth != 1 && !rowIsEmpty)
			{
				table.Rows.Add(row);
			}
		}
		
		return table;
	}
	
	/// <summary>
	/// Gets the json data for the specified spreadsheet in the specified DataSet
	/// </summary>
	/// <returns>The spread sheet json.</returns>
	/// <param name="excelDataSet">Excel data set.</param>
	/// <param name="sheetName">Sheet name.</param>
	private string GetSpreadSheetJson(DataSet excelDataSet, string sheetName)
	{
		// Get the specified table
		DataTable dataTable = excelDataSet.Tables[sheetName];
		
		// Remove empty columns
		for (int col = dataTable.Columns.Count - 1; col >= 0; col--)
		{
			bool removeColumn = true;
			foreach (DataRow row in dataTable.Rows)
			{
				if (!row.IsNull(col))
				{
					removeColumn = false;
					break;
				}
			}
			
			if (removeColumn)
			{
				dataTable.Columns.RemoveAt(col);
			}
		}
		
		// Remove columns which start with '~'
		Regex columnNameRegex = new Regex(@"^~.*$");
		for (int i = dataTable.Columns.Count - 1; i >= 0; i--)
		{
			if (columnNameRegex.IsMatch(dataTable.Columns[i].ColumnName))
			{
				dataTable.Columns.RemoveAt(i);
			}
		}
		
		// Serialze the data table to json string
		return Newtonsoft.Json.JsonConvert.SerializeObject(dataTable);
	}
	
	/// <summary>
	/// Writes the specified text to the specified file, overwriting it.
	/// Creates file if it does not exist.
	/// </summary>
	/// <param name="text">Text.</param>
	/// <param name="filePath">File path.</param>
	private void WriteTextToFile(string text, string filePath)
	{
		System.IO.File.WriteAllText(filePath, text);
	}
	
	/// <summary>
	/// Removes files which have not been modified since they were last processed
	/// from the process list
	/// </summary>
	/// <param name="excelFiles">Excel files.</param>
	private List<string> RemoveUnmodifiedFilesFromProcessList(List<string> excelFiles, string outputDirectory)
	{
		List<string> sheetNames;
		bool removeFile = true;
		
		// ignore sheets whose name starts with '~'
		Regex sheetNameRegex = new Regex(@"^~.*$");
		
		for (int i = excelFiles.Count - 1; i >= 0; i--)
		{
			sheetNames = GetSheetNamesInFile(excelFiles[i]);
			removeFile = true;
			
			for (int j = 0; j < sheetNames.Count; j++)
			{
				if (sheetNameRegex.IsMatch(sheetNames[j]))
				{
					continue;
				}
				
				string outputFile = outputDirectory + "/" + sheetNames[j] + ".json";
				if (!File.Exists(outputFile) ||
				    File.GetLastWriteTimeUtc(excelFiles[i]) > File.GetLastWriteTimeUtc(outputFile))
				{
					removeFile = false;
				}
			}
			
			if (removeFile)
			{
				excelFiles.RemoveAt(i);
			}
		}
		
		return excelFiles;
	}
	
	/// <summary>
	/// Gets the list of sheet names in the specified excel file
	/// </summary>
	/// <returns>The sheet names in file.</returns>
	/// <param name="filePath">File path.</param>
	private List<string> GetSheetNamesInFile(string filePath)
	{
		List<string> sheetNames = new List<string>();
		
		// Get the excel data reader with the excel data
		IExcelDataReader excelReader = GetExcelDataReaderForFile(filePath);
		
		if (excelReader == null)
		{
			return sheetNames;
		}
		
		do
		{
			// Add the sheet name to the list
			sheetNames.Add(excelReader.Name);
		}
		while(excelReader.NextResult()); // Read the next sheet
		
		return sheetNames;
	}
}
