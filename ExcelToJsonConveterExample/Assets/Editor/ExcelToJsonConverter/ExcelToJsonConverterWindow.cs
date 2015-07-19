using UnityEngine;
using UnityEditor;
using System.Collections;
using System.IO;

public class ExcelToJsonConverterWindow : EditorWindow 
{
	public static string kExcelToJsonConverterInputPathPrefsName = "ExcelToJson.InputPath";
	public static string kExcelToJsonConverterOuputPathPrefsName = "ExcelToJson.OutputPath";
	public static string kExcelToJsonConverterModifiedFilesOnlyPrefsName = "ExcelToJson.OnlyModifiedFiles";
	
	private string _inputPath;
	private string _outputPath;
	private bool _onlyModifiedFiles;

	private ExcelToJsonConverter _excelProcessor;

	[MenuItem ("Tools/Excel To Json Converter")]
	public static void ShowWindow() 
	{
		EditorWindow.GetWindow(typeof(ExcelToJsonConverterWindow), true, "Excel To Json Converter", true);
	}

	public void OnEnable()
	{
		if (_excelProcessor == null)
		{
			_excelProcessor = new ExcelToJsonConverter();
		}

		_inputPath = EditorPrefs.GetString(kExcelToJsonConverterInputPathPrefsName, Application.dataPath);
		_outputPath = EditorPrefs.GetString(kExcelToJsonConverterOuputPathPrefsName, Application.dataPath);
		_onlyModifiedFiles = EditorPrefs.GetBool(kExcelToJsonConverterModifiedFilesOnlyPrefsName, false);
	}
	
	public void OnDisable()
	{
		EditorPrefs.SetString(kExcelToJsonConverterInputPathPrefsName, _inputPath);
		EditorPrefs.SetString(kExcelToJsonConverterOuputPathPrefsName, _outputPath);
		EditorPrefs.SetBool(kExcelToJsonConverterModifiedFilesOnlyPrefsName, _onlyModifiedFiles);
	}

	void OnGUI()
	{
		GUILayout.BeginHorizontal();

		GUIContent inputFolderContent = new GUIContent("Input Folder", "Select the folder where the excel files to be processed are located.");
		EditorGUIUtility.labelWidth = 120.0f;
		EditorGUILayout.TextField(inputFolderContent, _inputPath, GUILayout.MinWidth(120), GUILayout.MaxWidth(500));
		if (GUILayout.Button(new GUIContent("Select Folder"), GUILayout.MinWidth(80), GUILayout.MaxWidth(100)))
		{
			_inputPath = EditorUtility.OpenFolderPanel("Select Folder with Excel Files", _inputPath, Application.dataPath);
		}

		GUILayout.EndHorizontal();

		GUILayout.BeginHorizontal();

		GUIContent outputFolderContent = new GUIContent("Output Folder", "Select the folder where the converted json files should be saved.");
		EditorGUILayout.TextField(outputFolderContent, _outputPath, GUILayout.MinWidth(120), GUILayout.MaxWidth(500));
		if (GUILayout.Button(new GUIContent("Select Folder"), GUILayout.MinWidth(80), GUILayout.MaxWidth(100)))
		{
			_outputPath = EditorUtility.OpenFolderPanel("Select Folder to save json files", _outputPath, Application.dataPath);
		}
		
		GUILayout.EndHorizontal();

		GUIContent modifiedToggleContent = new GUIContent("Modified Files Only", "If checked, only excel files which have been newly added or updated since the last conversion will be processed.");
		_onlyModifiedFiles = EditorGUILayout.Toggle(modifiedToggleContent, _onlyModifiedFiles);

		if (string.IsNullOrEmpty(_inputPath) || string.IsNullOrEmpty(_outputPath))
		{
			GUI.enabled = false;
		}

		GUILayout.BeginArea(new Rect((Screen.width / 2) - (200 / 2), (Screen.height / 2) - (25 / 2), 200, 25));

		if (GUILayout.Button("Convert Excel Files"))
		{
			_excelProcessor.ConvertExcelFilesToJson(_inputPath, _outputPath, _onlyModifiedFiles);
		}

		GUILayout.EndArea();

		GUI.enabled = true;
	}
}

[InitializeOnLoad]
public class ExcelToJsonAutoConverter 
{	
	/// <summary>
	/// Class attribute [InitializeOnLoad] triggers calling the static constructor on every refresh.
	/// </summary>
	static ExcelToJsonAutoConverter() 
	{
		string inputPath = EditorPrefs.GetString(ExcelToJsonConverterWindow.kExcelToJsonConverterInputPathPrefsName, Application.dataPath);
		string outputPath = EditorPrefs.GetString(ExcelToJsonConverterWindow.kExcelToJsonConverterOuputPathPrefsName, Application.dataPath);
		bool onlyModifiedFiles = EditorPrefs.GetBool(ExcelToJsonConverterWindow.kExcelToJsonConverterModifiedFilesOnlyPrefsName, false);
		
		ExcelToJsonConverter excelProcessor = new ExcelToJsonConverter();
		excelProcessor.ConvertExcelFilesToJson(inputPath, outputPath, onlyModifiedFiles);
	}
}
