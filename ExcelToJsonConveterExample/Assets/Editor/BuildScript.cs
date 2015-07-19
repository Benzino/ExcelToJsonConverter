using UnityEditor;
using UnityEngine;
using System;
using System.Diagnostics;
using System.Collections.Generic;

public class ScriptBatch
{
	private static Action _excelSuccessCallback;
	private static BuildOptions _buildOptions;

	[MenuItem("MyTools/IOS Release Build")]
	public static void PerformIOSReleaseBuild()
	{
		_excelSuccessCallback = DoIOSReleaseBuild;

		_buildOptions = BuildOptions.Il2CPP;

		// Process the Excel Files
		ProcessExcelFiles();
	}

	[MenuItem("MyTools/IOS Release Build & Run")]
	public static void PerformIOSReleaseBuildAndRun()
	{
		_excelSuccessCallback = DoIOSReleaseBuild;

		_buildOptions = BuildOptions.Il2CPP | BuildOptions.AutoRunPlayer;
		
		// Process the Excel Files
		ProcessExcelFiles();
	}

	/// <summary>
	/// Performs the iOS Release Build
	/// </summary>
	private static void DoIOSReleaseBuild()
	{
		// Get filename.
		string path = EditorUtility.SaveFolderPanel("Choose Location of Built Game", "", "");
		if (string.IsNullOrEmpty(path))
		{
			return;
		}

		// Build player.
		BuildPipeline.BuildPlayer(GetEnabledBuildScenes(), path, BuildTarget.iOS, _buildOptions);
		
		// Run the game (Process class from System.Diagnostics).
		Process proc = new Process();
		proc.StartInfo.FileName = path;
		proc.Start();
	}

	[MenuItem("MyTools/IOS Debug Build")]
	public static void PerformIOSDebugBuild()
	{
		_excelSuccessCallback = DoIOSDebugBuild;

		_buildOptions = BuildOptions.Il2CPP | BuildOptions.Development | BuildOptions.AllowDebugging | BuildOptions.SymlinkLibraries;

		// Process the Excel Files
		ProcessExcelFiles();
	}

	[MenuItem("MyTools/IOS Debug Build & Run")]
	public static void PerformIOSDebugBuildAndRun()
	{
		_excelSuccessCallback = DoIOSDebugBuild;

		_buildOptions = BuildOptions.Il2CPP | BuildOptions.Development | 
			BuildOptions.AllowDebugging | BuildOptions.SymlinkLibraries  | BuildOptions.AutoRunPlayer;
		
		// Process the Excel Files
		ProcessExcelFiles();
	}

	/// <summary>
	/// Performs the iOS Debug Build
	/// </summary>
	private static void DoIOSDebugBuild()
	{
		// Get filename.
		string path = EditorUtility.SaveFolderPanel("Choose Location of Built Game", "", "");
		if (string.IsNullOrEmpty(path))
		{
			return;
		}
		
		// Build player.
		BuildPipeline.BuildPlayer(GetEnabledBuildScenes(), path, BuildTarget.iOS, _buildOptions);
		
		// Run the game (Process class from System.Diagnostics).
		Process proc = new Process();
		proc.StartInfo.FileName = path;
		proc.Start();
	}

	/// <summary>
	/// Gets the list of enabled scenes that are added to the build settings
	/// via the Build Settings window
	/// </summary>
	/// <returns>The enabled build scenes.</returns>
	private static string[] GetEnabledBuildScenes()
	{
		List<EditorBuildSettingsScene> scenes = new List<EditorBuildSettingsScene>(EditorBuildSettings.scenes);
		List<string> enabledScenes = new List<string>();
		foreach (EditorBuildSettingsScene scene in scenes)
		{
			if (scene.enabled)
			{
				enabledScenes.Add(scene.path);
			}
		}

		return enabledScenes.ToArray();
	}
	
	/// <summary>
	/// Processes the excel files.
	/// </summary>
	private static void ProcessExcelFiles()
	{
		ExcelToJsonConverter excelProcessor = new ExcelToJsonConverter();
		excelProcessor.ConversionToJsonSuccessfull += ExcelSuccessCallback;
		excelProcessor.ConvertExcelFilesToJson(EditorPrefs.GetString(ExcelToJsonConverterWindow.kExcelToJsonConverterInputPathPrefsName, Application.dataPath), 
		                                 EditorPrefs.GetString(ExcelToJsonConverterWindow.kExcelToJsonConverterOuputPathPrefsName, Application.dataPath),
		                                 false);
	}

	/// <summary>
	/// Callback method for Processing Excel sheets
	/// </summary>
	private static void ExcelSuccessCallback()
	{
		if (_excelSuccessCallback != null)
		{
			_excelSuccessCallback();
			_excelSuccessCallback = null;
		}
	}
}