﻿using System.IO;
using System.Text.Json;
using System.Windows;
using Microsoft.Win32;
using Path = System.IO.Path;

namespace ExcelExportTool;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private const string _PRIVATE_DATA_FOLDER_NAME = "_data";
    private const string _OPTIONAL_DATA_FILE_NAME = "optional_data.json";
    
    private LogWindow _curLogWindow = null;
    
    private string _toolPrivateDataFolderPath = null;
    
    private OptionalData _optionalData = null;

    public MainWindow()
    {
        InitializeComponent();

        Check_CompletionToolPrivateDataFolder();
        _optionalData = LoadOptionalData();
        
        InitExcelPathEditorBtn();
        InitXMLPathEditorBtn();
        InitStartBtn();
    }

    private void InitExcelPathEditorBtn()
    {
        ExcelPathTextBox.Text = _optionalData.curExcelsFolderPath;
        ExcelPathEditorButton.Click += (sender, args) =>
        {
            var folderDialog = new OpenFolderDialog();
            if (folderDialog.ShowDialog() == true)
            {
                _optionalData.curExcelsFolderPath = ExcelPathTextBox.Text = folderDialog.FolderName;
            }
        };
    }

    private void InitXMLPathEditorBtn()
    {
        XMLPathTextBox.Text = _optionalData.curXMLsFolderPath;
        XmlPathEditorButton.Click += (sender, args) =>
        {
            var folderDialog = new OpenFolderDialog();
            if (folderDialog.ShowDialog() == true)
            {
                _optionalData.curXMLsFolderPath = XMLPathTextBox.Text = folderDialog.FolderName;
            }
        };
    }

    private void InitStartBtn()
    {
        StartButton.Click += (sender, args) =>
        {
            var logWindow = new LogWindow();
            logWindow.Show();
            
            _curLogWindow?.Close();
            _curLogWindow = logWindow;
            _curLogWindow.AddLog(FinishResults.Success, "Tool is running!!!");
            
            SaveOptionalData();
            
            if (!CheckPathConfig())
            {
                _curLogWindow.AddLog(FinishResults.Failure, "Work failed, please fix bug by log, try again");
                return;
            }
        };
    }

    /// <summary>
    /// 检查路径设置
    /// </summary>
    /// <returns></returns>
    private bool CheckPathConfig()
    {
        if (string.IsNullOrEmpty(_optionalData.curExcelsFolderPath) || Directory.Exists(_optionalData.curExcelsFolderPath))
        {
            _curLogWindow.AddLog(FinishResults.Failure, "Excel path is Error");
            return false;
        }
        if (string.IsNullOrEmpty(_optionalData.curXMLsFolderPath) || Directory.Exists(_optionalData.curXMLsFolderPath))
        {
            _curLogWindow.AddLog(FinishResults.Failure, "XML path is Error");
            return false;
        }

        _curLogWindow.AddLog(FinishResults.Success, "Path Check is Success");

        return true;
    }

    /// <summary>
    /// 检查并补全工具私有数据文件夹
    /// </summary>
    private void Check_CompletionToolPrivateDataFolder()
    {
        _toolPrivateDataFolderPath = Path.Combine(Directory.GetCurrentDirectory(), _PRIVATE_DATA_FOLDER_NAME);
        if (!Directory.Exists(_toolPrivateDataFolderPath))
        {
            Directory.CreateDirectory(_toolPrivateDataFolderPath);
            var directoryInfo = new DirectoryInfo(_toolPrivateDataFolderPath);
            directoryInfo.Attributes |= FileAttributes.Hidden;
        }
    }

    /// <summary>
    /// 可选数据路径生成工厂
    /// </summary>
    /// <returns></returns>
    private string OptionalDataFileFactory()
    {
        if (string.IsNullOrEmpty(_toolPrivateDataFolderPath) || !Directory.Exists(_toolPrivateDataFolderPath))
            Check_CompletionToolPrivateDataFolder();
        
        return Path.Combine(_toolPrivateDataFolderPath, _OPTIONAL_DATA_FILE_NAME);
    }

    /// <summary>
    /// 保存可选数据
    /// </summary>
    private void SaveOptionalData()
    {
        var dataString = JsonSerializer.Serialize(_optionalData, new JsonSerializerOptions()
        {
            WriteIndented = true
        });
        File.WriteAllText(OptionalDataFileFactory(), dataString);
    }

    /// <summary>
    /// 加载可选数据
    /// </summary>
    /// <returns></returns>
    private OptionalData LoadOptionalData()
    {
        if (string.IsNullOrEmpty(_toolPrivateDataFolderPath) || !Directory.Exists(_toolPrivateDataFolderPath) || !File.Exists(OptionalDataFileFactory()))
            return new OptionalData();
        
        var data = JsonSerializer.Deserialize<OptionalData>(File.ReadAllText(OptionalDataFileFactory()));
        
        return data ?? new OptionalData();
    }
}