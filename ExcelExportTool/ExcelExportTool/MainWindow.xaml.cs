using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Documents;
using ExcelExportTool.Util;
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
    private const string _DIRTY_EXCEL_FILE_NAME = "excel_dirty_data.json";

    private LogWindow _curLogWindow = null;

    private string _toolPrivateDataFolderPath = null;

    //key:Full Path , value:last load time
    private Dictionary<string, DateTime> _excelDirtyDataDic;
    private List<string> _dirtyExcels = [];

    private OptionalData _optionalData = null;

    public MainWindow()
    {
        InitializeComponent();

        Check_CompletionToolPrivateDataFolder();
        _optionalData = LoadOptionalData();

        InitExcelPathEditorBtn();
        InitXMLPathEditorBtn();
        InitStartBtn();
        InitCleanDirtyDataBtn();
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

            ExportWorkFlow();
        };
    }

    private void InitCleanDirtyDataBtn()
    {
        CleanDirtyDataBtn.Click += (sender, args) =>
        {
            var dirtyDataPath = ExcelDirtyDataFilePathFactory();
            if(File.Exists(dirtyDataPath))
            {
                File.Delete((dirtyDataPath));
            }
        };
    }

    private void SaveData()
    {
        SaveOptionalData();
        SaveDirtyData();
    }

    /// <summary>
    /// 导出工作流程
    /// </summary>
    private void ExportWorkFlow()
    {
        var workFlowPipeline = new NonTransitivePipeline()
            .AddStep(CheckPathConfig, "Path Settings failed")
            .AddStep(LoadDirtyData, "Load Dirty Data failed")
            .AddStep(GetDirtyExcels, "Get Dirty Excels failed")
            .AddStep(GenerationProcess, "Generation Process failed");

        var result = workFlowPipeline.Execute();
        if (!result.IsSuccess)
        {
            _curLogWindow.AddLog(FinishResults.Failure, result.ErrorMessage);
            _curLogWindow.AddLog(FinishResults.Failure, "Please fix bugs by logs, then try again");
        }
        else
        {
            _curLogWindow.AddLog(FinishResults.Success, "Export Work Finished");
        }

        SaveData();
    }

    /// <summary>
    /// 检查路径设置
    /// </summary>
    /// <returns></returns>
    private bool CheckPathConfig()
    {
        if (string.IsNullOrEmpty(_optionalData.curExcelsFolderPath) ||
            !Directory.Exists(_optionalData.curExcelsFolderPath))
        {
            _curLogWindow.AddLog(FinishResults.Failure, "Excel path is Error");
            return false;
        }

        if (string.IsNullOrEmpty(_optionalData.curXMLsFolderPath) || !Directory.Exists(_optionalData.curXMLsFolderPath))
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
    private string OptionalDataFilePathFactory()
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
        try
        {
            var dataString = JsonSerializer.Serialize(_optionalData, new JsonSerializerOptions()
            {
                WriteIndented = true
            });
            File.WriteAllText(OptionalDataFilePathFactory(), dataString);
        }
        catch (Exception e)
        {
            _curLogWindow.AddLog(FinishResults.Failure, "Save Optional Data Error :: " + e);
        }
    }

    /// <summary>
    /// 加载可选数据
    /// </summary>
    /// <returns></returns>
    private OptionalData LoadOptionalData()
    {
        try
        {
            var data = JsonSerializer.Deserialize<OptionalData>(File.ReadAllText(OptionalDataFilePathFactory()));
            return data ?? new OptionalData();
        }
        catch (Exception ex)
        {
            _curLogWindow.AddLog(FinishResults.Failure, $"Error loading optional data: {ex.Message}");
            return new OptionalData();
        }
    }

    /// <summary>
    /// Excel脏数据文件夹路径生成工厂
    /// </summary>
    /// <returns></returns>
    private string ExcelDirtyDataFilePathFactory()
    {
        if (string.IsNullOrEmpty(_toolPrivateDataFolderPath) || !Directory.Exists(_toolPrivateDataFolderPath))
            Check_CompletionToolPrivateDataFolder();

        return Path.Combine(_toolPrivateDataFolderPath, _DIRTY_EXCEL_FILE_NAME);
    }

    /// <summary>
    /// 加载脏数据
    /// </summary>
    private bool LoadDirtyData()
    {
        var dirtyFilePath = ExcelDirtyDataFilePathFactory();

        try
        {
            _excelDirtyDataDic = File.Exists(dirtyFilePath)
                ? _excelDirtyDataDic =
                    JsonSerializer.Deserialize<Dictionary<string, DateTime>>(File.ReadAllText(dirtyFilePath))
                : new Dictionary<string, DateTime>();

            return true;
        }
        catch (Exception e)
        {
            _curLogWindow.AddLog(FinishResults.Failure, "Load Dirty Data Error :: " + e);
            return false;
        }
    }

    /// <summary>
    /// 保存脏数据
    /// </summary>
    private void SaveDirtyData()
    {
        try
        {
            var dataString = JsonSerializer.Serialize(_excelDirtyDataDic, new JsonSerializerOptions()
            {
                WriteIndented = true
            });
            File.WriteAllText(ExcelDirtyDataFilePathFactory(), dataString);
        }
        catch (Exception e)
        {
            _curLogWindow.AddLog(FinishResults.Failure, "Save Dirty Data Error :: " + e);
        }
    }

    /// <summary>
    /// 获取脏Excel
    /// </summary>
    /// <returns></returns>
    private bool GetDirtyExcels()
    {
        if (string.IsNullOrEmpty(_optionalData.curExcelsFolderPath))
        {
            _curLogWindow.AddLog(FinishResults.Failure, "Excel Read Path is Null");
            return false;
        }
        
        var excels = Directory.GetFiles(_optionalData.curExcelsFolderPath, "*.xlsx", SearchOption.AllDirectories).ToList();
        if (excels.Count <= 0)
        {
            _curLogWindow.AddLog(FinishResults.Warning, "There are no '.xlsx' files in the current read directory");
            return false;
        }

        _dirtyExcels.Clear();
        foreach (var excel in excels)
        {
            if (_excelDirtyDataDic.TryGetValue(excel, out var lastLoadTime))
            {
                if (lastLoadTime < File.GetLastWriteTime(excel))
                {
                    _dirtyExcels.Add(excel);
                }
            }
            else
            {
                _excelDirtyDataDic.Add(excel, new DateTime());
                _dirtyExcels.Add(excel);
            }
        }
        if (_dirtyExcels.Count <= 0)
        {
            _curLogWindow.AddLog(FinishResults.Warning, "There are no updated file in the current read directory");
            return false;
        }
        
        return true;
    }

    /// <summary>
    /// 生成流程
    /// </summary>
    /// <returns></returns>
    private bool GenerationProcess()
    {
        var generationWorkFlow = new TransitivePipeline<dynamic>();
        foreach (var excel in _dirtyExcels)
        {
            generationWorkFlow.ClearSteps();
            generationWorkFlow.AddStep(LoadExcelData2Array, "");
        }

        return true;
    }

    private bool LoadExcelData2Array(dynamic excelPath)
    {
        if (!File.Exists(excelPath))
        {
            return false;
        }
        
        _curLogWindow.AddLog(FinishResults.Default, $"Start Load Excel Data :: {excelPath}");
        
        return true;
    }
}