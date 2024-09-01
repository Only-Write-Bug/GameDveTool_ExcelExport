using System.IO;
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
    
    private LogWindow _curLogWindow = null;
    
    private string _toolPrivateDataPath= null;
    private string _curExcelsPath = null;
    private string _curXMLsPath = null; 

    public MainWindow()
    {
        InitializeComponent();

        Check_CompletionToolPrivateDataFolder();
        InitExcelPathEditorBtn();
        InitXMLPathEditorBtn();
        InitStartBtn();
    }

    private void InitExcelPathEditorBtn()
    {
        ExcelPathEditorButton.Click += (sender, args) =>
        {
            var folderDialog = new OpenFolderDialog();
            if (folderDialog.ShowDialog() == true)
            {
                _curExcelsPath = ExcelPathTextBox.Text = folderDialog.FolderName;
            }
        };
    }

    private void InitXMLPathEditorBtn()
    {
        XmlPathEditorButton.Click += (sender, args) =>
        {
            var folderDialog = new OpenFolderDialog();
            if (folderDialog.ShowDialog() == true)
            {
                _curXMLsPath = XMLPathTextBox.Text = folderDialog.FolderName;
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
        if (string.IsNullOrEmpty(_curExcelsPath) || Directory.Exists(_curExcelsPath))
        {
            _curLogWindow.AddLog(FinishResults.Failure, "Excel path is Error");
            return false;
        }
        if (string.IsNullOrEmpty(_curXMLsPath) || Directory.Exists(_curXMLsPath))
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
        _toolPrivateDataPath = Path.Combine(Directory.GetCurrentDirectory(), _PRIVATE_DATA_FOLDER_NAME);
        if (!Directory.Exists(_toolPrivateDataPath))
        {
            Directory.CreateDirectory(_toolPrivateDataPath);
            var directoryInfo = new DirectoryInfo(_toolPrivateDataPath);
            directoryInfo.Attributes |= FileAttributes.Hidden;
        }
    }
}