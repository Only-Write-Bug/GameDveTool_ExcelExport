using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;

namespace ExcelExportTool;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private LogWindow _curLogWindow = null;

    public MainWindow()
    {
        InitializeComponent();

        InitExcelPathEditorBtn();
        InitXMLPathEditorBtn();
        InitStartBtn();
    }

    private void InitPath()
    {
        
    }

    private void InitExcelPathEditorBtn()
    {
        ExcelPathEditorButton.Click += (sender, args) =>
        {
            var folderDialog = new OpenFolderDialog();
            if (folderDialog.ShowDialog() == true)
            {
                ExcelPathTextBox.Text = folderDialog.FolderName;
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
                XMLPathTextBox.Text = folderDialog.FolderName;
            }
        };
    }

    private void InitStartBtn()
    {
        StartButton.Click += (sender, args) =>
        {
            var logWindow = new LogWindow();
            logWindow.Show();
            _curLogWindow = logWindow;
            _curLogWindow.AddLog(FinishResults.Default, Directory.GetCurrentDirectory());
        };
    }
}