using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ExcelExportTool;

public partial class LogWindow : Window
{
    public LogWindow()
    {
        InitializeComponent();
    }
    
    public void AddLog(FinishResults finishResult, string content)
    {
        logBox.Items.Add(LogTemplate(finishResult, content));
    }

    private TextBlock LogTemplate(FinishResults finishResult, string content)
    {
        var result = new TextBlock();
        result.TextWrapping = TextWrapping.Wrap;
        result.VerticalAlignment = VerticalAlignment.Stretch;
        result.FontSize = 15;
        result.Text = content;
        result.Width = 1100;
        
        var brush = new SolidColorBrush(Colors.White);
        switch (finishResult)
        {
            case FinishResults.Success:
                brush = new SolidColorBrush(Colors.Green);
                break;
            case FinishResults.Failure:
                brush = new SolidColorBrush(Colors.Red);
                break;
            case FinishResults.Warning:
                brush = new SolidColorBrush(Colors.Yellow);
                break;
            case FinishResults.Default:
            default:
                break;
        }
        result.Foreground = brush;

        return result;
    }
}