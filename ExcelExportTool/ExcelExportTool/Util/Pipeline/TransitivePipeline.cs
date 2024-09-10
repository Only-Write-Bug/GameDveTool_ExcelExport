namespace ExcelExportTool.Util;

/// <summary>
/// 传递性管道
/// 结果会跟随步骤传递
/// </summary>
public class TransitivePipeline
{
    private readonly List<Func<dynamic, (dynamic parameter, bool isSucceed)>> _steps = [];
    private readonly List<string> _errorMessages = [];
    private string _errorMessage = string.Empty;
    
    public TransitivePipeline AddStep(Func<dynamic, (dynamic parameter, bool isSucceed)> step, string errorMessage)
    {
        _steps.Add(parameter =>
        {
            var result = step(parameter);
            return (result.parameter, result.isSucceed);
        });
        _errorMessages.Add(errorMessage);
        return this;
    }

    public (bool IsSuccess, string ErrorMessage) Execute(dynamic value)
    {
        for (var i = 0; i < _steps.Count; i++)
        {
            var result = _steps[i](value);
            value = result.parameter;
            if (!result.isSucceed)
            {
                _errorMessage = _errorMessages[i];
                return (false, _errorMessage);
            }
        }
        return (true, string.Empty);
    }

    public void ClearSteps()
    {
        _steps.Clear();
        _errorMessages.Clear();
        _errorMessage = string.Empty;
    }
}