namespace ExcelExportTool.Util;

/// <summary>
/// 非传递性管线
/// 结果不会跟随步骤传递
/// </summary>
public class NonTransitivePipeline
{
    private readonly List<Func<bool>> _steps = [];
    private readonly List<string> _errorMessages = [];
    private string _errorMessage = string.Empty;

    public NonTransitivePipeline AddStep(Func<bool> step, string errorMessage)
    {
        _steps.Add(step);
        _errorMessages.Add(errorMessage);
        return this;
    }

    public (bool IsSuccess, string ErrorMessage) Execute()
    {
        for (var i = 0; i < _steps.Count; i++)
        {
            if (!_steps[i]())
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