namespace ExcelExportTool.Util;

/// <summary>
/// 传递性管道
/// 结果会跟随步骤传递
/// </summary>
/// <typeparam name="T">结果类型约束，如果结构复杂，可以使用dynamic，但需要自己在步骤中维护</typeparam>
public class TransitivePipeline<T>
{
    private readonly List<Func<T, (T parameter, bool isSucceed)>> _steps = [];
    private readonly List<string> _errorMessages = [];
    private string _errorMessage = string.Empty;
    
    public TransitivePipeline<T> AddStep(Func<T, bool> step, string errorMessage)
    {
        _steps.Add(value =>
        {
            var isSucceed = step(value);
            return (value, isSucceed);
        });
        _errorMessages.Add(errorMessage);
        return this;
    }

    public (bool IsSuccess, string ErrorMessage) Execute(T value)
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