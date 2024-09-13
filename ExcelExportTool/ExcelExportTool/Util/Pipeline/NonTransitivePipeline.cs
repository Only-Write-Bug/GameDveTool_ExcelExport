namespace ExcelExportTool.Util;

/// <summary>
/// 非传递性管线
/// 结果不会跟随步骤传递
/// </summary>
public class NonTransitivePipeline
{
    private readonly Queue<Func<bool>> _steps = [];
    private readonly Queue<string> _errorMessages = [];

    public NonTransitivePipeline AddStep(Func<bool> step, string errorMessage)
    {
        _steps.Enqueue(step);
        _errorMessages.Enqueue(errorMessage);
        return this;
    }

    public (bool IsSuccess, string ErrorMessage) Execute()
    {
        while (_steps.Count > 0)
        {
            var func = _steps.Dequeue();
            var errorMessage = _errorMessages.Dequeue();

            if (!func())
                return (false, errorMessage);
        }

        return (true, string.Empty);
    }

    public void ClearSteps()
    {
        _steps.Clear();
        _errorMessages.Clear();
    }

    ~NonTransitivePipeline()
    {
        ClearSteps();
    }
}