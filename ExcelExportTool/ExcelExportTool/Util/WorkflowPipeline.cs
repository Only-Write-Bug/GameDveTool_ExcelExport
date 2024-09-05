namespace ExcelExportTool.Util;

public class WorkflowPipeline
{
    private readonly List<Func<bool>> _steps = [];
    private readonly List<string> _errorMessages = [];
    private string _errorMessage = string.Empty;

    public WorkflowPipeline AddStep(Func<bool> step, string errorMessage)
    {
        _steps.Add(step);
        _errorMessages.Add(errorMessage);
        return this;
    }

    public (bool IsSuccess, string ErrorMessage) Execute()
    {
        for (int i = 0; i < _steps.Count; i++)
        {
            if (!_steps[i]())
            {
                _errorMessage = _errorMessages[i];
                return (false, _errorMessage);
            }
        }
        return (true, string.Empty);
    }
}