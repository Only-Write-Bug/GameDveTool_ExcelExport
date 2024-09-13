namespace ExcelExportTool.Util;

/// <summary>
/// 传递性管道
/// 结果会跟随步骤传递
/// </summary>
public class TransitivePipeline
{
    private Queue<Func<object[], (bool isSucces, object[] result)>> _steps = [];
    private Queue<string> _errorMessages = [];

    public TransitivePipeline AddStep(Func<object[], (bool isSucces, object[] result)> step, string errorMessage)
    {
        _steps.Enqueue(step);
        _errorMessages.Enqueue(errorMessage);
        
        return this;
    }

    public (bool isSucces, object[] result) Execute(object[] startParams)
    {
        (bool isSucces, object[] result) lastStepResult = (true, startParams);
        while (_steps.Count > 0)
        {
            var step = _steps.Dequeue();
            var errorMessage = _errorMessages.Dequeue();

            if (!lastStepResult.isSucces)
            {
                return (false, new object[] { errorMessage });
            }
            lastStepResult = step.Invoke(lastStepResult.result);
        }
        
        return (true, [string.Empty]);
    }

    public void Clear()
    {
        _steps.Clear();
        _errorMessages.Clear();
    }

    ~TransitivePipeline()
    {
        Clear();
    }
}