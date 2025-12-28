namespace Simulator.Module.Common.MsgBoxService
{
    public interface IMsgBoxService
    {
        void Info(string msg, bool preventDup = false, bool modeless = false);
        void Warn(string msg, bool preventDup = false, bool modeless = false);
        void Error(string msg, bool preventDup = false, bool modeless = false);

        Task<MessageBoxResult> ShowAsync(string message,
        string caption,
        MessageBoxButton buttons,
        MessageBoxImage icon,
        bool preventDup = false,
        bool modeless = false,
        TimeSpan? coolDown = null,
        string? key = null,
        CancellationToken ct = default);
    }
}