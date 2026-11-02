namespace QuickExcel.Runtime;

internal interface IWorkerProcess : IDisposable
{
    Stream Input { get; }
    Stream Output { get; }
    Stream Error { get; }
    int Id { get; }
    string ProfileName { get; }
    string PayloadDirectory { get; }
    void Kill();
}
