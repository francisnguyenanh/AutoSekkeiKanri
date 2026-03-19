using System.Threading;

namespace AutoFiller.UIDriver
{
    /// <summary>
    /// Provides a process-wide singleton <see cref="OcrEngine"/> instance.
    /// Initialisation is deferred until first access and thread-safe.
    /// </summary>
    public static class OcrEngineProvider
    {
        private static readonly Lazy<OcrEngine> _instance =
            new Lazy<OcrEngine>(() => new OcrEngine("ja"),
                LazyThreadSafetyMode.ExecutionAndPublication);

        public static OcrEngine Instance => _instance.Value;
    }
}
