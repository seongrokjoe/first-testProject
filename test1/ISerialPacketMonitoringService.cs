namespace Simulator.Module.VxStudio.Models.Monitor.Serial
{
    public interface ISerialPacketMonitoringService : IDisposable
    {
        void Initialize(string basePath);
        IEnumerable<SerialPacketEndpointKey> ListEndpointsByUnit(string unit, bool includeRuntime = true);
        SerialPacketSnapshot GetSnapshot(SerialPacketEndpointKey key);
        void StartMonitoring(SerialPacketEndpointKey key);
        void StopMonitoring(SerialPacketEndpointKey key);
        IReadOnlyList<PacketLine> DrainBuffer(SerialPacketEndpointKey key);
        event Action<SerialPacketEndpointKey, PacketLine> LineArrived;
        event Action<SerialPacketEndpointKey, IReadOnlyList<PacketLine>> LinesPublished;
    }

    public readonly record struct SerialPacketEndpointKey(string Unit, string Device, string Port)
    {
        public override string ToString() => $"{Unit}|{Device}|{Port}";
        public static BuildKey(string unit, string device, string port) => $"{unit}|{device}|{port}";
    }

    public sealed class PacketLine
    {
        public DateTime Timestamp {get; init;}
        public bool IsRecv {get; init;}
        public string Hex {get; init;}
        public PacketLine(DateTime timestamp, bool isRecv, string hex)
        {
            Timestamp = timestamp;
            IsRecv = isRecv;
            Hex = hex ?? string.Empty;
        }
    }

    public sealed class SerialPacketSnapshot
    {
        public System.DateTime? LastRecvAt {get; set;}
        public string? LastRecvHex {get; set; }
        public System.DateTime? LastSentAt {get; set; }
        public string? LastSendHex {get; set; }
    }
}