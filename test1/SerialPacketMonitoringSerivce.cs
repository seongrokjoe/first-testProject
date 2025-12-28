using System.Collections.Concurrent;

namespace Simulator.Module.VxStudio.Models.Monitor.Serial
{
    public sealed class SerialPacketMonitoringSerivce : ISerialPacketMonitoringService
    {
        private string? _basePath;
        private bool _initialized;

        private readonly ConcurrentDictionary<SerialPacketEndpointKey, byte> _watch = new();
        private readonly ConcurrentDictionary<SerialPacketEndpointKey, ConcurrentQueue<PacketLine>> _queues = new();
        private readonly ConcurrentDictionary<SerialPacketEndpointKey, SerialPacketSnapshot> _snapshots = new();

        private IDisposable? _feedSubscription;
        private CancellationTokenSource? _cts;
        private readonly TimeSpan _flushInterval = TimeSpan.FromMilliseconds(500);

        public event Action<SerialPacketEndpointKey, PacketLine>? LineArrived;
        public event Action<SerialPacketEndpointKey, IReadOnlyList<PacketLine>>? LinesPublished;

        public void Initialize(string basePath)
        {
            if (_initialized) return;
            _basePath = basePath ?? throw new ArgumentNullException(nameof(basePath));

            _feedSubscription = SerialPacketFeed.Subscribe(OnScriptEnginePacket);

            _cts = new CancellationTokenSource();
            _ = FlushLoopAsync(_cts.Token);

            _initialized = true;
        }

        public void Dispose()
        {
            try
            {
                _feedSubscription?.Dispose();
                _cts?.Cancel();
                _cts?.Dispose();    
            }
            catch
            {
                throw new ObjectDisposedException(nameof(SerialPacketMonitoringSerivce));
            }
        }

        public IEnumerable<SerialPacketEndpointKey> ListEndpointsByUnit(string unit, bool includeRuntime = true)
        {
            if (string.IsNullOrWhiteSpace(unit))
                return Enumerable.Empty<SerialPacketEndpointKey>();
            
            IEnumerable<(string Unit, string Device, string Port)> all;
            if (includeRuntime)
            {
                var tuples = SerialDeviceManager.Instance.ListAllKnownEndpoints(_basePath);
                all = tuples.Select(t => (t.UnitName, t.DeviceName, t.PortName));
            }
            else
            {
                var tuples = SerialDeviceManager.Instance.ListDeviceEndpointsFromConfig(_basePath);
                all = tuples.Select(t -> (t.UnitName, t.DeviceName, t.PortName));
            }

            return all.Where(t => string.Equals(t.Unit, unit, StringComparison.Ordinal)).Select(t => new SerialPacketEndpointKey(t.Unit, t.Device, t.Port)).Distinct().OrderBy(k => k.Device).ThenBy(k => k.Port);
        }

        public SerialPacketSnapshot GetSnapshot(SerialPacketEndpointKey key)
        {
            if (_snapshots.TryGetValue(key, out var s))
            {
                return s;
            }
            return new SerialPacketSnapshot();
        }

        public void StartMonitoring(SerialPacketEndpointKey key)
        {
            _watch[key] = 1;
            _queues.GetOrAdd(key, _ => new ConcurrentQueue<PacketLine>());
        }

        public void StopMonitoring(SerialPacketEndpointKey key)
        {
            _watch.TryRemove(key, out _);
        }

        public IReadOnlyList<PacketLine> DrainBuffer(SerialPacketEndpointKey key)
        {
            if (!_queues.TryGetValue(key, out var q) || q.IsEmpty)
            {
                return Array.Empty<PacketLine>();
            }

            var list = new List<PacketLine>(128);
            while (q.TryDequeue(out var l))
            {
                list.Add(l);
            }
            return list;
        }

        private void OnScriptEnginePacket(DeviceKey dk, bool isRecv, byte[] data)
        {
            var key = new SerialPacketEndpointKey(dk.Unit, dk.Device, dk.Port);
            var now = DateTime.Now;
            var hex = BytesToHex(data);

            var s = _snapshots.GetOrAdd(key, _ => new SerialPacketSnapshot());

            if (isRecv)
            {
                if (!string.Equals(s.LastRecvHex, hex, StringComparison.Ordinal))
                {
                    s.LastRecvHex = hex;
                    s.LastRecvAt = now;
                }
            }
            else
            {
                if (!string.Equals(s.LastSentHex, hex, StringComparison.Ordinal))
                {
                    s.LastSentHex = hex;
                    s.LastSentAt = now;
                }
            }

            var line = new PacketLine(now, isRecv, hex);
            LineArrived?.Invoke(key, line);

            if (_watch.ContainsKey(key))
            {
                var q = _queues.GetOrAdd(key, _ => new ConcurrentQueue<PacketLine>());
                q.Enqueue(line);
            }
        }

        private async Task FlushLoopAsync(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {

                    await Task.Delay(_flushInterval, ct);

                    foreach (var key in _watch.Keys.ToArray())
                    {
                        if (!_queues.TryGetValue(key, out var q) || q.IsEmpty)
                            continue;

                        var batch = new List<PacketLine>(256);
                        while (q.TryDequeue(out var l))
                        {
                            batch.Add(l);
                        }

                        if (batch.Count > 0)
                        {
                            LinesPublished?.Invoke(key, batch);
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch {}
                
                
            }
        }

        private static string BytesToHex(byte[] data)
        {
            if (data == null || data.Length == 0)
                return string.Empty;
            
            var chars = new char[data.Length * 5 - 1];
            int p = 0;

            for (int i = 0; i < data.Length; i++)
            {
                if (i > 0) chars[p++] = ' ';
                chars[p++] = '0';
                chars[p++] = 'x';
                var b = data[i];
                chars[p++] = GexHex((b >> 4) & 0xF);
                chars[p++] = GexHex(b & 0xF);
            }
            return new string(chars);
        }

        private static char GetHex(int v) => (char)(v < 10 ? ('0' + v) : ('A' + (v - 10)));



    }
}