using System.Collections.Concurrent;
using System.Diagnostics.Contracts;

namespace Simulator.Module.VxStudio.Models.IOConfig
{
    public sealed class ReadSpec
    {
        public string Key {get; private set;}
        public IoType IoType {get; private set;}
        public Dir Direction {get; private set;}
        public int ByteAddress {get; private set;}
        public int BitAddress {get; private set;}
        public int Length {get; private set;}
        public int DecimalPoint {get; private set;}
        public int ByteOffset {get; private set;}
        public int Min {get; private set;}
        public int Max {get; private set;}
        public int DrvMin {get; private set;}
        public int DrvMax {get; private set;}

        public static ReadSpec FromIO(IOConfigItem io) => new ReadSpec
        {
            Key = $"{io.UnitName}|{io.Variable}",
            IoType = io.Type,
            Direction = io.Direction,
            ByteAddress = io.ByteAddress,
            BitAddress = io.BitAddress,
            Length = io.Length,
            DecimalPoint = io.DecimalPoint,
            ByteOffset = io.ByteOffset,
            Min = io.Min,
            Max = io.Max,
            DrvMin = io.DrvMin,
            DrvMax = io.DrvMax
        };
    }

    public sealed class SimpleRead
    {
        public string Key {get; init;}
        public object Value {get; init;}
    }

    public sealed class ValueHistory
    {
        public DateTime Timestamp {get; init;}
        public double Value {get; init;}
    }

    public interface IWMXIOMonitoringService
    {
        Task<IReadOnlyList<SimpleRead>> ReadAsync(IEnumerable<ReadSpec> specs);
        Task ApplySimplesAsync(IDictionary<ReadSpec, double> overrides);
        IReadOnlyList<ValueHistory> GetHistory(string key);
    }

    public sealed class WMXIOMonitoringService : IWMXIOMonitoringService
    {
        private readonly WMXService _wmxService;
        public WMXIOMonitoringService(WMXService wmxService)
        {
            _wmxService = wmxService;
        }

        private sealed class ChannelState
        {
            public double CurrentValue;
            public List<ValueHistory> History = new();
            public bool IsDigital;
        }

        private readonly ConcurrentDictionary<string, ChannelState> _channels = new();
        private readonly ConcurrentDictionary<string, IoType> _keyTypes = new();
        private readonly object _historyLock = new();

        private const int MaxHistory = 5000;

        public Task<IReadOnlyList<SimpleRead> ReadAsync(IEnumerable<ReadSpec> specs)
        {
            if (specs == null)
                return Task.FromResult((IReadOnlyList<SimpleRead>)Array.Empty<SimpleRead>());

            var nowUtc = DateTime.UtcNow;
            var result = new List<SimpleRead>();

            foreach (var spec in specs)
            {
                var key = spec.Key ?? "";

                _keyTypes[key] = spec.IoType;

                var st = _channels.GetOrAdd(key, _ =>
                {
                    return new ChannelState
                    {
                        CurrentValue = 0.0,
                        IsDigital = spec.IoType == IoType.Digital
                    };
                });

                if (st.IsDigital)
                {
                    byte byteData;
                    if (spec.Direction == Dir.In)
                    {
                        if (_wmxService.GetInBit(spec.ByteAddress, spec.BitAddress, out byteData))
                            st.CurrentValue = byteData;
                    }
                    else
                    {
                        if (_wmxService.GetOutBit(spec.ByteAddress, spec.BitAddress, out byteData))
                            st.CurrentValue = byteData;
                    }
                }
                else
                {
                    int size = spec.Length / 8;
                    byte[] bytesData = new byte[size];
                    if (spec.Direction == Dir.In)
                    {
                        if (_wmxService.GetInBytes(spec.ByteAddress, size, out bytesData))
                        {
                            int cvtlValue = ConvertBytesToInt(size, bytesData, spec.ByteOffset);
                            st.CurrentValue = ScaleFromDriver(spec, cvtlValue);
                        }
                    }
                    else
                    {
                        if (_wmxService.GetOutBytes(spec.ByteAddress, size, out bytesData))
                        {
                            int cvtlValue = ConvertBytesToInt(size, bytesData, spec.ByteOffset);
                            st.CurrentValue = ScaleFromDriver(spec, cvtlValue);
                        }
                    }
                }

                var record = new ValueHistory
                {
                    Timestamp = nowUtc,
                    Value = st.CurrentValue
                };
                lock (_historyLock)
                {
                    st.History.Add(record);
                    if (st.History.Count > MaxHistory)
                        st.History.RemoveRange(0, st.History.Count - MaxHistory);
                }
                result.Add(new SimpleRead {Key = key, Value = st.CurrentValue});    

            }
            return Task.FromResult((IReadOnlyList<SimpleRead>)result);
        }

        public Task ApplySimplesAsync(IDictionary<ReadSpec, double> overrides)
        {
            if (overrides == null)
                return Task.CompletedTask;

            var nowUtc = DateTime.UtcNow;

            foreach (var kv in overrides)
            {
                var key = kv.Key.Key ?? "";
                var type = kv.Key.IoType;

                var st = _channels.GetOrAdd(key, _ => new ChannelState
                {
                    CurrentValue = 0,
                    IsDigital = type == IoType.Digital
                });

                if (st.IsDigital)
                {
                    byte byteData = kv.Value == 0 ? (byte)0 : (byte)1;
                    if (kv.Key.Direction == Dir.In)
                    {
                        if (_wmxService.SetInBit(kv.Key.ByteAddress, kv.Key.BitAddress, byteData))
                            st.CurrentValue = byteData;
                    }
                    else
                    {
                        if (_wmxService.SetOutBit(kv.Key.ByteAddress, kv.Key.BitAddress, byteData))
                            st.CurrentValue = byteData;
                    }
                }
                else
                {
                    int size = kv.Key.Length / 8;
                    byte[] bytesData = new byte[size];
                    int iValue = ScaleToDriver(kv.Key, kv.Value);
                    bytesData = ConvertIntToBytes(size, iValue, kv.Key.ByteOffset);

                    if(kv.Key.Direction == Dir.In)
                    {
                        if (_wmxService.SetInBytes(kv.Key.ByteAddress, size, bytesData))
                            st.CurrentValue = kv.Value;
                    }
                    else
                    {
                        if (_wmxService.SetOutBytes(kv.Key.ByteAddress, size, bytesData))
                            st.CurrentValue = kv.Value;
                    }
                }

                var record = new ValueHistory
                {
                    Timestamp = nowUtc,
                    Value = st.CurrentValue
                };
                lock (_historyLock)
                {
                    st.History.Add(record);
                    if (st.History.Count > MaxHistory)
                        st.History.RemoveRange(0, st.History.Count - MaxHistory);
                }
            }
            return Task.CompletedTask;
        }

        public IReadOnlyList<ValueHistory> GetHistory(string key)
        {
            if (_channels.TryGetValue(key ?? "", out var st))
            {
                lock (_historyLock)
                    return st.History.ToArray();
            }
            return Array.Empty<ValueHistory>();
        }

        private double ScaleFromDriver(ReadSpec spec, int driverValue)
        {
            if (spec.DrvMax == spec.DrvMin) return spec.Min;

            int drv = Math.Max(spec.DrvMin, Math.Min(driverValue, spec.DrvMax));

            double drvRange = spec.DrvMax - spec.DrvMin;
            double frac = (drv - spec.DrvMin) / drvRange;

            double physRange = spec.Max - spec.Min;
            double scaled = spec.Min + frac * physRange;

            int convertedDecimal = spec.DecimalPoint == 0 ? 1 : sepc.DecimalPoint;
            return scaled / convertedDecimal;
        }

        private int ScaleToDriver(ReadSpec spec, double value)
        {
            int convertedDecimal = spec.DecimalPoint == 0 ? 1 : spec.DecimalPoint;
            value *= convertedDecimal;
            
            if (spec.Max == spec.Min) return spec.DrvMin;
            value = Math.Max(spec.Min, Math.Min(value, spec.Max));
            double frac = (value - spec.Min) / (spec.Max - spec.Min);
            return (int)(spec.DrvMin + frac * (spec.DrvMax - spec.DrvMin));
        }

        private int ConvertBytesToInt(int byteSize, byte[] data, int byteOffset = 0)
        {
            if (byteOffset >= byteSize) byteOffset = 0;
            if (data.Length < byteSize) byteSize = data.Length;
            int intValue = 0;
            for (int i = byteOffset; i < byteSize; i++)
            {
                intValue |= data[i] << (8 * (i - byteOffset));
            }
            return intValue;
        }

        public static byte[] ConvertIntToBytes(int byteSize, int iValue, int byteOffset = 0)
        {
            if (byteOffset >= byteSize) byteOffset = 0;
            
            byte[] result = new byte[byteSize];
            for (int i = byteOffset; i < byteSize; i++)
            {
                result[i] = (byte)(iValue & 0xFF);
                iValue >>= 8;
            }
            return result;
        }
    }
}