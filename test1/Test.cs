using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using Prism.Commands;
using Prism.Mvvm;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace YourNamespace     
{
    public class WMXIOMonitoringViewModel : DialogViewModelBase
    {
        private readonly IOConfigFileService _fileService;
        private readonly IWMXIOMonitoringService _monService;
        private readonly IMsgBoxService _msgBoxService;
        private readonly ISerialPacketMonitoringService _spms;
        private readonly IRecordService _recordService;

        private readonly IEventAggregator _ea;
        private readonly IDialogService _ds;
        private readonly IRegionManager _rm;

        private readonly string _serialBasePath;
        private bool _spmsInitialized;

        private readonly int _uiBudgetPerTick = 400;
        private List<string> _rrKeys = new();
        private int _rrIndex = 0;
        private volatile Dictionary<string, double> _lastValueMap = new();
        private int _fastUiApplyScheduled = 0;

        private readonly Dictionary<(string unit, Cat cat), HashSet<string>> _selectedKeysPerUnit = new();
        private readonly Dictionary<string, PacketMonitorPaneViewModel> _packetPaneIndex = new();
        private bool _suppressLeftSelection;
        private bool _suppressTrendDetach;
        private readonly Timer _pollTimer;

        public enum EndpointValueKind { Digital, Analog }

        private enum Cat { None, Digital, Analog, Linked, SetAnalog, SerialPacket }
        private Cat _selectedCat = Cat.None;
        private bool _isDigitalCategory, _isAnalogCategory, _isLinkedDigitalCategory, _isSetAnalogCategory, _isSerialPacketCategory = false;

        public bool IsDigitalCategory
        {
            get => _isDigitalCategory;
            set
            {
                if (SetProperty(ref _isDigitalCategory, value))
                {
                    if (value)
                    {
                        _selectedCat = Cat.Digital;
                        RebuildCurrentUnitCategoryView();
                    }
                    else if (_selectedCat == Cat.Digital)
                    {
                        _selectedCat = Cat.None;
                        RebuildCurrentUnitCategoryView();
                    }
                }
            }
        }
        public bool IsAnalogCategory
        {
            get => _isAnalogCategory;
            set
            {
                if (SetProperty(ref _isAnalogCategory, value))
                {
                    if (value)
                    {
                        _selectedCat = Cat.Analog;
                        RebuildCurrentUnitCategoryView();
                    }
                    else if (_selectedCat == Cat.Analog)
                    {
                        _selectedCat = Cat.None;
                        RebuildCurrentUnitCategoryView();
                    }
                }
            }
        }
        public bool IsLinkedDigitalCategory
        {
            get => _isLinkedDigitalCategory;
            set
            {
                if (SetProperty(ref _isLinkedDigitalCategory, value))
                {
                    if (value)
                    {
                        _selectedCat = Cat.LinkedDigital;
                        RebuildCurrentUnitCategoryView();
                    }
                    else if (_selectedCat == Cat.LinkedDigital)
                    {
                        _selectedCat = Cat.None;
                        RebuildCurrentUnitCategoryView();
                    }
                }
            }
        }
        public bool IsSetAnalogCategory
        {
            get => _isSetAnalogCategory;
            set
            {
                if (SetProperty(ref _isSetAnalogCategory, value))
                {
                    if (value)
                    {
                        _selectedCat = Cat.SetAnalog;
                        RebuildCurrentUnitCategoryView();
                    }
                    else if (_selectedCat == Cat.SetAnalog)
                    {
                        _selectedCat = Cat.None;
                        RebuildCurrentUnitCategoryView();
                    }
                }
            }
        }

        public bool IsSerialPacketCategory
        {
            get => _isSerialPacketCategory;
            set
            {
                if (SetProperty(ref _isSerialPacketCategory, value))
                {
                    if (value)
                    {
                        _selectedCat = Cat.SerialPacket;
                        RebuildCurrentUnitCategoryView();
                    }
                    else if (_selectedCat == Cat.SerialPacket)
                    {
                        _selectedCat = Cat.None;
                        RebuildCurrentUnitCategoryView();
                    }
                }
            }
        }

        private string _selectedUnit;
        public string SelectedUnit
        {
            get => _selectedUnit;
            set
            {
                if (SetProperty(ref _selectedUnit, value))
                {
                    ClearCategorySelection(); // 라디오 해제 + 좌측 리스트 비움
                    CurrentUnitCategoryItems.Clear();
                    _currentCategoryAllChecked = false;
                    RaisePropertyChanged(nameof(CurrentCategoryAllChecked));
                    RaisePropertyChanged(nameof(CurrentCategoryTitle));
                }
            }
        }

        public string CurrentCategoryTitle => _selectedCat == Cat.None ? "Select Category" : _selectedCat.ToString();

        private bool _currentCategoryAllChecked;
        public bool CurrentCategoryAllChecked
        {
            get => _currentCategoryAllChecked;
            set
            {
                if (SetProperty(ref _currentCategoryAllChecked, value))
                {
                   _ = HandleMasterSelectionChangedAsync(value);
                }
            }
        }

        private bool _isRecordEnabled;
        public bool _isRecordEnabled
        {
            get => _isRecordEnabled;
            set
            {
                if (SetProperty(ref _isRecordEnabled, value))
                {
                    if (_isRecordEnabled)
                    {
                        _recordService.SetRecording(true);
                    }
                    else
                    {
                        _recordService.SetRecording(false);
                    }
                }
            }
        }

        private readonly Dictionary<string, SerialPacketRowVM> _serialRowIndex = new();
        private int _selectedTabIndex;
        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set => SetProperty(ref _selectedTabIndex, value);
        }

        public ObservableCollection<UnitScopeItemVM> CurrentUnitCategoryItems { get; set;} = new();

        public ObservableCollection<IOConfigItem> DigitalIO { get; } = new();
        public ObservableCollection<IOConfigItem> AnalogIO { get; } = new();
        public ObservableCollection<LinkedDigitalIO> LinkedItems { get; } = new();
        public ObservableCollection<SetAnalogIO> SetAnalogItems { get; } = new();

        public ObservableCollection<string> Units { get; } = new();

        public ListCollectionView SelectedDigitalItemsView { get; }
        public ListCollectionView SelectedAnalogItemsView { get; }
        public ListCollectionView SelectedLinkedItemsView { get; }
        public ListCollectionView SelectedSetAnalogItemsView { get; }
        public ListCollectionView SelectedSerialPacketItemsView { get; }

        public ObservableCollection<DigitalRowVM> SelectedDigitalItems { get; } = new();
        public ObservableCollection<AnalogRowVM> SelectedAnalogItems { get; } = new();
        public ObservableRangeCollection<LinkedEndpointRowVM> SelectedLinkedItems { get; } = new();
        public ObservableCollection<SetAnalogEndpointRowVM> SelectedSetAnalogItems { get; } = new();
        public ObservableCollection<SerialPacketRowVM> SelectedSerialPacketItems { get; } = new();

        public ObservableCollection<TrendPaneViewModel> TrendPanes { get; } = new();
        public ObservableCollection<PacketMonitorPaneViewModel> PacketMonitorPanes {get;} = new();

        public DelegateCommand AddTrendFromDigitalCommand { get; }
        public DelegateCommand AddTrendFromAnalogCommand { get; }
        public DelegateCommand AddTrendFromLinkedCommand { get; }
        public DelegateCommand AddTrendFromSetAnalogCommand { get; }
        public DelegateCommand<string> CloseTrendCommand { get; }
        public DelegateCommand<string> ClosePacketMonitorCommand {get;}
        public DelegateCommand<UnitScopeItemVM> ToggleLeftSelectionCommand { get; }

        public DelegateCommand ApplyDigitalChangedCommand {get;}
        public DelegateCommand ApplyAnalogChangeCommand {get;}
        public DelegateCommand ApplyLinkedChangedCommand {get;}
        public DelegateCommand ApplySetAnalogChangedCommand {get;}
        public DelegateCommand AddMonitorFromSerialCommand {get;}


        public WMXIOMonitoringViewModel(IOConfigFileService fileService, IEventAggregator ea, IDialogService ds, IRegionManager rm, IWMXIOMonitoringService monService, IMsgBoxService msgService, ISerialPacketMonitoringService spms, IRecordService _recordService) : base(rm, ea)
        {
            _ea = ea;
            _ds = ds;
            _rm = rm;

            _fileService = fileService;
            _monService = monService;
            _msgBoxService = msgService;
            _spms = spms;
            _recordService = _recordService;

            // 초기 로드
            LoadAll();

            // Units
            foreach (var u in EnumerateUnitsFromAll()) 
                Units.Add(u);

            // 라디오 초기 상태: 아무것도 선택되지 않음
            _selectedCat = Cat.None;
            
            SelectedUnit = Units.FirstOrDefault();

            // View(Grouping)
            SelectedDigitalItemsView = (ListCollectionView)CollectionViewSource.GetDefaultView(SelectedDigitalItems);
            SelectedDigitalItemsView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(DigitalRowVM.UnitName)));
            SelectedAnalogItemsView = (ListCollectionView)CollectionViewSource.GetDefaultView(SelectedAnalogItems);
            SelectedAnalogItemsView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(AnalogRowVM.UnitName)));
            SelectedLinkedItemsView = (ListCollectionView)CollectionViewSource.GetDefaultView(SelectedLinkedItems);
            SelectedLinkedItemsView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(LinkedEndpointRowVM.UnitName)));
            SelectedSetAnalogItemsView = (ListCollectionView)CollectionViewSource.GetDefaultView(SelectedSetAnalogItems);
            SelectedSetAnalogItemsView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(SetAnalogEndpointRowVM.UnitName)));
            SelectedSerialPacketItemsView = (ListCollectionView)CollectionViewSource.GetDefaultView(SelectedSerialPacketItems);
            SelectedSerialPacketItemsView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(SerialPacketRowVM.UnitName)));
            // Commands
            AddTrendFromDigitalCommand = new DelegateCommand(() => AddTrendFrom(SelectedDigitalItems.Where(x => x.TrendChecked), true));
            AddTrendFromAnalogCommand  = new DelegateCommand(() => AddTrendFrom(SelectedAnalogItems.Where(x => x.TrendChecked), false));
            AddTrendFromLinkedCommand  = new DelegateCommand(() => AddTrendFrom(SelectedLinkedItems.Where(x => x.TrendChecked), true));
            AddTrendFromSetAnalogCommand = new DelegateCommand(() => AddTrendFrom(SelectedSetAnalogItems.Where(x => x.TrendChecked), false));

            CloseTrendCommand = new DelegateCommand<string>(async title => { await CloseTrendAsync(title); });
            
            ToggleLeftSelectionCommand = new DelegateCommand<UnitScopeItemVM>(OnToggleLeftSelectionAsync);

            ApplyDigitalChangeCommand = new DelegateCommand(async () => await ApplyChangesDigitalAsync());
            ApplyAnalogChangeCommand = new DelegateCommand(async () => await ApplyChangesAnalogAsync());
            ApplyLinkedChangeCommand = new DelegateCommand(async () => await ApplyChangesLinkedAsync());
            ApplySetAnalogChangeCommand = new DelegateCommand(async () => await ApplyChangesSetAnalogAsync());

            AddMonitorFromSerialCommand = new DelegateCommand(OnAddMonitorFromSerial);
            ClosePacketMonitorCommand = new DelegateCommand<string>(async id => await ClosePacketMonitorAsync(id));

            _spms.LineArrived += Spms_LineArrived;
            _spms.LinesPublished += Spms_LinesPublished;
            _serialBasePath = ConfigManager.Instance.ConfigPath;

            _recordService.Initialize(ConfigManager.Instance.Solutionpath);
            _pollTimer = new Timer(async _ => await PollAsync().ConfigureAwait(false), null, 200, 500);
        }

        private void LoadAll()
        {
            DigitalIO.Clear();
            AnalogIO.Clear();
            LinkedItems.Clear();
            SetAnalogItems.Clear();

            var (dig, ana) = _fileService.LoadDA();

            foreach (var d in dig) DigitalIO.Add(d);
            foreach (var a in ana) AnalogIO.Add(a);

            // SettingSystemIO.cs 들 -> LinkedDigital, SetAnalog
            foreach (var l in _fileService.LoadDigitalSettingsByScript())
                LinkedItems.Add(l);
            foreach (var sa in _fileService.LoadAnalogSettingsByScript(DigitalIO, AnalogIO))
                SetAnalogItems.Add(sa);
        }

        private IEnumerable<string> EnumerateUnitsFromAll()
        {
            var s1 = DigitalIO.Select(x => x.UnitName);
            var s2 = AnalogIO.Select(x => x.UnitName);
            var s3 = LinkedItems.Select(x => x.UnitName);
            var s4 = SetAnalogItems.Select(x => x.UnitName);
            return s1.Concat(s2).Concat(s3).Concat(s4)
                     .Where(s => !string.IsNullOrWhiteSpace(s))
                     .Distinct()
                     .OrderBy(s => s);
        }

        private void EnsureSpmsInitialized()
        {
            if (_spmsInitialized) return;

            _spms.Initialize(_serialBasePath);
            _spmsInitialized = true;
        }

        // ===== Left Row2 채우기 =====
        private void RebuildCurrentUnitCategoryView()
        {
            CurrentUnitCategoryItems.Clear();
            _currentCategoryAllChecked = false;
            RaisePropertyChanged(nameof(CurrentCategoryAllChecked));
            RaisePropertyChanged(nameof(CurrentCategoryTitle));

            if (_selectedCat == Cat.None || string.IsNullOrWhiteSpace(_selectedUnit)) 
                return;

            switch (_selectedCat)
            {
                case Cat.Digital:
                    foreach (var it in DigitalIO.Where(x => x.UnitName == _selectedUnit))
                        CurrentUnitCategoryItems.Add(UnitScopeItemVM.From(it));
                    break;
                case Cat.Analog:
                    foreach (var it in AnalogIO.Where(x => x.UnitName == _selectedUnit))
                        CurrentUnitCategoryItems.Add(UnitScopeItemVM.From(it));
                    break;
                case Cat.Linked:
                    foreach (var it in LinkedItems.Where(x => x.UnitName == _selectedUnit))
                        CurrentUnitCategoryItems.Add(UnitScopeItemVM.From(it));
                    break;
                case Cat.SetAnalog:
                    foreach (var it in SetAnalogItems.Where(x => x.UnitName == _selectedUnit))
                        CurrentUnitCategoryItems.Add(UnitScopeItemVM.From(it));
                    break;
                case Cat.SerialPacket:
                    {
                        EnsureSpmsInitialized();
                        var unit = SelectedUnit;
                        if (string.IsNullOrWhiteSpace(unit))
                        {
                            CurrentUnitCategoryItems.Clear();
                            return;
                        }
                        var eps = _spms.ListEndpointsByUnit(unit, includeRuntime: true).ToList();

                        CurrentUnitCategoryItems.Clear();
                        foreach(var ep in eps)
                        {
                            var item = UnitScopeItemVM.From(ep);
                            item.IsSelected = _serialRowIndex.ContainsKey(item.key);
                            CurrentUnitCategoryItems.Add(item);
                        }

                        RaisePropertyChanged(nameof(CurrentCategoryTitle));
                        RaisePropertyChanged(nameof(CurrentCategoryAllChecked));
                        break;
                    }
            }

            // 저장된 선택 상태 복원
            if (_selectedKeysPerUnit.TryGetValue((_selectedUnit, _selectedCat), out var saved))
            {
                foreach (var vm in CurrentUnitCategoryItems)
                    vm.IsSelected = saved.Contains(vm.Key);
            }
            UpdateMasterCheck();
            RaisePropertyChanged(nameof(CurrentCategoryTitle));
        }

        private void UpdateMasterCheck()
        {
            _currentCategoryAllChecked = CurrentUnitCategoryItems.Any() && CurrentUnitCategoryItems.All(x => x.IsSelected);
            RaisePropertyChanged(nameof(CurrentCategoryAllChecked));
        }

        private void SyncSelectionToTabs(Cat cat, string unit, bool selectAll)
        {
            foreach (var vm in CurrentUnitCategoryItems)
                SyncOneItem(cat, unit, vm, selectAll);

            SetActiveTabFor(cat);
            UpdateMasterCheck();

            TriggerImmediateUiApply();
        }

        private void SetActiveTabFor(Cat cat)
        {
            SelectedTabIndex = cat switch
            {
                Cat.Digital   => 0,
                Cat.Analog    => 1,
                Cat.Linked    => 2,
                Cat.SetAnalog => 3,
                _ => SelectedTabIndex
            };
        }

        private void SyncOneItem(Cat cat, string unit, UnitScopeItemVM vm, bool isSelected)
        {
            var key = (unit, cat);
            if (!_selectedKeysPerUnit.TryGetValue(key, out var set))
                _selectedKeysPerUnit[key] = set = new HashSet<string>();

            if (isSelected) 
                set.Add(vm.Key);
            else 
                set.Remove(vm.Key);

            switch (cat)
            {
                case Cat.Digital:
                {
                    if (isSelected)
                    {
                        if (SelectedDigitalItems.All(x => x.Key != vm.Key))
                        {
                            var row = DigitalRowVM.From(DigitalIO.First(d => d.UnitName == unit && d.Variable == vm.Variable));
                            HookRowSelectionSync(row);
                            row.IsSelected = true;
                            SelectedDigitalItems.Add(row);
                        }
                        TrySetLeftSelection(vm.Key, true);
                    }
                    else
                    {
                        var row = SelectedDigitalItems.FirstOrDefault(x => x.Key == vm.Key);
                        if (row != null)
                        {
                            row.IsSelected = false;
                            SelectedDigitalItems.Remove(row);
                        }
                        TrySetLeftSelection(vm.Key, false);
                    }
                    SetActiveTabFor(Cat.Digital);
                    break;
                }

                case Cat.Analog:
                {
                    if (isSelected)
                    {
                        if (SelectedAnalogItems.All(x => x.Key != vm.Key))
                        {
                            var row = AnalogRowVM.From(AnalogIO.First(a => a.UnitName == unit && a.Variable == vm.Variable));
                            HookRowSelectionSync(row);
                            row.IsSelected = true;
                            SelectedAnalogItems.Add(row);
                        }
                        TrySetLeftSelection(vm.Key, true);
                    }
                    else
                    {
                        var row = SelectedAnalogItems.FirstOrDefault(x => x.Key == vm.Key);
                        if (row != null)
                        {
                            row.IsSelected = false;
                            SelectedAnalogItems.Remove(row);
                        }
                        TrySetLeftSelection(vm.Key, false);
                    }
                    SetActiveTabFor(Cat.Analog);
                    break;
                }

                case Cat.Linked:
                {
                    var link = LinkedItems.First(l => l.UnitName == unit && (l.OutputItem.Variable + "|" + l.InputItem.Variable == vm.Variable));

                    var toAdd = new List<LinkedEndpointRowVM>();

                    var outKey = unit + "|" + link.OutputItem.Variable;
                    if (isSelected && SelectedLinkedItems.All(x => x.Key != outKey))
                    {
                        var outRow = LinkedEndpointRowVM.From(link, isOut: true);
                        HookRowSelectionSync(outRow);
                        outRow.IsSelected = true;
                        toAdd.Add(outRow);
                    }

                    var inKey = unit + "|" + link.InputItem.Variable;
                    if (isSelected && SelectedLinkedItems.All(x => x.Key != inKey))
                    {
                        var inRow = LinkedEndpointRowVM.From(link, isOut: false);
                        HookRowSelectionSync(inRow);
                        inRow.IsSelected = true;
                        toAdd.Add(inRow);
                    }

                    OnUi(() =>
                    {
                        if (toAdd.Count > 0) 
                            SelectedLinkedItems.AddRange(toAdd);

                        if (!isSelected)
                        {
                            SelectedLinkedItems.RemoveWhere(ep => 
                            ep.UnitName == unit &&
                            (ep.Variable == link.OutputItem.Variable || ep.Variable == link.InputItem.Variable));
                        }
                    });

                    TrySetLeftSelection(vm.Key, isSelected);
                    SetActiveTabFor(Cat.Linked);
                    break;
                }

                case Cat.SetAnalog:
                {
                    var sa = SetAnalogItems.First(s => s.UnitName == unit && s.AnalogIO.Variable == vm.Variable);

                    if (isSelected)
                    {
                        var anaKey = unit + "|" + sa.AnalogIO.Variable;
                        if (SelectedSetAnalogItems.All(x => x.Key != anaKey))
                        {
                            var anaRow = SetAnalogEndpointRowVM.FromAnalog(sa);
                            HookRowSelectionSync(anaRow);
                            anaRow.IsSelected = true;
                            SelectedSetAnalogItems.Add(anaRow);
                        }

                        if (!string.IsNullOrWhiteSpace(sa.OutSignalName))
                        {
                            var outIo = FindIOByVariable(unit, sa.OutSignalName);
                            if (outIo != null)
                            {
                                var outKey = unit + "|" + sa.OutSignalName;
                                if (SelectedSetAnalogItems.All(x => x.Key != outKey))
                                {
                                    var outRow = SetAnalogEndpointRowVM.FromOutSignal(sa, outIo);
                                    HookRowSelectionSync(outRow);
                                    outRow.IsSelected = true;
                                    SelectedSetAnalogItems.Add(outRow);
                                }
                            }
                        }

                        TrySetLeftSelection(vm.Key, true);
                    }
                    else
                    {
                        var anaToRemove = SelectedSetAnalogItems
                        .FirstOrDefault(x => x.Key == unit + "|" + sa.AnalogIO.Variable);
                        if (anaToRemove != null)
                        {
                            anaToRemove.IsSelected = false;
                            SelectedSetAnalogItems.Remove(anaToRemove);
                        }
                        if (!string.IsNullOrWhiteSpace(sa.OutSignalName))
                        {
                            var outToRemove = SelectedSetAnalogItems
                            .FirstOrDefault(x => x.Key == unit + "|" + sa.OutSignalName);
                            if (outToRemove != null)
                            {
                                outToRemove.IsSelected = false;
                                SelectedSetAnalogItems.Remove(outToRemove);
                            }
                        }
                        TrySetLeftSelection(vm.Key, false);
                    }

                    SetActiveTabFor(Cat.SetAnalog);
                    break;
                }

                case Cat.SerialPacket:
                    {
                        var serialKey = vm.Key;

                        if (isSelected)
                        {
                            if (_serialRowIndex.ContainsKey(serialKey)) break;

                            var parts = (vm.Variable ?? "").Split('|');
                            var device = parts.Length >= 1 ? parts[0] : "";
                            var port = parts.Length >= 2 ? parts[1] : "";

                            var epKey = new SerialPacketEndpointKey(vm.UnitName, device, port);
                            var row = new SerialPacketRowVM(epKey);
                            HookRowSelectionSync(row);
                            row.IsSelected = true;

                            _serialRowIndex[serialKey] = row;
                            Application.Current.Dispatcher.Invoke(
                            () => SelectedSerialPacketItems.Add(row),
                            System.Windows.Threading.DispatcherPriority.DataBind);
                        }
                        else
                        {
                            if (_serialRowIndex.TryGetValue(serialKey, out var row))
                            {
                                row.IsSelected = false;
                                Application.Current.Dispatcher.Invoke(
                                () => SelectedSerialPacketItems.Remove(row),
                                System.Windows.Threading.DispatcherPriority.DataBind);
                                _serialRowIndex.Remove(serialKey);
                                )
                            }
                        }
                        SetActiveTabFor(Cat.SerialPacket);
                        break;
                    }
            }

            UpdateMasterCheck();
            TriggerImmediateUiApply();
        }

        private void OnToggleLeftSelectionAsync(UnitScopeItemVM vm)
        {
            if (_suppressLeftSelection) return;

            // 체크 해제 시 Trend 포함 여부 확인
            if (!vm.IsSelected && _selectedCat != Cat.SerialPacket)
            {
                var seriesIds = BuildSeriesIdsForLeftItem(vm);
                var used = TrendPanes.Any(p => p.ContainsAnySeries(seriesIds));

                if (used)
                {
                    var r = MessageBox.Show(
                    "선택 해제하려는 항목이 현재 Trend 차트에 표시 중입니다.￦n이 항목을 Trend에서 제거하시겠습니까?",
                    "Trend 항목 제거 확인",
                    MessageBoxButton.YesNo, 
                    MessageBoxImage.Question);

                    if (r != MessageBoxResult.Yes)
                    {
                        // 취소: 체크 되돌리고 아무것도 제거하지 않음
                        _suppressLeftSelection = true;
                        _suppressTrendDetach = true;
                        vm.IsSelected = true;
                        _suppressTrendDetach = false;
                        _suppressLeftSelection = false;
                        return;
                    }

                    RemoveItemFromAllTrends(seriesIds);
                }
            }

            if (!vm.IsSelected && _selectedCat == Cat.SerialPacket)
            {
                var key = vm.Key;
                if (_serialRowIndex.TryGetValue(key, out var row))
                {
                    if (row.MonitorChecked)
                    {
                        var ok = MessageBox.Show($"'{row.DeviceName}|{row.PropertyName}' 모니터도 함께 종료할까요?", "종료 알림", MessageBoxButton.YesNo);
                        if (ok == MessageBoxResult.Yes)
                        {
                            row.MonitorChecked = false;
                            _spms.StopMonitoring(row.EndpointKey);
                            ClosePacketMonitorPaneByKey(row.Key);
                        }
                    }
                    SelectedSerialPacketItems.Remove(row);
                    _serialRowIndex.Remove(key);
                }
                return;
            }

            if (_selectedCat == Cat.SerialPacket)
            {
                var parts = (vm.Variable ?? "").Split('|');
                var device = parts.Length > 0 ? parts[0].Trim() : string.Empty;
                var port = parts.Length > 1 ? parts[1].Trim() : string.Empty;

                var epKey = new SerialPacketEndpointKey(vm.UnitName, device, port);
                var sKey = epKey.ToString();

                if (!_serialRowIndex.ContainsKey(skey))
                {
                    var row = new SerialPacketRowVM(epKey);
                    row.InitSnapshot(_spms);

                    row.PropertyChanged += (s,e) =>
                    {
                        if (e.PropertyName == nameof(SerialPacketRowVM.MonitorChecked))
                        {
                            if (row.MonitorChecked)
                            {
                                _spms.StartMonitoring(row.EndpointKey);
                                OpenOrFocusPacketMonitorPane(row.Key, row.EndpointKey);
                            }
                            else
                            {
                                _spms.StopMonitoring(row.EndpointKey);
                                ClosePacketMonitorPanelIfExists(row.Key);
                            }
                        }
                    };

                    _serialRowIndex[sKey] = row;
                    SelectedSerialPacketItems.Add(row);
                }
                SetActiveTabFor(Cat.SerialPacket);
                UpdateMasterCheck();
                return;
            }

            // 최종 동기화
            SyncOneItem(_selectedCat, _selectedUnit, vm, vm.IsSelected);
            UpdateMasterCheck();
        }

        private void Spms_LineArrived(SerialPacketEndpointKey key, PacketLine line)
        {
            var sKey = key.ToString();
            if (_serialRowIndex.TryGetValue(skey, out var row))
            {
                if (line.IsRecv)
                {
                    if (!string.Equals(row.LastRecvHex, line.Hex, StringComparison.Ordinal))
                    {
                        row.LastRecvHex = line.Hex;
                        row.LastRecvAt = line.Timestamp;
                    }
                }
                else
                {
                    if (!string.Equals(row.LastSendHex, line.Hex, StringComparison.Ordinal))
                    {
                        row.LastSendHex = line.Hex;
                        row.LastSendAt = line.Timestamp;
                    }
                }
            }
        }

        private void Spms_LinesPublished(SerialPacketEndpointKey key, IReadOnlyList<PacketLine> batch)
        {
            var skey = key.ToString();
            AppendPacketLinesToPanelIfWatching(skey, batch);
        }

        private void OnAddMonitorFromSerial()
        {
            if (_selectedCat != Cat.SerialPacket && SelectedTabIndex != 4)
            {
                return;
            }

            var targets = SelectedSerialPacketItems.Where(r => r.MonitorChecked).ToList();
            if (targets.Count == 0) return;

            foreach(var r in targets)
            {
                var id = r.Key;
                if (_packetPaneIndex.TryGetValue(id, out var pane))
                {
                    continue;
                }

                var epKey = new SerialPacketEndpointKey(r.UnitName, r.DeviceName, r.PortName);
                var p = new PacketMonitorPaneViewModel(_spms, epKey, _recordService);
                _packetPaneIndex[id] = p;
                PacketMonitorPanes.Add(p);

                p.Start(TimeSpan.FromMilliseconds(500));
                r.MonitorChecked = false;
            }
        }

        private void OpenOrFocusPacketMonitorPane(string id, SerialPacketEndpointKey endpointKey)
        {
            if (_packetPaneIndex.TryGetValue(id, out var existing))
            {
                Application.Current?.Dispatcher?.InvokeAsync(() =>
                    {
                        if (PacketMonitorPanes.Contains(existing))
                        {
                            PacketMonitorPanes.Remove(existing);
                            PacketMonitorPanes.Add(existing);
                        }
                    }, System.Windows.Threading.DispatcherPriority.Background);
                    return;
            }

            var pane = new PacketMonitorPaneViewModel(_spms, endpointKey, _recordService);
            _packetPaneIndex[id] = pane;

            Application.Current?.Dispatcher?.InvokeAsync(() =>
            {
                PacketMonitorPanes.Add(pane);
            }, System.Windows.Threading.DispatcherPriority.Background);

            pane.Start(TimeSpan.FromMilliseconds(500));
        }

        private void ClosePacketMonitorPanelIfExists(string id)
        {
            if (!_packetPaneIndex.TryGetValue(id, out var pane))
                return;
            
            _packetPaneIndex.Remove(id);

            Application.Current?.Dispatcher?.InvokeAsync(() =>
            {
                if (PacketMonitorPanes.Contains(pane))
                    PacketMonitorPanes.Remove(pane);
                
                try
                {
                    pane.Stop();
                }
                catch { }
            }, System.Windows.Threading.DispatcherPriority.Background);
        }

        private void AppendPacketLinesToPanelIfWatching(string id, IReadOnlyList<PacketLine> batch)
        {
            if (batch == null || batch.Count == 0)
                return;
            
            if (!_packetPaneIndex.TryGetValue(id, out var pane))
                return;

            Application.Current?.Dispatcher?.InvokeAsync(() =>
            {
                pane.AppendBatch(batch);
            }, System.Windows.Threading.DispatcherPriority.Background);
        }

        private void async Task ClosePacketMonitorAsync(string id)
        {
            var pane = PacketMonitorPanes.FirstOrDefault(p => p.EndpointKey.ToString() == id);
            if (pane != null)
            {
                pane.Stop();
                await Application.Current.Dispatcher.InvokeAsync(
                    () => PacketMonitorPanes.Remove(pane),
                    System.Windows.Threading.DispatcherPriority.Background);

                if (_serialRowIndex.TryGetValue(id, out var row))
                {
                    row.MonitorChecked = false;
                }
            }
        }

        private async Task ApplyChangesDigitalAsync()
        {
            var targets = SelectedDigitalItems.Where(r => !string.IsNullOrWhiteSpace(r.ChangeValue)).ToList();

            if (!targets.Any()) 
                return;

            var errors = new List<string>();
            var dict = new Dictionary<string, double>(); // key = Unit|Var

            foreach (var r in targets)
            {
                if (!TryParseDigital(r.ChangeValue, out var val))
                {
                    errors.Add($"[Digital] {r.UnitName}.{r.Variable} : 0 또는 1만 입력하세요 (입력: '{r.ChangeValue}')");
                    continue;
                }
                dict[ReadSpec.FromIO(r.Model)] = val;
            }

            if (errors.Any())
            {
                _msgBoxService.Warn(string.Join(Environment.NewLine, errors));
                return;
            }

            await _monService.ApplySamplesAsync(dict);

            foreach (var r in targets) 
                r.ChangeValue = string.Empty;
        }

        // [NEW] Analog 탭: (실수)
        private async Task ApplyChangesAnalogAsync()
        {
            var targets = SelectedAnalogItems.Where(r => !string.IsNullOrWhiteSpace(r.ChangeValue)).ToList();

            if (!targets.Any()) 
                return;

            var errors = new List<string>();
            var dict = new Dictionary<string, double>();

            foreach (var r in targets)
            {
                if (!TryParseAnalog(r.ChangeValue, out var val))
                {
                    errors.Add($"[Analog] {r.UnitName}.{r.Variable} : 실수 형식으로 입력하세요 (예: 12.3) (입력: '{r.ChangeValue}')");
                    continue;
                }
                dict[ReadSpec.FromIO(r.Model)] = val;
            }

            if (errors.Any())
            {
                _msgBoxService.Warn(string.Join(Environment.NewLine, errors));
                return;
            }

            await _monService.ApplySamplesAsync(dict);

            foreach (var r in targets) 
                r.ChangeValue = string.Empty;
        }

        // [NEW] Linked 탭: (Digital로 간주)
        private async Task ApplyChangesLinkedAsync()
        {
            var targets = SelectedLinkedItems.Where(r => !string.IsNullOrWhiteSpace(r.ChangeValue)).ToList();

            if (!targets.Any()) 
                return;

            var errors = new List<string>();
            var dict = new Dictionary<string, double>();

            foreach (var r in targets)
            {
                if (!TryParseDigital(r.ChangeValue, out var val))
                {
                    errors.Add($"[Linked] {r.UnitName}.{r.Variable} : 0 또는 1만 입력하세요 (입력: '{r.ChangeValue}')");
                    continue;
                }
                dict[ReadSpec.FromIO(r.Model)] = val;
            }

            if (errors.Any())
            {
                _msgBoxService.Warn(string.Join(Environment.NewLine, errors));
                return;
            }

            await _monService.ApplySamplesAsync(dict);

            foreach (var r in targets) 
                r.ChangeValue = string.Empty;
        }

        // [NEW] SetAnalog 탭: IsAnalog == true → 실수, false → 0/1
        private async Task ApplyChangesSetAnalogAsync()
        {
            var targets = SelectedSetAnalogItems.Where(r => !string.IsNullOrWhiteSpace(r.ChangeValue)).ToList();

            if (!targets.Any()) 
                return;

            var errors = new List<string>();
            var dict = new Dictionary<string, double>();

            foreach (var r in targets)
            {
                bool ok;
                double val;

                if (r.ValueKind == EndpointValueKind.Analog)
                {
                    ok = double.TryParse(r.ChangeValue?.Trim(), NumberStyle.Float, CultureInfo.InvariantCulture, out val);
                }
                else
                {
                    ok = (r.ChangeValue?.Trim() == "0" || r.ChangeValue?.Trim() == "1");
                    val = ok && r.ChangeValue!.Trim() == "1" ? 1.0 : 0.0;
                }

                if (!ok)
                {
                    var kind = r.ValueKind == EndpointValueKind.Analog ? "Analog(실수)" : "Digital(0/1)";
                    errors.Add($"[SetAnalog] {r.UnitName}.{r.Variable} : {kind} 형식으로 입력하세요 (입력: '{r.ChangeValue}')");
                    continue;
                }

                dict[ReadSpec.FromIO(r.Model)] = val;
            }

            if (errors.Any())
            {
                _msgBoxService.Warn(string.Join(Environment.NewLine, errors));
                return;
            }

            await _monService.ApplySamplesAsync(dict);

            foreach (var r in targets) 
                r.ChangeValue = string.Empty;
        }

        private List<string> BuildSeriesIdsForLeftItem(UnitScopeItemVM vm)
        {
            var unit = item.UnitName;
            var ids = new List<string>();

            switch (_selectedCat)
            {
                case Cat.Digital:
                    ids.Add($"{unit}:D:{item.Variable}");
                    break;
                case Cat.Analog:
                    ids.Add($"{unit}:A:{item.Variable}");
                    break;
                case Cat.Linked:
                    var tokens = (item.Variable ?? "").Split('|');
                    if (tokens.Length == 2)
                    {
                        ids.Add($"{unit}:LD:{tokens[0]}");
                        ids.Add($"{unit}:LD:{tokens[1]}");
                    }
                    break;

                case Cat.SetAnalog:
                    ids.Add($"{unit}:SA:{item.Variable}");
                    var sa = SetAnalogItems.FirstOrDefault(s => s.UnitName == unit && s.AnalogIO.Variable == item.Variable);
                    if (sa != null && !string.IsNullOrWhiteSpace(sa.OutSignalName))
                        ids.Add($"{unit}:SA:{sa.OutSignalName}");
                    break;
            }
            return ids;
        }

        private void RemoveItemFromAllTrends(IEnumerable<string> seriesIds)
        {
            var ids = new HashSet<string>(seriesIds ?? Array.Empty<string>());
            if (ids.Count == 0) return;

            var panes = TrendPanes.ToList();
            foreach(var pane in panes)
            {
                if (!pane.ContainsAnySeries(ids)) continue;

                pane.RemoveSeriesByIds(ids);

                if (pane.IsEmpty)
                {
                    _ = Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        pane.Stop();
                        TrendPanes.Remove(pane);
                    });
                }
            }
        }

        private void OnUi(Action action)
        {
            if (Application.Current?.Dispatcher?.CheckAccess() == true) 
                action();
            else Application.Current?.Dispatcher?.Invoke(action);
        }

        private void TrySetLeftSelection(string key, bool isSelected)
        {
            var item = CurrentUnitCategoryItems.FirstOrDefault(i => i.Key == key);
            if (item != null && item.IsSelected != isSelected)
            {
                _suppressLeftSelection = true;
                item.IsSelected = isSelected;
                _suppressLeftSelection = false;
            }
        }

        // 탭(RowVM)에서 체크 변경 → 좌측/사전 동기화 훅
        private void UpdateSelectionFromTab(Cat cat, string unit, string key, bool isSelected)
        {
            var tuple = (unit, cat);
            if (!_selectedKeysPerUnit.TryGetValue(tuple, out var set))
                _selectedKeysPerUnit[tuple] = set = new HashSet<string>();

            if (isSelected) 
                set.Add(key);
            else 
                set.Remove(key);

            // 현재 좌측 리스트에 같은 key가 보이는 경우 동기화
            TrySetLeftSelection(key, isSelected);
        }

        // 타입별 훅 등록기
        private void HookRowSelectionSync(DigitalRowVM r)
        {
            r.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(DigitalRowVM.IsSelected))
                    UpdateSelectionFromTab(Cat.Digital, r.UnitName, r.Key, r.IsSelected);
            };
        }
        private void HookRowSelectionSync(AnalogRowVM r)
        {
            r.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(AnalogRowVM.IsSelected))
                    UpdateSelectionFromTab(Cat.Analog, r.UnitName, r.Key, r.IsSelected);
            };
        }
        private void HookRowSelectionSync(LinkedEndpointRowVM r)
        {
            r.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(LinkedEndpointRowVM.IsSelected))
                    UpdateSelectionFromTab(Cat.Linked, r.UnitName, r.Key, r.IsSelected);
            };
        }
        private void HookRowSelectionSync(SetAnalogEndpointRowVM r)
        {
            r.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(SetAnalogEndpointRowVM.IsSelected))
                    UpdateSelectionFromTab(Cat.SetAnalog, r.UnitName, r.Key, r.IsSelected);
            };
        }

        private void HookRowSelectionSync(SerialPacketRowVM r)
        {
            r.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(SerialPacketRowVM.IsSelected))
                    UpdateSelectionFromTab(Cat.SerialPacket, r.UnitName, r.Key, r.IsSelected);
            };
        }


        // ===== Trend 추가/닫기/Export =====
        private async void AddTrendFrom(IEnumerable<ITrendSource> rows, bool isDigital)
        {
            var groups = GroupTrendTargets(rows)?.ToList();
            if (groups == null || groups.Count == 0) 
                return;

            if (rows.Count() >= 7)
            {
                _msgBoxService.Info("한 번에 6개 이하 항목만 트렌드로 추가할 수 있습니다.");
                return;
            }

            foreach (var group in groups)
            {
                var already = TrendPanes.Any(p =>
                    p.Title == group.Title &&
                    p.Series.Count == group.SeriesTargets.Count &&
                    group.SeriesTargets.All(t => p.Series.Any(s => s.Name == t.SeriesTitle)));

                if (already) continue;

                var vm = new TrendPaneViewModel(group.Title, group.SeriesTargets, isDigital);

                await Application.Current.Dispatcher.InvokeAsync(() => { TrendPanes.Add(vm); },
                    System.Windows.Threading.DispatcherPriority.Background);
                await Application.Current.Dispatcher.InvokeAsync(() => { vm.Start(); },
                    System.Windows.Threading.DispatcherPriority.Background);

                foreach (var r in rows) r.TrendChecked = false;
            }
        }

        private void ClosePacketMonitorPaneByKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) 
                return;
            
            var pane = PacketMonitorPanes.FirstOrDefault(p => p.Key == key);

            if (pane == null)
                return;

                pane.Stop();
                Application.Current.Dispatcher.Invoke(() =>
                {
                    PacketMonitorPanes.Remove(pane);
                }, DispatcherPriority.Background);
        }

        private bool IsVmInAnyTrend(UnitScopeItemVM vm)
        {
            var ids = BuildSeriesIdsForLeftItem(vm);
            return TrendPanes.Any(p => p.ContainsAnySeries(ids));
        }

        private void UnSelectList(IEnumerable<UnitScopeItemVM> vms)
        {
            var list = vms.ToList();

            if (list.Count == 0) 
                return;

            _suppressLeftSelection = true;

            foreach (var vm in list)
            {
                if (vm.IsSelected) vm.IsSelected = false;
                SyncOneItem(_selectedCat, _selectedUnit, vm, isSelected: false);    
            }
            _suppressLeftSelection = false;
            UpdateMasterCheck();
        }

        private Task HandleMasterSelectionChangedAsync(bool isChecked)
        {
            if (_selectedCat == Cat.None || string.IsNullOrWhiteSpace(_selectedUnit))
            {
                UpdateMasterCheck();
                return Task.CompletedTask;
            }

            if (isChecked)
            {
                _suppressLeftSelection = true;
                foreach(var vm in CurrentUnitCategoryItems)
                {
                    if (!vm.IsSelected) 
                        vm.IsSelected = true;
                }
                _suppressLeftSelection = false;

                SyncSelectionToTabs(_selectedCat, _selectedUnit, selectAll: true);
                SetActiveTabFor(_selectedCat);
                UpdateMasterCheck();
                return Task.CompletedTask;
            }

            var currentlySelected = CurrentUnitCategoryItems.Where(x => x.IsSelected).ToList();
            if (currentlySelected.Count == 0)
            {
                UpdateMasterCheck();
                return Task.CompletedTask;
            }

            var trending = currentlySelected.Where(IsVmInAnyTrend).ToList();
            var nonTrending = currentlySelected.Except(trending).ToList();

            if (trending.Count == 0)
            {
                UnselectList(currentlySelected);
                return Task.CompletedTask;
            }

            var msg =
            $"현재 트렌드에 포함된 항목이 {trending.Count}개 있습니다. \n" +
            $"모두 트렌드에서 제거하고 전체 해제하시겠습니까? \n\n" +
            $"[예] 트렌드에서 제거 후 전부 해제\n" +
            $"[아니오] 트렌드 항목은 유지, 트렌드 미포함 항목만 해제";
            var result = MessageBox.Show(msg, "전체 해제  확인", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                var seriesIdsToRemove = trending.SelectMany(BuildSeriesIdsForLeftItem).ToList();
                RemoveItemFromAllTrends(seriesIdsToRemove);
                UnselectList(currentlySelected);
            }
            else
            {
                UnselectList(nonTrending);
            }
            return Task.CompletedTask;
        }

        private static string IoKeyFromSeriesId(string seriesId)
        {
            if (string.IsNullOrWhiteSpace(seriesId)) 
                return null;

            var parts = seriesId.Split(':');

            if (parts.Length <3) 
                return null;

            var unit = parts[0];
            var variable = parts[2];
            return $"{unit}|{variable}";
        }

        private HashSet<string> BuildPriorityKeysSnapshot()
        {
            var pri = new HashSet<string>();

            foreach (var p in TrendPanes)
            {
                var ids = p.GetSeriesIdsSnapshot();
                
                if (ids == null) 
                    continue;

                foreach(var sid in ids)
                {
                    var k = IoKeyFromSeriesId(sid);
                    if (k != null) pri.Add(k);
                    }
                }
            
            switch(SelectedTabIndex)
            {
                case 0:
                    foreach (var r in SelectedDigitalItems.Where(x => x.IsSelected))
                        pri.Add(r.Key);
                    break;
                case 1:
                    foreach (var r in SelectedAnalogItems.Where(x => x.IsSelected))
                        pri.Add(r.Key);
                    break;
                case 2:
                    foreach (var r in SelectedLinkedItesm.Where(x => x.IsSelected))
                        pri.Add(r.Key);
                    break;
                case 3:
                    foreach (var r in SelectedSetAnalogItems.Where(x => x.IsSelected))
                        pri.Add(r.Key);
                    break;
            }
            return pri;
        }

        private async Task CloseTrendAsync(string id)
        {
            var pane = TrendPanes.FirstOrDefault(p => p.Title == id);
            if (pane != null)
            {
                pane.Stop();
                await Application.Current.Dispatcher.InvokeAsync(() => { TrendPanes.Remove(pane); },
                    System.Windows.Threading.DispatcherPriority.Background);
            }
        }

        private void ClearCategorySelection()
        {
            _isDigitalCategory = _isAnalogCategory = _isLinkedDigitalCategory = _isSetAnalogCategory = false;
            RaisePropertyChanged(nameof(IsDigitalCategory));
            RaisePropertyChanged(nameof(IsAnalogCategory));
            RaisePropertyChanged(nameof(IsLinkedDigitalCategory));
            RaisePropertyChanged(nameof(IsSetAnalogCategory));
            _selectedCat = Cat.None;
            RebuildCurrentUnitCategoryView();
        }

        private static bool IsIoDigital(IOConfigItem io)
        {
            if (io == null) return false;

            var t = io.Type;
            if (t == IoType.Digital) return true;
            else return false;
        }
        private bool TryParseDigital(string s, out double v)
        {
            v = 0;
            if (s == null) return false;
            s = s.Trim();
            if (s == "0")
            {
                v = 0;
                return true;

            }   

            if (s == "1")
            {
                v = 1;
                return true;
            }
            return false;
        }

        private bool TryParseAnalog(string s, out double v)
        {
            return double.TryParse(s?.Trim(), NumberStyle.Float, CultureInfo.InvariantCulture, out v);
        }


        // 인덱스(빠른 갱신용): key = Unit|Variable
        private readonly Dictionary<string, DigitalRowVM> _digitalIndex = new();
        private readonly Dictionary<string, AnalogRowVM> _analogIndex = new();
        private readonly Dictionary<string, LinkedEndpointRowVM> _linkedIndex = new();
        private readonly Dictionary<string, SetAnalogEndpointRowVM> _setAnalogIndex = new();

        private IEnumerable<ReadSpec> BuildReadSpecs()
        {
            DigitalRowVM[] dItems = Array.Empty<DigitalRowVM>();
            AnalogRowVM[] aItems = Array.Empty<AnalogRowVM>();
            LinkedEndpointRowVM[] lItems = Array.Empty<LinkedEndpointRowVM>();
            SetAnalogEndpointRowVM[] sItems = Array.Empty<SetAnalogEndpointRowVM>();

            Application.Current.Dispatcher.Invoke(() =>
            {
                dItems = SelectedDigitalItems.ToArray();
                aItems = SelectedAnalogItems.ToArray();
                lItems = SelectedLinkedItems.ToArray();
                sItems = SelectedSetAnalogItems.ToArray();

                _digitalIndex.Clear();
                _analogIndex.Clear();
                _linkedIndex.Clear();
                _setAnalogIndex.Clear();

                foreach (var d in dItems) _digitalIndex[d.Key] = d;
                foreach (var a in aItems) _analogIndex[a.Key] = a;
                foreach (var l in lItems) _linkedIndex[l.Key] = l;
                foreach (var s in sItems) _setAnalogIndex[s.Key] = s;
            });

            var list = new List<ReadSpec>(dItems.Length + aItems.Length + lItems.Length + sItems.Length);

            foreach (var d in dItems) list.Add(ReadSpec.FromIO(d.Model));
            foreach (var a in aItems) list.Add(ReadSpec.FromIO(a.Model));
            foreach (var l in lItems) list.Add(ReadSpec.FromIO(l.Model));
            foreach (var s in sItems) list.Add(ReadSpec.FromIO(s.Model));

            return list.GroupBy(x => x.Key).Select(g => g.First());
        }

        // ===== 폴링 =====
        private async Task PollAsync()
        {
            var specs = BuildReadSpecs();
            if (!specs.Any()) return;

            var values = await _monService.ReadAsync(specs).ConfigureAwait(false);

            var approxCount = _digitalIndex.Count + _analogIndex.Count + _linkedIndex.Count + _setAnalogIndex.Count;
            var map = approxCount > 0 ? new Dictionary<string, double>(approxCount) : new Dictionary<string, double>();

            foreach (var v in values)
                map[v.Key] = v.Value;

            _lastValueMap = map;

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                var priKeys = BuildPriorityKeysSnapshot();
            
                IEnumerable<string> allKeys = _digitalIndex.Keys
                .Concat(_analogIndex.Keys)
                .Concat(_linkedIndex.Keys)
                .Concat(_setAnalogIndex.Keys)
                .Distinct()
                .ToList();

                foreach (var k in allKeys)
                {
                    if (!priKeys.Contains(k)) continue;
                    if (!_lastValueMap.TryGetValue(k, out var y)) continue;

                    if (_digitalIndex.TryGetValue(k, out var d)) d.LiveValue = y;
                    if (_analogIndex.TryGetValue(k, out var a)) a.LiveValue = y;
                    if (_linkedIndex.TryGetValue(k, out var le)) le.LiveValue = y;
                    if (_setAnalogIndex.TryGetValue(k, out var se)) se.LiveValue = y;
                }

                var nonPri = allKeys.Where(k => !priKeys.Contains(k)).ToList();
                if (_rrKeys.Count != nonPri.Count || !_rrKeys.SequenceEqual(nonPri))
                {
                    _rrKeys = nonPri;
                    _rrIndex = 0;
                }

                if (_rrKeys.Count > 0)
                {
                    var budget = _uiBudgetPerTick;
                    var n = _rrKeys.Count;
                    if (_rrIndex >= n) _rrIndex %= n;

                    int processed = 0;
                    int i = _rrIndex;
                    while (processed < budget && processed < n)
                    {
                        var k = _rrKeys[i];
                        if (_lastValueMap.TryGetValue(k, out var y))
                        {
                            if (_digitalIndex.TryGetValue(k, out var d)) d.LiveValue = y;
                            if (_analogIndex.TryGetValue(k, out var a)) a.LiveValue = y;
                            if (_linkedIndex.TryGetValue(k, out var le)) le.LiveValue = y;
                            if (_setAnalogIndex.TryGetValue(k, out var se)) se.LiveValue = y;
                        }

                        processed++;
                        i++;
                        if (i >= n) i = 0;
                    }
                    _rrIndex = i;
                }
            }, System.Windows.Threading.DispatcherPriority.DataBind);
        }

        private void TriggerImmediateUiApply()
        {
            if (System.Threading.Interlocked.Exchange(ref _uiApplyRequested, 1) == 1)
                return;
            
            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                try
                {
                    IEnumverable<string> allKeys = _digitalIndex.Keys
                    .Concat(_analogIndex.Keys)
                    .Concat(_linkedIndex.Keys)
                    .Concat(_setAnalogIndex.Keys)
                    .Distinct()
                    .ToList();

                    foreach(var k in allKeys)
                    {
                        if (!_lastValueMap.TryGetValue(k, out var y)) continue;

                        if (_digitalIndex.TryGetValue(k, out var d)) d.LiveValue = y;
                        if (_analogIndex.TryGetValue(k, out var a)) a.LiveValue = y;
                        if (_linkedIndex.TryGetValue(k, out var le)) le.LiveValue = y;
                        if (_setAnalogIndex.TryGetValue(k, out var se)) se.LiveValue = y;
                    }
                }
                finally
                {
                    _fastUiApplyScheduled = 0;
                }
            }, System.Windows.Threading.DispatcherPriority.DataBind);
        }

        private IOConfigItem FindIOByVariable(string unit, string variable)
        => DigitalIO.Concat(AnalogIO).FirstOrDefault(x => x.UnitName == unit && x.Variable == variable);

        private IEnumerable<TrendGroup> GroupTrendTargets(IEnumerable<ITrendSource> rows)
        {
            var items = rows.ToList();
            if (!items.Any()) yield break;

            var seriesTargets = new List<TrendSeriesTarget>();
            foreach (var item in items)
                seriesTargets.AddRange(r.AsTrendSeries());
            
            var title = string.Join(", ", seriesTargets.Select(s => s.SeriesTitle).Take(3));
            var id = string.Join(";", seriesTargets.Select(s => s.SeriesId).OrderBy(s => s));
            yield return new TrendGroup
            {
                Title = title, Id = id, SeriesTargets = seriesTargets };
        }

        public override bool CanCloseDialog()
        {
            var result = MessageBox.Show("이 창을 닫으면, 현재 화면에 설정된 정보가 모두 초기화 됩니다. 진행하시겠습니까?", "종료", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void Dispose()
        {
            _pollTimer?.Stop();
            _pollTimer?.Dispose();
            _spms.LineArrived -= Spms_LineArrived;
            _spms.LinesPublished -= Spms_LinesPublished;

            foreach (var pane in PacketMonitorPanes)
            {
                try
                {
                    pane.Stop();
                }
                catch { }
            }
            PacketMonitorPanes.Clear();

            foreach (var pane in TrendPanes)
            {
                try
                {
                    pane.Stop();
                }
                catch { }
            }
            TrendPanes.Clear();
        }

        public interface IKeyItem { string Key {get;}}
    }
}