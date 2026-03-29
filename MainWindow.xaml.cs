using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;

namespace WowUtils
{
    public class AppSettings
    {
        public string? LastWowPath { get; set; }
        public string? LastSelectedAccount { get; set; }
        public List<string> ProtectedAddons { get; set; } = new();
    }

    public class ConfigFileItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        public string Name { get; set; } = "";
        public string FullPath { get; set; } = "";
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    public class BackupRecord
    {
        public string BackupTime { get; set; } = "";
        public string Account { get; set; } = "";
        public int FileCount { get; set; }
        public string Size { get; set; } = "";
        public string FilePath { get; set; } = "";
        public string Remark { get; set; } = "";
        public List<string> Files { get; set; } = new();
    }

    public class AddonItem : INotifyPropertyChanged
    {
        private bool _isProtected;
        public string Name { get; set; } = "";
        public string FolderPath { get; set; } = "";
        public bool IsProtected
        {
            get => _isProtected;
            set
            {
                _isProtected = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsProtected)));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }

    internal sealed class ProtectedCleanupSnapshot
    {
        public string TempRoot { get; set; } = "";
        public int PreservedAddonFolders { get; set; }
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string? wowPath;
        private ObservableCollection<ConfigFileItem> configFiles = new();
        private ObservableCollection<ConfigFileItem> allConfigFiles = new();
        private ObservableCollection<BackupRecord> backupRecords = new();
        private ObservableCollection<AddonItem> installedAddons = new();
        private ObservableCollection<AddonItem> filteredAddons = new();
        private string backupFolder = "";
        private readonly string settingsFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "WowUtils", "settings.json");
        private int sessionCleanCount = 0;
        private DateTime? lastCleanTime = null;
        private AppSettings appSettings = new();

        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(settingsFile))
                {
                    var json = File.ReadAllText(settingsFile);
                    appSettings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                    if (appSettings.LastWowPath != null && ValidateWowPath(appSettings.LastWowPath))
                    {
                        SetWowPath(appSettings.LastWowPath);
                        AddLog($"已加载上次使用的目录: {appSettings.LastWowPath}");
                        
                        // 恢复上次选择的账号
                        if (!string.IsNullOrEmpty(appSettings.LastSelectedAccount))
                        {
                            for (int i = 0; i < cmbAccounts.Items.Count; i++)
                            {
                                if (cmbAccounts.Items[i].ToString() == appSettings.LastSelectedAccount)
                                {
                                    cmbAccounts.SelectedIndex = i;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AddLog($"加载配置失败: {ex.Message}");
            }
        }

        private void SaveSettings()
        {
            try
            {
                var protectedAddons = installedAddons.Count > 0
                    ? installedAddons.Where(item => item.IsProtected).Select(item => item.Name).OrderBy(name => name).ToList()
                    : appSettings.ProtectedAddons;

                var settings = new AppSettings
                {
                    LastWowPath = wowPath,
                    LastSelectedAccount = cmbAccounts.SelectedItem?.ToString(),
                    ProtectedAddons = protectedAddons
                };
                appSettings = settings;
                
                var dir = Path.GetDirectoryName(settingsFile);
                if (!string.IsNullOrEmpty(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(settingsFile, json);
            }
            catch { }
        }

        private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            lstConfigs.ItemsSource = configFiles;
            dgBackups.ItemsSource = backupRecords;
            lstProtectedAddons.ItemsSource = filteredAddons;
            
            // 初始化清理统计显示
            UpdateCleanStats();
            UpdateAddonStats();
            
            AddLog("程序启动...");
            
            // 先尝试加载上次的配置
            LoadSettings();
            
            // 如果没有加载到配置，才进行自动搜索
            if (string.IsNullOrEmpty(wowPath))
            {
                AddLog("正在自动搜索魔兽世界安装目录...");
                await Task.Run(() => AutoDetectWowPath());
            }
        }

        private void AutoDetectWowPath()
        {
            // 常见的安装路径
            var possiblePaths = new List<string>
            {
                @"C:\Program Files (x86)\World of Warcraft",
                @"C:\Program Files\World of Warcraft",
                @"D:\World of Warcraft",
                @"E:\World of Warcraft",
                @"F:\World of Warcraft",
            };

            // 尝试从注册表读取
            try
            {
                using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Blizzard Entertainment\World of Warcraft");
                if (key != null)
                {
                    var installPath = key.GetValue("InstallPath")?.ToString();
                    if (!string.IsNullOrEmpty(installPath))
                    {
                        possiblePaths.Insert(0, installPath);
                    }
                }
            }
            catch { }

            // 检查每个可能的路径
            foreach (var path in possiblePaths)
            {
                if (ValidateWowPath(path))
                {
                    Dispatcher.Invoke(() => SetWowPath(path));
                    return;
                }
            }

            Dispatcher.Invoke(() =>
            {
                AddLog("未能自动检测到魔兽世界安装目录，请手动选择或使用全局搜索");
                txtPathStatus.Text = "请手动选择魔兽世界安装目录";
                txtPathStatus.Foreground = System.Windows.Media.Brushes.Orange;
            });
        }

        private bool ValidateWowPath(string path)
        {
            if (string.IsNullOrEmpty(path) || !Directory.Exists(path))
                return false;

            // 检查 World of Warcraft Launcher.exe
            var launcherPath = Path.Combine(path, "World of Warcraft Launcher.exe");
            if (!File.Exists(launcherPath))
                return false;

            // 检查 _retail_/Wow.exe
            var retailPath = Path.Combine(path, "_retail_");
            var wowExePath = Path.Combine(retailPath, "Wow.exe");
            if (!Directory.Exists(retailPath) || !File.Exists(wowExePath))
                return false;

            return true;
        }

        private void SetWowPath(string path)
        {
            wowPath = path;
            txtWowPath.Text = path;
            txtPathStatus.Text = "✓ 目录验证成功";
            txtPathStatus.Foreground = System.Windows.Media.Brushes.Green;
            btnClean.IsEnabled = true;
            AddLog($"已检测到魔兽世界安装目录: {path}");
            
            // 初始化备份文件夹
            backupFolder = Path.Combine(path, "WowUtils_Backups");
            Directory.CreateDirectory(backupFolder);
            
            // 刷新账号列表
            RefreshAccounts();
            LoadBackupRecords();
            LoadInstalledAddons();
            
            // 保存配置
            SaveSettings();
        }

        private async void BtnGlobalSearch_Click(object sender, RoutedEventArgs e)
        {
            btnGlobalSearch.IsEnabled = false;
            btnBrowse.IsEnabled = false;
            AddLog("========================================");
            AddLog("开始全局搜索魔兽世界安装目录...");
            AddLog("这可能需要几分钟时间，请耐心等待...");

            string? foundPath = null;
            await Task.Run(() =>
            {
                var drives = DriveInfo.GetDrives()
                    .Where(d => d.DriveType == DriveType.Fixed && d.IsReady)
                    .ToList();

                AddLog($"准备搜索 {drives.Count} 个磁盘: {string.Join(", ", drives.Select(d => d.Name))}");

                foreach (var drive in drives)
                {
                    AddLog($"正在搜索磁盘 {drive.Name}...");
                    foundPath = SearchWowPath(drive.RootDirectory.FullName);
                    if (foundPath != null)
                    {
                        AddLog($"✓ 找到魔兽世界安装目录: {foundPath}");
                        break;
                    }
                }

                if (foundPath == null)
                {
                    AddLog("✗ 未找到魔兽世界安装目录");
                }
            });

            if (foundPath != null)
            {
                SetWowPath(foundPath);
            }
            else
            {
                MessageBox.Show("未能找到魔兽世界安装目录，请手动选择。", "搜索完成", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }

            btnGlobalSearch.IsEnabled = true;
            btnBrowse.IsEnabled = true;
        }

        private string? SearchWowPath(string rootPath)
        {
            try
            {
                // 先检查常见文件夹名
                var commonNames = new[] { "World of Warcraft", "魔兽世界", "WoW" };
                foreach (var name in commonNames)
                {
                    var testPath = Path.Combine(rootPath, name);
                    if (ValidateWowPath(testPath))
                        return testPath;
                }

                // 搜索目录（限制深度避免太慢）
                return SearchWowPathRecursive(rootPath, 0, 3);
            }
            catch
            {
                return null;
            }
        }

        private string? SearchWowPathRecursive(string currentPath, int depth, int maxDepth)
        {
            if (depth > maxDepth)
                return null;

            try
            {
                // 检查当前目录
                if (ValidateWowPath(currentPath))
                    return currentPath;

                // 搜索子目录
                var directories = Directory.GetDirectories(currentPath);
                foreach (var dir in directories)
                {
                    // 跳过系统目录
                    var dirName = Path.GetFileName(dir).ToLower();
                    if (dirName == "windows" || dirName == "program files" || 
                        dirName == "$recycle.bin" || dirName == "system volume information")
                        continue;

                    var result = SearchWowPathRecursive(dir, depth + 1, maxDepth);
                    if (result != null)
                        return result;
                }
            }
            catch { }

            return null;
        }

        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog
            {
                Title = "选择魔兽世界安装目录",
                Multiselect = false
            };

            if (dialog.ShowDialog() == true)
            {
                var selectedPath = dialog.FolderName;
                if (ValidateWowPath(selectedPath))
                {
                    SetWowPath(selectedPath);
                }
                else
                {
                    wowPath = null;
                    txtPathStatus.Text = "✗ 目录验证失败！请选择正确的魔兽世界安装目录";
                    txtPathStatus.Foreground = System.Windows.Media.Brushes.Red;
                    btnClean.IsEnabled = false;
                    AddLog("目录验证失败！");
                    MessageBox.Show(
                        "所选目录不是有效的魔兽世界安装目录！\n\n" +
                        "有效的目录应包含：\n" +
                        "• World of Warcraft Launcher.exe\n" +
                        "• _retail_\\Wow.exe",
                        "目录验证失败",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
        }

        private void RefreshAccounts()
        {
            cmbAccounts.Items.Clear();
            if (string.IsNullOrEmpty(wowPath))
                return;

            var wtfPath = Path.Combine(wowPath, "_retail_", "WTF", "Account");
            if (!Directory.Exists(wtfPath))
            {
                AddLog("未找到 WTF/Account 目录");
                return;
            }

            var accounts = Directory.GetDirectories(wtfPath)
                .Select(Path.GetFileName)
                .Where(name => !string.IsNullOrEmpty(name))
                .ToList();

            foreach (var account in accounts)
            {
                cmbAccounts.Items.Add(account);
            }

            if (cmbAccounts.Items.Count > 0)
            {
                cmbAccounts.SelectedIndex = 0;
                AddLog($"找到 {cmbAccounts.Items.Count} 个账号");
            }
            else
            {
                AddLog("未找到任何账号");
            }
        }

        private void BtnRefreshAccounts_Click(object sender, RoutedEventArgs e)
        {
            RefreshAccounts();
        }

        private void CmbAccounts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbAccounts.SelectedItem == null)
                return;

            LoadConfigFiles(cmbAccounts.SelectedItem.ToString()!);
            SaveSettings();
        }

        private void LoadConfigFiles(string accountName)
        {
            allConfigFiles.Clear();
            configFiles.Clear();

            if (string.IsNullOrEmpty(wowPath))
                return;

            var accountPath = Path.Combine(wowPath, "_retail_", "WTF", "Account", accountName);
            var savedVarsPath = Path.Combine(accountPath, "SavedVariables");

            // 加载账号级SavedVariables
            if (Directory.Exists(savedVarsPath))
            {
                var files = Directory.GetFiles(savedVarsPath, "*.lua");
                foreach (var file in files)
                {
                    allConfigFiles.Add(new ConfigFileItem
                    {
                        Name = Path.GetFileName(file),
                        FullPath = file,
                        IsSelected = false
                    });
                }
            }

            // 加载服务器/角色级SavedVariables
            try
            {
                var serverDirs = Directory.GetDirectories(accountPath)
                    .Where(d => !Path.GetFileName(d).Equals("SavedVariables", StringComparison.OrdinalIgnoreCase));

                foreach (var serverDir in serverDirs)
                {
                    var charDirs = Directory.GetDirectories(serverDir);
                    foreach (var charDir in charDirs)
                    {
                        var charSavedVars = Path.Combine(charDir, "SavedVariables");
                        if (Directory.Exists(charSavedVars))
                        {
                            var files = Directory.GetFiles(charSavedVars, "*.lua");
                            foreach (var file in files)
                            {
                                var fileName = Path.GetFileName(file);
                                // 避免重复添加
                                if (!allConfigFiles.Any(cf => cf.Name == fileName))
                                {
                                    allConfigFiles.Add(new ConfigFileItem
                                    {
                                        Name = fileName,
                                        FullPath = file,
                                        IsSelected = false
                                    });
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            // 更新显示列表
            foreach (var item in allConfigFiles.OrderBy(f => f.Name))
            {
                configFiles.Add(item);
                item.PropertyChanged += (s, e) => UpdateConfigStats();
            }

            AddLog($"加载了 {configFiles.Count} 个插件配置文件");
            UpdateConfigStats();
        }

        private void LoadInstalledAddons()
        {
            installedAddons.Clear();
            filteredAddons.Clear();

            if (string.IsNullOrEmpty(wowPath))
            {
                txtAddonSettingsStatus.Text = "请先选择有效的魔兽世界目录。";
                txtAddonSettingsStatus.Foreground = System.Windows.Media.Brushes.Gray;
                UpdateAddonStats();
                return;
            }

            var addOnsPath = Path.Combine(wowPath, "_retail_", "Interface", "AddOns");
            if (!Directory.Exists(addOnsPath))
            {
                txtAddonSettingsStatus.Text = "未找到 Interface/AddOns 目录。";
                txtAddonSettingsStatus.Foreground = System.Windows.Media.Brushes.Orange;
                UpdateAddonStats();
                return;
            }

            var protectedAddons = appSettings.ProtectedAddons.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var addonFolders = Directory.GetDirectories(addOnsPath)
                .Select(path => new
                {
                    Name = Path.GetFileName(path),
                    Path = path
                })
                .Where(item => !string.IsNullOrWhiteSpace(item.Name))
                .OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var addon in addonFolders)
            {
                var item = new AddonItem
                {
                    Name = addon.Name!,
                    FolderPath = addon.Path,
                    IsProtected = protectedAddons.Contains(addon.Name!)
                };
                item.PropertyChanged += AddonItem_PropertyChanged;
                installedAddons.Add(item);
            }

            txtAddonSettingsStatus.Text = "插件列表来自 Interface/AddOns 目录。勾选项只影响一键清理时的插件目录保留。";
            txtAddonSettingsStatus.Foreground = System.Windows.Media.Brushes.Green;
            AddLog($"加载了 {installedAddons.Count} 个已安装插件");
            ApplyAddonFilter();
            UpdateAddonStats();
        }

        private void AddonItem_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(AddonItem.IsProtected))
            {
                UpdateAddonStats();
                SaveSettings();
            }
        }

        private void UpdateAddonStats()
        {
            var totalCount = installedAddons.Count;
            var visibleCount = filteredAddons.Count;
            var protectedCount = installedAddons.Count(item => item.IsProtected);
            txtAddonStats.Text = $"共 {totalCount} 个插件，当前显示 {visibleCount} 个，保留 {protectedCount} 个";
        }

        private List<string> GetProtectedAddonNames()
        {
            return installedAddons
                .Where(item => item.IsProtected)
                .Select(item => item.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void BtnRefreshAddons_Click(object sender, RoutedEventArgs e)
        {
            LoadInstalledAddons();
            SaveSettings();
        }

        private void BtnProtectAllAddons_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in filteredAddons)
            {
                item.IsProtected = true;
            }

            UpdateAddonStats();
            SaveSettings();
        }

        private void BtnClearProtectedAddons_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in filteredAddons)
            {
                item.IsProtected = false;
            }

            UpdateAddonStats();
            SaveSettings();
        }

        private void TxtAddonSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyAddonFilter();
        }

        private void ApplyAddonFilter()
        {
            var searchText = txtAddonSearch.Text.Trim();
            filteredAddons.Clear();

            var filtered = string.IsNullOrWhiteSpace(searchText)
                ? installedAddons
                : installedAddons.Where(item => item.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase));

            foreach (var item in filtered)
            {
                filteredAddons.Add(item);
            }

            UpdateAddonStats();
        }

        private void TxtConfigSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            var searchText = txtConfigSearch.Text.ToLower();
            configFiles.Clear();

            var filtered = string.IsNullOrWhiteSpace(searchText)
                ? allConfigFiles
                : allConfigFiles.Where(f => f.Name.ToLower().Contains(searchText));

            foreach (var item in filtered)
            {
                configFiles.Add(item);
            }
            
            UpdateConfigStats();
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in configFiles)
            {
                item.IsSelected = true;
            }
            UpdateConfigStats();
        }

        private void BtnSelectNone_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in configFiles)
            {
                item.IsSelected = false;
            }
            UpdateConfigStats();
        }

        private void UpdateConfigStats()
        {
            var totalCount = allConfigFiles.Count;
            var selectedCount = allConfigFiles.Count(f => f.IsSelected);
            txtConfigStats.Text = $"共 {totalCount} 个插件配置，已选择 {selectedCount} 个";
        }

        private void BtnOpenWowFolder_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(wowPath))
            {
                MessageBox.Show("请先选择魔兽世界安装目录！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                OpenPathInExplorer(wowPath, false);
                AddLog($"已打开目录: {wowPath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开目录失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateCleanStats()
        {
            if (sessionCleanCount == 0)
            {
                txtCleanStats.Text = "本次会话：未清理";
            }
            else
            {
                txtCleanStats.Text = $"本次会话：已清理 {sessionCleanCount} 次";
            }

            if (lastCleanTime.HasValue)
            {
                txtLastCleanTime.Text = $"上次清理：{lastCleanTime.Value:yyyy-MM-dd HH:mm:ss}";
            }
            else
            {
                txtLastCleanTime.Text = "上次清理：无记录";
            }
        }

        private async void BtnBackupConfigs_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = allConfigFiles.Where(f => f.IsSelected).ToList();
            if (selectedItems.Count == 0)
            {
                MessageBox.Show("请至少选择一个配置文件！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var accountName = cmbAccounts.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(accountName))
                return;

            // 显示将要备份的插件列表
            var pluginList = string.Join("\n", selectedItems.Select(item => $"  • {item.Name}"));
            var message = $"确定要备份以下 {selectedItems.Count} 个插件配置吗？\n\n" +
                         $"账号：{accountName}\n\n" +
                         $"插件列表：\n{pluginList}";
            
            var result = MessageBox.Show(message, "确认备份", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
                return;

            // 输入备注
            var remarkWindow = new Window
            {
                Title = "输入备注",
                Width = 400,
                Height = 180,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ResizeMode = ResizeMode.NoResize
            };
            
            var grid = new Grid { Margin = new Thickness(15) };
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Auto) });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Auto) });
            
            var label = new TextBlock { Text = "请输入备份备注（可选）：", Margin = new Thickness(0, 0, 0, 10) };
            Grid.SetRow(label, 0);
            
            var textBox = new TextBox 
            { 
                AcceptsReturn = true, 
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                MaxLength = 200
            };
            Grid.SetRow(textBox, 1);
            
            var buttonPanel = new StackPanel 
            { 
                Orientation = Orientation.Horizontal, 
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };
            var okButton = new Button { Content = "确定", Width = 80, Height = 30, Margin = new Thickness(0, 0, 10, 0), IsDefault = true };
            okButton.Click += (s, args) => { remarkWindow.DialogResult = true; remarkWindow.Close(); };
            var cancelButton = new Button { Content = "取消", Width = 80, Height = 30, IsCancel = true };
            cancelButton.Click += (s, args) => { remarkWindow.DialogResult = false; remarkWindow.Close(); };
            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);
            Grid.SetRow(buttonPanel, 2);
            
            grid.Children.Add(label);
            grid.Children.Add(textBox);
            grid.Children.Add(buttonPanel);
            remarkWindow.Content = grid;
            
            string remark = "";
            if (remarkWindow.ShowDialog() == true)
            {
                remark = textBox.Text.Trim();
            }
            else
            {
                return;
            }

            btnBackupConfigs.IsEnabled = false;

            try
            {
                await Task.Run(() => BackupConfigs(accountName, selectedItems, remark));
                MessageBox.Show($"成功备份 {selectedItems.Count} 个配置文件！", "备份完成", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                LoadBackupRecords();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"备份失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnBackupConfigs.IsEnabled = true;
            }
        }

        private void BackupConfigs(string accountName, List<ConfigFileItem> items, string remark)
        {
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var backupFileName = $"Config_Backup_{accountName}_{timestamp}.zip";
            var backupPath = Path.Combine(backupFolder, backupFileName);

            AddLog("========================================");
            AddLog($"开始备份配置文件...");
            AddLog($"账号: {accountName}");
            AddLog($"插件数量: {items.Count}");
            AddLog("");
            AddLog("将备份以下插件配置:");
            foreach (var item in items)
            {
                AddLog($"  • {item.Name}");
            }
            AddLog("");

            var accountPath = Path.Combine(wowPath!, "_retail_", "WTF", "Account", accountName);
            var filesToBackup = new List<string>();

            using (var zip = ZipFile.Open(backupPath, ZipArchiveMode.Create))
            {
                foreach (var item in items)
                {
                    AddLog($"  正在备份: {item.Name}");

                    // 备份账号级SavedVariables
                    var savedVarsPath = Path.Combine(accountPath, "SavedVariables", item.Name);
                    if (File.Exists(savedVarsPath))
                    {
                        zip.CreateEntryFromFile(savedVarsPath, $"SavedVariables/{item.Name}");
                        filesToBackup.Add($"SavedVariables/{item.Name}");
                    }

                    // 备份所有服务器/角色级SavedVariables
                    try
                    {
                        var serverDirs = Directory.GetDirectories(accountPath)
                            .Where(d => !Path.GetFileName(d).Equals("SavedVariables", StringComparison.OrdinalIgnoreCase));

                        foreach (var serverDir in serverDirs)
                        {
                            var serverName = Path.GetFileName(serverDir);
                            var charDirs = Directory.GetDirectories(serverDir);
                            foreach (var charDir in charDirs)
                            {
                                var charName = Path.GetFileName(charDir);
                                var charConfigPath = Path.Combine(charDir, "SavedVariables", item.Name);
                                if (File.Exists(charConfigPath))
                                {
                                    var entryName = $"{serverName}/{charName}/SavedVariables/{item.Name}";
                                    zip.CreateEntryFromFile(charConfigPath, entryName);
                                    filesToBackup.Add(entryName);
                                }
                            }
                        }
                    }
                    catch { }
                }
            }

            // 保存备份记录
            var fileInfo = new FileInfo(backupPath);
            var record = new BackupRecord
            {
                BackupTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                Account = accountName,
                FileCount = filesToBackup.Count,
                Size = FormatFileSize(fileInfo.Length),
                FilePath = backupPath,
                Files = filesToBackup,
                Remark = remark
            };

            SaveBackupRecord(record);

            AddLog($"✓ 备份完成！");
            AddLog($"  备份文件: {backupFileName}");
            AddLog($"  文件数量: {filesToBackup.Count}");
            AddLog($"  文件大小: {record.Size}");
        }

        private void SaveBackupRecord(BackupRecord record)
        {
            var recordFile = Path.Combine(backupFolder, "backup_records.json");
            List<BackupRecord> records = new();

            if (File.Exists(recordFile))
            {
                try
                {
                    var json = File.ReadAllText(recordFile);
                    records = JsonSerializer.Deserialize<List<BackupRecord>>(json) ?? new();
                }
                catch { }
            }

            records.Insert(0, record);
            var newJson = JsonSerializer.Serialize(records, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(recordFile, newJson);
        }

        private void LoadBackupRecords()
        {
            backupRecords.Clear();
            if (string.IsNullOrEmpty(backupFolder))
                return;

            var recordFile = Path.Combine(backupFolder, "backup_records.json");
            if (!File.Exists(recordFile))
                return;

            try
            {
                var json = File.ReadAllText(recordFile);
                var records = JsonSerializer.Deserialize<List<BackupRecord>>(json);
                if (records != null)
                {
                    foreach (var record in records)
                    {
                        // 检查文件是否还存在
                        if (File.Exists(record.FilePath))
                        {
                            backupRecords.Add(record);
                        }
                    }
                }
            }
            catch { }
        }

        private void DgBackups_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selected = dgBackups.SelectedItem as BackupRecord;
            btnRestore.IsEnabled = selected != null;
            btnDeleteBackup.IsEnabled = selected != null;

            if (selected != null)
            {
                var previewFiles = selected.Files.Take(5).ToList();
                var details = previewFiles.Count > 0
                    ? string.Join("\n", previewFiles)
                    : "无文件详情";

                if (selected.Files.Count > previewFiles.Count)
                {
                    details += $"\n... 另有 {selected.Files.Count - previewFiles.Count} 个文件";
                }

                txtBackupDetails.Text =
                    $"双击可打开备份位置。\n文件总数：{selected.FileCount}\n预览：\n{details}";
            }
            else
            {
                txtBackupDetails.Text = "双击备份记录可打开备份文件所在位置。";
            }
        }

        private void DgBackups_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (dgBackups.SelectedItem is not BackupRecord selected)
                return;

            if (!File.Exists(selected.FilePath))
            {
                MessageBox.Show("备份文件不存在，列表将刷新。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                LoadBackupRecords();
                return;
            }

            try
            {
                OpenPathInExplorer(selected.FilePath, true);
                AddLog($"已打开备份位置: {selected.FilePath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开备份位置失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void BtnRestore_Click(object sender, RoutedEventArgs e)
        {
            var selected = dgBackups.SelectedItem as BackupRecord;
            if (selected == null)
                return;

            var result = MessageBox.Show(
                $"确定要恢复此备份吗？\n\n" +
                $"备份时间: {selected.BackupTime}\n" +
                $"账号: {selected.Account}\n" +
                $"文件数: {selected.FileCount}\n\n" +
                $"注意：这将覆盖当前的配置文件！",
                "确认恢复",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
                return;

            btnRestore.IsEnabled = false;

            try
            {
                await Task.Run(() => RestoreBackup(selected));
                MessageBox.Show("配置恢复成功！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"恢复失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnRestore.IsEnabled = true;
            }
        }

        private void RestoreBackup(BackupRecord record)
        {
            AddLog("========================================");
            AddLog($"开始恢复备份...");
            AddLog($"备份时间: {record.BackupTime}");
            AddLog($"账号: {record.Account}");

            var accountPath = Path.Combine(wowPath!, "_retail_", "WTF", "Account", record.Account);

            using (var zip = ZipFile.OpenRead(record.FilePath))
            {
                foreach (var entry in zip.Entries)
                {
                    if (entry.FullName.EndsWith("/"))
                        continue;

                    AddLog($"  正在恢复: {entry.FullName}");

                    var destPath = Path.Combine(accountPath, entry.FullName);
                    var destDir = Path.GetDirectoryName(destPath)!;
                    Directory.CreateDirectory(destDir);

                    entry.ExtractToFile(destPath, true);
                }
            }

            AddLog($"✓ 恢复完成！共恢复 {record.FileCount} 个文件");
        }

        private void BtnDeleteBackup_Click(object sender, RoutedEventArgs e)
        {
            var selected = dgBackups.SelectedItem as BackupRecord;
            if (selected == null)
                return;

            var result = MessageBox.Show(
                $"确定要删除此备份吗？\n\n" +
                $"备份时间: {selected.BackupTime}\n" +
                $"账号: {selected.Account}\n" +
                $"文件数: {selected.FileCount}\n\n" +
                $"此操作不可恢复！",
                "确认删除",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes)
                return;

            try
            {
                if (File.Exists(selected.FilePath))
                {
                    File.Delete(selected.FilePath);
                }
                backupRecords.Remove(selected);
                
                // 更新记录文件
                var recordFile = Path.Combine(backupFolder, "backup_records.json");
                var records = backupRecords.ToList();
                var json = JsonSerializer.Serialize(records, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(recordFile, json);

                AddLog($"已删除备份: {Path.GetFileName(selected.FilePath)}");
                MessageBox.Show("备份已删除！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void BtnClean_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(wowPath))
            {
                MessageBox.Show("请先选择魔兽世界安装目录！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 检查是否至少选择了一个文件夹
            if (!chkCache.IsChecked.GetValueOrDefault() &&
                !chkFonts.IsChecked.GetValueOrDefault() &&
                !chkInterface.IsChecked.GetValueOrDefault() &&
                !chkWTF.IsChecked.GetValueOrDefault())
            {
                MessageBox.Show("请至少选择一个要清理的文件夹！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var foldersToDelete = new List<string>();
            if (chkCache.IsChecked.GetValueOrDefault()) foldersToDelete.Add("Cache");
            if (chkFonts.IsChecked.GetValueOrDefault()) foldersToDelete.Add("Fonts");
            if (chkInterface.IsChecked.GetValueOrDefault()) foldersToDelete.Add("Interface");
            if (chkWTF.IsChecked.GetValueOrDefault()) foldersToDelete.Add("WTF");

            var message = $"确定要清理以下文件夹吗？\n\n{string.Join("\n", foldersToDelete)}";
            var protectedAddons = GetProtectedAddonNames();
            if (protectedAddons.Count > 0 && foldersToDelete.Contains("Interface"))
            {
                message += $"\n\n设置中已配置保留 {protectedAddons.Count} 个插件，对应插件目录会被保留。";
            }
            if (chkBackup.IsChecked.GetValueOrDefault())
            {
                message += "\n\n将在删除前创建备份。";
            }

            var result = MessageBox.Show(message, "确认清理", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
                return;

            // 禁用按钮防止重复点击
            btnClean.IsEnabled = false;
            btnBrowse.IsEnabled = false;
            btnGlobalSearch.IsEnabled = false;

            try
            {
                await Task.Run(() => PerformCleanup(foldersToDelete));
                sessionCleanCount++;
                lastCleanTime = DateTime.Now;
                UpdateCleanStats();
                AddLog("========================================");
                AddLog("清理完成！");
                MessageBox.Show("插件清理完成！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AddLog($"错误: {ex.Message}");
                MessageBox.Show($"清理过程中发生错误：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                btnClean.IsEnabled = true;
                btnBrowse.IsEnabled = true;
                btnGlobalSearch.IsEnabled = true;
            }
        }

        private void PerformCleanup(List<string> foldersToDelete)
        {
            var retailPath = Path.Combine(wowPath!, "_retail_");
            var protectedAddons = Dispatcher.Invoke(GetProtectedAddonNames)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            ProtectedCleanupSnapshot? snapshot = null;
            
            AddLog("========================================");
            AddLog("开始清理操作...");
            AddLog($"目标目录: {retailPath}");
            if (protectedAddons.Count > 0 && foldersToDelete.Contains("Interface"))
            {
                AddLog($"保留插件数量: {protectedAddons.Count}");
                AddLog($"保留插件列表: {string.Join(", ", protectedAddons.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))}");
            }
            AddLog("");

            try
            {
                snapshot = CaptureProtectedCleanupSnapshot(retailPath, foldersToDelete, protectedAddons);

                // 备份
                if (Dispatcher.Invoke(() => chkBackup.IsChecked.GetValueOrDefault()))
                {
                    try
                    {
                        var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        var backupFileName = $"Cleanup_Backup_{timestamp}.zip";
                        var backupPath = Path.Combine(wowPath!, backupFileName);

                        AddLog($"正在创建备份: {backupFileName}");
                        AddLog("压缩中，请稍候...");

                        using (var zip = ZipFile.Open(backupPath, ZipArchiveMode.Create))
                        {
                            foreach (var folderName in foldersToDelete)
                            {
                                var folderPath = Path.Combine(retailPath, folderName);
                                if (Directory.Exists(folderPath))
                                {
                                    AddLog($"  正在压缩: {folderName}");
                                    AddFilesToZip(zip, folderPath, folderName);
                                }
                            }
                        }

                        var fileInfo = new FileInfo(backupPath);
                        AddLog($"✓ 备份创建成功！");
                        AddLog($"  文件大小: {FormatFileSize(fileInfo.Length)}");
                        AddLog($"  保存位置: {backupPath}");
                        AddLog("");
                    }
                    catch (Exception ex)
                    {
                        AddLog($"✗ 备份失败: {ex.Message}");
                        AddLog("继续执行删除操作...");
                        AddLog("");
                    }
                }

                // 删除文件夹
                AddLog("开始删除文件夹...");
                foreach (var folderName in foldersToDelete)
                {
                    var folderPath = Path.Combine(retailPath, folderName);
                    if (Directory.Exists(folderPath))
                    {
                        try
                        {
                            AddLog($"正在删除: {folderName}");
                            Directory.Delete(folderPath, true);
                            AddLog($"✓ {folderName} 删除成功");
                        }
                        catch (Exception ex)
                        {
                            AddLog($"✗ {folderName} 删除失败: {ex.Message}");
                        }
                    }
                    else
                    {
                        AddLog($"○ {folderName} 文件夹不存在，跳过");
                    }
                }

                RestoreProtectedCleanupSnapshot(retailPath, snapshot);
            }
            finally
            {
                CleanupProtectedCleanupSnapshot(snapshot);
            }
        }

        private ProtectedCleanupSnapshot? CaptureProtectedCleanupSnapshot(string retailPath, IReadOnlyCollection<string> foldersToDelete, HashSet<string> protectedAddons)
        {
            var shouldPreserveInterface = foldersToDelete.Contains("Interface") && protectedAddons.Count > 0;
            if (!shouldPreserveInterface)
                return null;

            var snapshot = new ProtectedCleanupSnapshot
            {
                TempRoot = Path.Combine(Path.GetTempPath(), "WowUtils", $"cleanup_{Guid.NewGuid():N}")
            };
            Directory.CreateDirectory(snapshot.TempRoot);

            if (shouldPreserveInterface)
            {
                snapshot.PreservedAddonFolders = CopyProtectedAddonFolders(retailPath, snapshot.TempRoot, protectedAddons);
                AddLog($"已暂存 {snapshot.PreservedAddonFolders} 个需要保留的插件目录");
            }

            return snapshot;
        }

        private int CopyProtectedAddonFolders(string retailPath, string tempRoot, HashSet<string> protectedAddons)
        {
            var addOnsPath = Path.Combine(retailPath, "Interface", "AddOns");
            if (!Directory.Exists(addOnsPath))
                return 0;

            var copiedCount = 0;
            foreach (var addonDir in Directory.GetDirectories(addOnsPath))
            {
                var addonName = Path.GetFileName(addonDir);
                if (string.IsNullOrWhiteSpace(addonName) || !protectedAddons.Contains(addonName))
                    continue;

                var targetDir = Path.Combine(tempRoot, "Interface", "AddOns", addonName);
                CopyDirectory(addonDir, targetDir);
                copiedCount++;
            }

            return copiedCount;
        }

        private void RestoreProtectedCleanupSnapshot(string retailPath, ProtectedCleanupSnapshot? snapshot)
        {
            if (snapshot == null)
                return;

            if (snapshot.PreservedAddonFolders > 0)
            {
                var sourcePath = Path.Combine(snapshot.TempRoot, "Interface");
                if (Directory.Exists(sourcePath))
                {
                    CopyDirectory(sourcePath, Path.Combine(retailPath, "Interface"));
                    AddLog($"已恢复 {snapshot.PreservedAddonFolders} 个保留插件目录");
                }
            }
        }

        private void CleanupProtectedCleanupSnapshot(ProtectedCleanupSnapshot? snapshot)
        {
            if (snapshot == null || string.IsNullOrWhiteSpace(snapshot.TempRoot) || !Directory.Exists(snapshot.TempRoot))
                return;

            try
            {
                Directory.Delete(snapshot.TempRoot, true);
            }
            catch
            {
            }
        }

        private void CopyDirectory(string sourceDir, string targetDir)
        {
            Directory.CreateDirectory(targetDir);

            foreach (var file in Directory.GetFiles(sourceDir))
            {
                var targetFile = Path.Combine(targetDir, Path.GetFileName(file));
                File.Copy(file, targetFile, true);
            }

            foreach (var directory in Directory.GetDirectories(sourceDir))
            {
                var targetSubDir = Path.Combine(targetDir, Path.GetFileName(directory));
                CopyDirectory(directory, targetSubDir);
            }
        }

        private void AddFilesToZip(ZipArchive zip, string folderPath, string entryPrefix)
        {
            var files = Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                try
                {
                    var relativePath = Path.GetRelativePath(folderPath, file);
                    var entryName = Path.Combine(entryPrefix, relativePath).Replace("\\", "/");
                    zip.CreateEntryFromFile(file, entryName);
                }
                catch { }
            }
        }

        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }

        private void AddLog(string message)
        {
            Dispatcher.Invoke(() =>
            {
                var timestamp = DateTime.Now.ToString("HH:mm:ss");
                txtLog.AppendText($"[{timestamp}] {message}\n");
                txtLog.ScrollToEnd();
            });
        }

        private void OpenPathInExplorer(string path, bool selectFile)
        {
            var arguments = selectFile
                ? $"/select,\"{path}\""
                : $"\"{path}\"";

            Process.Start(new ProcessStartInfo("explorer.exe", arguments)
            {
                UseShellExecute = true
            });
        }
    }
}
