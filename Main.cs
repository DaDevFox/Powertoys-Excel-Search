using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ManagedCommon;
using Microsoft.Office.Interop.Excel;
using Microsoft.PowerToys.Settings.UI.Library;
using Wox.Plugin;
using Wox.Plugin.Logger;
using Excel = Microsoft.Office.Interop.Excel;

namespace Community.Powertoys.Run.Plugin.ExcelSearch
{

    /// <summary>
    /// Main class of this plugin that implement all used interfaces.
    /// </summary>
    public class Main : IPlugin, IContextMenu, ISettingProvider, IDisposable
    {
        private static Excel.RecentFiles Files { get; } = (new Excel.Application()).RecentFiles;

        /// <summary>
        /// ID of the plugin.
        /// </summary>
        public static string PluginID => "A2bd4fbacda4041f3a7fe22b33e79a99b";

        /// <summary>
        /// Name of the plugin.
        /// </summary>
        public string Name => "ExcelSearch";

        /// <summary>
        /// Description of the plugin.
        /// </summary>
        public string Description => "Searches recent excel sheets.";

        /// <summary>
        /// Additional options for the plugin.
        /// </summary>
        public IEnumerable<PluginAdditionalOption> AdditionalOptions => [
            new()
            {
                Key = nameof(IndexSearch),
                DisplayLabel = "Index search",
                DisplayDescription = "Index search db",
                PluginOptionType = PluginAdditionalOption.AdditionalOptionType.Checkbox,
                Value = IndexSearch,
            }
        ];

        private bool IndexSearch { get; set; }

        private PluginInitContext? Context { get; set; }

        private string? IconPath { get; set; }

        private bool Disposed { get; set; }

        /// <summary>
        /// Return a filtered list, based on the given query.
        /// </summary>
        /// <param name="query">The query to filter the list.</param>
        /// <returns>A filtered list, can be empty when nothing was found.</returns>
        public List<Result> Query(Query query)
        {
            Log.Info("Query: " + query.Search, GetType());
            bool showQueryInSearch = false;

            return (from file in Files.OfType<RecentFile>()
                    where File.Exists(file.Name) && file.Name.FuzzyBitap(query.Search, 2) > 5
                    orderby file.Name.FuzzyBitap(query.Search, 2) descending
                    select new Result()
                    {
                        QueryTextDisplay = query.Search,
                        IcoPath = IconPath,
                        Title =
                          showQueryInSearch ?
                            file.Name.EllipsifyInterpolatedQuery(query.Search) :
                            Path.GetFileNameWithoutExtension(file.Name),
                        SubTitle = $"Last modified {new FileInfo(file.Name).LastWriteTime.ToString("d")}",
                        ToolTipData = new ToolTipData("File Path", $"{file.Name}"),
                        ContextData = (file.Name, file.Index),
                    }).ToList();
        }

        /// <summary>
        /// Initialize the plugin with the given <see cref="PluginInitContext"/>.
        /// </summary>
        /// <param name="context">The <see cref="PluginInitContext"/> for this plugin.</param>
        public void Init(PluginInitContext context)
        {
            Log.Info("Init", GetType());

            Context = context ?? throw new ArgumentNullException(nameof(context));
            Context.API.ThemeChanged += OnThemeChanged;
            UpdateIconPath(Context.API.GetCurrentTheme());
        }

        /// <summary>
        /// Return a list context menu entries for a given <see cref="Result"/> (shown at the right side of the result).
        /// </summary>
        /// <param name="selectedResult">The <see cref="Result"/> for the list with context menu entries.</param>
        /// <returns>A list context menu entries.</returns>
        public List<ContextMenuResult> LoadContextMenus(Result selectedResult)
        {
            Log.Info("LoadContextMenus", GetType());

            if (selectedResult?.ContextData is (string fileName, int index))
            {
                return
                [
                    new ContextMenuResult
                    {
                        PluginName = Name,
                        Title = "Open (Enter)",
                        FontFamily = "Segoe Fluent Icons,Segoe MDL2 Assets",
                        Glyph = "📄", // Open File Icon
                        AcceleratorKey = Key.Enter,
                        Action = _ => {
                            Excel.Application excel = new();
                            excel.Workbooks.Open(fileName);
                            return true;
                            }
                    }
                ];
            }

            return [];
        }

        /// <summary>
        /// Creates setting panel.
        /// </summary>
        /// <returns>The control.</returns>
        /// <exception cref="NotImplementedException">method is not implemented.</exception>
        public Control CreateSettingPanel() => throw new NotImplementedException();

        /// <summary>
        /// Updates settings.
        /// </summary>
        /// <param name="settings">The plugin settings.</param>
        public void UpdateSettings(PowerLauncherPluginSettings settings)
        {
            Log.Info("UpdateSettings", GetType());

            IndexSearch = settings.AdditionalOptions.SingleOrDefault(x => x.Key == nameof(IndexSearch))?.Value ?? false;
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            Log.Info("Dispose", GetType());

            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Wrapper method for <see cref="Dispose()"/> that dispose additional objects and events form the plugin itself.
        /// </summary>
        /// <param name="disposing">Indicate that the plugin is disposed.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (Disposed || !disposing)
            {
                return;
            }

            if (Context?.API != null)
            {
                Context.API.ThemeChanged -= OnThemeChanged;
            }

            Disposed = true;
        }

        private void UpdateIconPath(Theme theme) => IconPath = theme == Theme.Light || theme == Theme.HighContrastWhite ? Context?.CurrentPluginMetadata.IcoPathLight : Context?.CurrentPluginMetadata.IcoPathDark;

        private void OnThemeChanged(Theme currentTheme, Theme newTheme) => UpdateIconPath(newTheme);

        private static bool CopyToClipboard(string? value)
        {
            if (value != null)
            {
                Clipboard.SetText(value);
            }

            return true;
        }
    }
}
