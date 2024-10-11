using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ManagedCommon;
using Microsoft.PowerToys.Settings.UI.Library;
using Wox.Plugin;
using Wox.Plugin.Logger;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Community.PowerToys.Run.Plugin.OfficeSearch;

namespace Community.PowerToys.Run.Plugin.OfficeSearch
{
    /// <summary>
    /// Main class of this plugin that implement all used interfaces.
    /// </summary>
    public class Main : IPlugin, IContextMenu, ISettingProvider, IDisposable
    {
        enum SupportedApplication
        {
            Excel,
            Word
        }

        private static Excel.RecentFiles ExcelFiles { get; } = (new Excel.Application()).RecentFiles;
        private static Word.RecentFiles WordFiles { get; } = (new Word.Application()).RecentFiles;

        /// <summary>
        /// ID of the plugin.
        /// </summary>
        public static string PluginID => "A2bd4fbacda4041f3a7fe22b33e79a99b";

        /// <summary>
        /// Name of the plugin.
        /// </summary>
        public string Name => "OfficeSearch";

        /// <summary>
        /// Description of the plugin.
        /// </summary>
        public string Description => "Searches recent excel sheets and word documents.";

        /// <summary>
        /// Additional options for the plugin.
        /// </summary>
        public IEnumerable<PluginAdditionalOption> AdditionalOptions => [
            new()
            {
                Key = nameof(IndexOneDrive),
                DisplayLabel = "Index OneDrive",
                DisplayDescription = "Index OneDrive files for search",
                PluginOptionType = PluginAdditionalOption.AdditionalOptionType.Checkbox,
                Value = IndexOneDrive,
            }
        ];

        private bool IndexOneDrive { get; set; }

        private PluginInitContext? Context { get; set; }

        private string ExcelLightIconPath { get; } = "Images\\ExcelLogoSmall.contrast-white_scale-180.png";
        private string ExcelDarkIconPath { get; } = "Images\\ExcelLogoSmall.contrast-black_scale-180.png";
        private string WordLightIconPath { get; } = "Images\\WordLogoSmall.contrast-white_scale-180.png";
        private string WordDarkIconPath { get; } = "Images\\WordLogoSmall.contrast-black_scale-180.png";

        private string? ExcelIconPath { get; set; }
        private string? WordIconPath { get; set; }

        private bool Disposed { get; set; }

        /// <summary>
        /// Return a filtered list, based on the given query.
        /// </summary>
        /// <param name="query">The query to filter the list.</param>
        /// <returns>A filtered list, can be empty when nothing was found.</returns>
        public List<Result> Query(Query query)
        {
            Log.Info("Query: " + query.Search, GetType());
            Log.Info("Word Recent: " + WordFiles.OfType<Word.RecentFile>().First().Name, GetType());
            Log.Info("Word Recent: " + WordFiles.OfType<Word.RecentFile>().First().Path, GetType());
            bool showQueryInSearch = false;

            return (from file in ExcelFiles.OfType<Excel.RecentFile>()
                    where File.Exists(file.Name) && file.Name.FuzzyBitap(query.Search, 2) > 5
                    orderby file.Name.FuzzyBitap(query.Search, 2) descending
                    select new Result()
                    {
                        QueryTextDisplay = query.Search,
                        IcoPath = ExcelIconPath,
                        Title =
                          showQueryInSearch ?
                            file.Name.EllipsifyInterpolatedQuery(query.Search) :
                            Path.GetFileNameWithoutExtension(file.Name),
                        SubTitle = $"Last modified {new FileInfo(file.Name).LastWriteTime.ToString("d")}",
                        ToolTipData = new ToolTipData("Excel Spreadsheet", $"File Path: {file.Name}"),
                        ContextData = (file.Name, SupportedApplication.Excel),
                    })
                    .Union(
                   from file in WordFiles.OfType<Word.RecentFile>()
                   where File.Exists(Path.Combine(file.Path, file.Name)) && Path.Combine(file.Path, file.Name).FuzzyBitap(query.Search, 2) > 5
                   orderby Path.Combine(file.Path, file.Name).FuzzyBitap(query.Search, 2) descending
                   select new Result()
                   {
                       QueryTextDisplay = query.Search,
                       IcoPath = WordIconPath,
                       Title =
                         showQueryInSearch ?
                           file.Name.EllipsifyInterpolatedQuery(query.Search) :
                           Path.GetFileNameWithoutExtension(file.Name),
                       SubTitle = $"Last modified {new FileInfo(Path.Combine(file.Path, file.Name)).LastWriteTime.ToString("d")}",
                       ToolTipData = new ToolTipData("Word Document", $"File Path: {Path.Combine(file.Path, file.Name)}"),
                       ContextData = (Path.Combine(file.Path, file.Name), SupportedApplication.Word),
                   }

                ).ToList();
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

            if (selectedResult?.ContextData is (string fileName, SupportedApplication app))
            {
                return
                [
                    new ContextMenuResult
                    {
                        PluginName = Name,
                        Title = "Open (Enter)",
                        FontFamily = "Segoe Fluent Icons,Segoe MDL2 Assets",
                        Glyph = "\xE8E5", // Open File Icon
                        AcceleratorKey = Key.Enter,
                        Action = _ => {
                            switch (app){
                                case SupportedApplication.Excel:
                                    Excel.Application excel = new(){ Visible = true };
                                    excel.Workbooks.Open(fileName);
                                    break;
                                case SupportedApplication.Word:
                                    Word.Application word = new(){ Visible = true };
                                    word.Documents.Open(fileName);
                                    break;
                            }
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

            IndexOneDrive = settings.AdditionalOptions.SingleOrDefault(x => x.Key == nameof(IndexOneDrive))?.Value ?? false;
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

        private void UpdateIconPath(Theme theme)
        {
            ExcelIconPath = theme == Theme.Light || theme == Theme.HighContrastWhite ? ExcelLightIconPath : ExcelDarkIconPath;
            WordIconPath = theme == Theme.Light || theme == Theme.HighContrastWhite ? WordLightIconPath : WordDarkIconPath;
        }

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
