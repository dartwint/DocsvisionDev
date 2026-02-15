using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

using DevExpress.Utils;
using DevExpress.Utils.Menu;
using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Menu;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraLayout;

using DocsVision.BackOffice.ObjectModel;
using DocsVision.BackOffice.ObjectModel.Services;
using DocsVision.BackOffice.WinForms;
using DocsVision.BackOffice.WinForms.Controls;
using DocsVision.BackOffice.WinForms.Design.LayoutItems;
using DocsVision.BackOffice.WinForms.Design.PropertyControls;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectModel;

using DvScriptsBase.Resolvers;

using Excel = Microsoft.Office.Interop.Excel;

namespace BackOffice
{
    [CardKindScriptClass]
    public class ReportOnDocumentsScript : CardDocumentScript
    {
        #region Fields

        private readonly SimpleLogger _logger;
        private bool _enableLogging = true;
        private static readonly string _workingDir = "DocsVision\\ReportOnMyCards";
        private readonly string _loggerTempFilePath = Path.Combine(Path.GetTempPath(), _workingDir, "ReportOnMyDocuments.log");
        private readonly string _loggerFilePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\" + _workingDir + "\\ReportOnMyDocuments.log";

        private LayoutItemHelper _layoutItemHelper;
        private bool _cardActivated = false;
        private bool _cardClosing = false;
        private const string frameCaption = "Форма поиска документов";

        private CardUrlHelper _cardUrlHelper;

        private XtraGridRepository _xtraGridRepository;
        private SqlServerExtensionResolver _sqlServerExtResolver;

        //private static int _panelWithFilesTableHeight;

        private const int excelExportBatchSize = 1000;

        private Size _tableFilesFileIconSize = new Size(16, 16);
        private int _tableFilesRowHeight = 20;

        private static class SessionPropertyNames
        {
            public const string ServerUrlSessionPropertyName = "ServerUrl";
            public const string IsAdminSessionPropertyName = "IsAdmin";
        }

        private Guid _userId;
        private DbSelectDocsOptions _currentSelectOptions;

        private static Process _currentProcess = Process.GetCurrentProcess();

        private System.Windows.Forms.Label _loadingOverlay;

        private static readonly DateTime _minDBDateTime = new DateTime(1900, 1, 1, 0, 0, 0);
        private const string _minDbDateStr = "1900-01-01T00:00:00";

        private int _tableFilesSearchPanelThreshold = 5;
        private int _defaultDocsLimit = 3000;

        private static readonly string _docsGridViewXml = Path.Combine(Path.GetTempPath(), _workingDir, "DocumentsGridView.xml");
        private static readonly string _filesGridViewXml = Path.Combine(Path.GetTempPath(), _workingDir, "FilesGridView.xml");

        private GridViewData.PersistenceOptions _gridDocsViewPersistenceOptions = GridViewData.PersistenceOptions.Session;
        private GridViewData.PersistenceOptions _gridFilesViewPersistenceOptions = GridViewData.PersistenceOptions.Session;

        private const string _SEDefaultName = "GAZServerExtension";
        private const string _extMethodName = "ExecuteTextSQLCommand";

        private Control _exportToExcelBtn, _searchBtn;
        private DateEditorButtonsHandler _dateEditorButtonsHandler = new DateEditorButtonsHandler();
        private Control _filePreviewCtrl;

        private UserProfileCardData _userProfileCardData;
        private static readonly Guid SettingObjectId = new Guid("3fe834de-7018-4b4f-99a6-020c6ebde055");
        //private const int ThisCardSettingsType = 90256251;
        private const int SettingId = 775114;
        private DataPersistenceContext _dataPersistenceContext = DataPersistenceContext.User;

        private IPersistantDataController _dataController;
        private ISerializationStrategy _serializationStrategy = new XmlSerializer();

        // misc
        [Obsolete] private string _pathToLayoutXML = "C:\\Users\\KokurinMR\\Documents\\Layout.xml";

        #endregion

        #region Extra Types

        private enum ImageAlias
        {
            /// <summary>
            /// Экспорт в Excel
            /// </summary>
            ToExcel,

            /// <summary>
            /// Иконка для архивного файла
            /// </summary>
            Archive,

            /// <summary>
            /// Иконка для аудиофайла
            /// </summary>
            Audio,

            /// <summary>
            /// Иконка для файла изображения
            /// </summary>
            Image,

            /// <summary>
            /// Иконка для видеофайла
            /// </summary>
            Video,

            /// <summary>
            /// Иконка для документов Microsoft Word
            /// </summary>
            Word,

            /// <summary>
            /// Иконка для документов Microsoft Excel
            /// </summary>
            Excel,

            /// <summary>
            /// Иконка для презентаций Microsoft PowerPoint
            /// </summary>
            PowerPoint,

            /// <summary>
            /// Иконка для pdf-документов
            /// </summary>
            Pdf,

            /// <summary>
            /// Иконка для файлов с расширением .txt
            /// </summary>
            Text,

            /// <summary>
            /// Иконка для файла с неопознанным расширением
            /// </summary>
            Undefined,

            /// <summary>
            /// Из встроенного контекстного меню при клике по колонке GridView 'Группировать'
            /// </summary>
            GroupColumn,

            /// <summary>
            /// Иконка настройки видимых столбцов
            /// </summary>
            ColumnsVisibility,

            /// <summary>
            /// Очистить группировку в GridView
            /// </summary>
            UngroupAll,

            /// <summary>
            /// Иконка автоподбора по ширине содержимого колонки
            /// </summary>
            BestFitColumn,

            /// <summary>
            /// Показать панель группировки в GridView
            /// </summary>
            ShowGroupPanel,

            /// <summary>
            /// Скрыть панель группировки в GridView
            /// </summary>
            HideGroupPanel,

            /// <summary>
            /// Развернуть все группы в GridView
            /// </summary>
            ExpandAllGroupLevels,

            /// <summary>
            /// Редактор фильтра колонки в GridView
            /// </summary>
            FilterEditor,

            /// <summary>
            /// Свернуть все группы в GridView
            /// </summary>
            CollapseAllGroupLevels,

            /// <summary>
            /// Отменить группировку
            /// </summary>
            ClearGrouping,

            /// <summary>
            /// Очистка всех фильтров
            /// </summary>
            ClearFilters,

            /// <summary>
            /// Отсортировать данные по возрастанию
            /// </summary>
            SortByAsc,

            /// <summary>
            /// Отсортировать по убыванию
            /// </summary>
            SortByDesc,
        }

        private interface IImageLoader<TImage> : IDisposable
        {
            string ImageData { get; }
            bool IsDataCompressed { get; }
            TImage CachedImage { get; }
            TImage Load();

            void ClearCache();
        }

        private abstract class BaseImageLoader<TImage> : IImageLoader<TImage>
        {
            public SimpleLogger logger = new SimpleLogger();

            public virtual string ImageData { get; protected set; }
            public virtual TImage CachedImage
            {
                get
                {
                    if (_cached == null && ImageData != null && !_disposed)
                    {
                        _cached = Load();
                    }

                    return _cached;
                }
            }

            public bool IsDataCompressed { get { return _isDataCompressed; } }

            protected TImage _cached;
            protected bool _isDataCompressed;
            protected bool _disposed = false;

            public BaseImageLoader(bool compressed, string data)
            {
                this._isDataCompressed = compressed;
                this.ImageData = data;
            }

            public abstract TImage Load();

            public virtual void ClearCache()
            {
                if (!_disposed && _cached != null)
                {
                    if (_cached is IDisposable)
                    {
                        IDisposable disposable = _cached as IDisposable;
                        disposable.Dispose();
                    }
                    _cached = default(TImage);
                }
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!_disposed)
                {
                    if (disposing)
                    {
                        ClearCache();
                        ImageData = null;
                    }

                    _disposed = true;
                }
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        private class ImageLoader : BaseImageLoader<System.Drawing.Image>
        {
            public ImageLoader(bool compressed, string data) : base(compressed, data)
            {

            }

            public override System.Drawing.Image Load()
            {
                return Utils.StringToImage(ImageData, IsDataCompressed);
            }
        }

        private class SvgImageLoader : BaseImageLoader<DevExpress.Utils.Svg.SvgImage>
        {
            public SvgImageLoader(bool compressed, string data) : base(compressed, data)
            {

            }

            public override DevExpress.Utils.Svg.SvgImage Load()
            {
                return Utils.StringToSvgImage(ImageData, IsDataCompressed);
            }
        }

        private class ImageManager
        {
            private Dictionary<ImageAlias, object> _loaders = new Dictionary<ImageAlias, object>();
            private Dictionary<string, ImageAlias> _fileExtensionIconCollection = new Dictionary<string, ImageAlias>();
            private SvgImageCollection _svgImages = new SvgImageCollection();

            public SvgImageCollection SvgImages { get { return _svgImages; } }

            public TImage Get<TImage>(ImageAlias alias)
            {
                if (!_loaders.ContainsKey(alias))
                {
                    return default(TImage);
                }

                var loader = _loaders[alias];

                if (loader is IImageLoader<TImage>)
                {
                    return ((IImageLoader<TImage>)loader).CachedImage;
                }

                return default(TImage);
            }

            public IImageLoader<TImage> GetLoader<TImage>(ImageAlias alias)
            {
                if (!_loaders.ContainsKey(alias))
                {
                    return default(IImageLoader<TImage>);
                }

                return (IImageLoader<TImage>)_loaders[alias];
            }

            public void Add<TImage>(ImageAlias alias, Func<IImageLoader<TImage>> loaderFactory)
            {
                if (!_loaders.ContainsKey(alias))
                {
                    if (loaderFactory == null)
                    {
                        throw new ArgumentNullException("Loader factory reference is not set");
                    }

                    _loaders[alias] = loaderFactory();
                }
            }

            public void AddSvgImage(ImageAlias alias, DevExpress.Utils.Svg.SvgImage image)
            {
                _svgImages.Add(alias.ToString(), image);
            }

            public void AddSvgImage(ImageAlias alias)
            {
                object loader;
                if (_loaders.TryGetValue(alias, out loader))
                {
                    if (loader == null)
                    {
                        throw new NullReferenceException("Image loader reference isn't set. ImageAlias=" + alias.ToString());
                    }

                    if (loader is SvgImageLoader)
                    {
                        SvgImageLoader svgImageLoader = loader as SvgImageLoader;
                        if (svgImageLoader != null)
                        {
                            _svgImages.Add(alias.ToString(), svgImageLoader.CachedImage);
                        }
                        else
                        {
                            throw new NullReferenceException("Loader object was casted to SvgImageLoader but it's reference isn't set");
                        }
                    }
                }
                else
                {
                    throw new ArgumentException("Image loaders collection hasn't item with ImageAlias=" + alias.ToString());
                }
            }

            public void BindFileExtension(string extension, ImageAlias alias)
            {
                if (!_fileExtensionIconCollection.ContainsKey(extension))
                {
                    _fileExtensionIconCollection.Add(extension, alias);
                }
                else
                {
                    _fileExtensionIconCollection[extension] = alias;
                }
            }

            public void BindFileExtensions(ImageAlias alias, params string[] extensions)
            {
                if (extensions != null)
                {
                    int n = extensions.Length;
                    for (int i = 0; i < n; i++)
                    {
                        if (!_fileExtensionIconCollection.ContainsKey(extensions[i]))
                        {
                            _fileExtensionIconCollection.Add(extensions[i], alias);
                        }
                        else
                        {
                            _fileExtensionIconCollection[extensions[i]] = alias;
                        }
                    }
                }
            }

            public ImageAlias GetAliasFromBindedFileExtensions(string extension)
            {
                if (_fileExtensionIconCollection.ContainsKey(extension))
                {
                    return _fileExtensionIconCollection[extension];
                }

                return ImageAlias.Undefined;
            }

            public ImageManager()
            {
                // DevExpress Svg Icons
                Add(ImageAlias.GroupColumn, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxnIG9wYWNpdHk9IjAuNSIgY2xhc3M9InN0MiI+DQogICAgICA8cGF0aCBkPSJNMjQsMjBIMTRjLTEuMSwwLTItMC45LTItMlY4YzAtMS4xLDAuOS0yLDItMmgxMGMxLjEsMCwyLDAuOSwyLDJ2MTBDMjYsMTkuMSwyNS4xLDIwLDI0LDIweiIgZmlsbD0iIzcyNzI3MiIgb3BhY2l0eT0iMC41IiBjbGFzcz0iQmxhY2siIC8+DQogICAgPC9nPg0KICAgIDxwYXRoIGQ9Ik0xOCwyNkg4Yy0xLjEsMC0yLTAuOS0yLTJWMTRjMC0xLjEsMC45LTIsMi0yaDEwYzEuMSwwLDIsMC45LDIsMnYxMEMyMCwyNS4xLDE5LjEsMjYsMTgsMjZ6IiBmaWxsPSIjMTE3N0Q3IiBjbGFzcz0iQmx1ZSIgLz4NCiAgICA8cGF0aCBkPSJNMiwyOXYtM0gwdjRjMCwxLjEsMC45LDIsMiwyaDR2LTJIM0MyLjQsMzAsMiwyOS42LDIsMjl6IiBmaWxsPSIjNzI3MjcyIiBjbGFzcz0ic3QzIiAvPg0KICAgIDxwYXRoIGQ9Ik0zLDJoM1YwSDJDMC45LDAsMCwwLjksMCwydjRoMlYzQzIsMi40LDIuNCwyLDMsMnoiIGZpbGw9IiM3MjcyNzIiIGNsYXNzPSJzdDMiIC8+DQogICAgPHBhdGggZD0iTTMwLDN2M2gyVjJjMC0xLjEtMC45LTItMi0yaC00djJoM0MyOS42LDIsMzAsMi40LDMwLDN6IiBmaWxsPSIjNzI3MjcyIiBjbGFzcz0ic3QzIiAvPg0KICAgIDxwYXRoIGQ9Ik0yOSwzMGgtM3YyaDRjMS4xLDAsMi0wLjksMi0ydi00aC0ydjNDMzAsMjkuNiwyOS42LDMwLDI5LDMweiIgZmlsbD0iIzcyNzI3MiIgY2xhc3M9InN0MyIgLz4NCiAgPC9nPg0KICA8ZyBpZD0iTGF5ZXJfMiIgLz4NCjwvc3ZnPg=="));
                Add(ImageAlias.ColumnsVisibility, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxnIG9wYWNpdHk9IjAuNSIgY2xhc3M9InN0MiI+DQogICAgICA8cGF0aCBkPSJNMiwwdjMwaDI4VjBIMnogTTI4LDI4SDRWMmgyNFYyOHoiIGZpbGw9IiM3MjcyNzIiIG9wYWNpdHk9IjAuNSIgY2xhc3M9IkJsYWNrIiAvPg0KICAgIDwvZz4NCiAgICA8cGF0aCBkPSJNMjUsMTBIN2MtMC42LDAtMS0wLjQtMS0xVjVjMC0wLjYsMC40LTEsMS0xaDE4YzAuNiwwLDEsMC40LDEsMXY0QzI2LDkuNiwyNS42LDEwLDI1LDEweiIgZmlsbD0iIzExNzdENyIgY2xhc3M9IkJsdWUiIC8+DQogICAgPHBhdGggZD0iTTI1LDE4SDdjLTAuNiwwLTEtMC40LTEtMXYtNGMwLTAuNiwwLjQtMSwxLTFoMThjMC42LDAsMSwwLjQsMSwxdjRDMjYsMTcuNiwyNS42LDE4LDI1LDE4eiIgZmlsbD0iIzExNzdENyIgY2xhc3M9IkJsdWUiIC8+DQogICAgPHBhdGggZD0iTTI1LDI2SDdjLTAuNiwwLTEtMC40LTEtMXYtNGMwLTAuNiwwLjQtMSwxLTFoMThjMC42LDAsMSwwLjQsMSwxdjRDMjYsMjUuNiwyNS42LDI2LDI1LDI2eiIgZmlsbD0iIzExNzdENyIgY2xhc3M9IkJsdWUiIC8+DQogIDwvZz4NCiAgPGcgaWQ9IkxheWVyXzIiIC8+DQo8L3N2Zz4="));
                Add(ImageAlias.UngroupAll, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxnIG9wYWNpdHk9IjAuNSIgY2xhc3M9InN0MiI+DQogICAgICA8cGF0aCBkPSJNMjgsMTZIMThjLTEuMSwwLTItMC45LTItMlY0YzAtMS4xLDAuOS0yLDItMmgxMGMxLjEsMCwyLDAuOSwyLDJ2MTBDMzAsMTUuMSwyOS4xLDE2LDI4LDE2eiIgZmlsbD0iIzcyNzI3MiIgb3BhY2l0eT0iMC41IiBjbGFzcz0iQmxhY2siIC8+DQogICAgPC9nPg0KICAgIDxwYXRoIGQ9Ik0xNCwzMEg0Yy0xLjEsMC0yLTAuOS0yLTJWMThjMC0xLjEsMC45LTIsMi0yaDEwYzEuMSwwLDIsMC45LDIsMnYxMEMxNiwyOS4xLDE1LjEsMzAsMTQsMzB6IiBmaWxsPSIjMTE3N0Q3IiBjbGFzcz0iQmx1ZSIgLz4NCiAgPC9nPg0KICA8ZyBpZD0iTGF5ZXJfMiIgLz4NCjwvc3ZnPg=="));
                Add(ImageAlias.BestFitColumn, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxwYXRoIGQ9Ik0xOC4zLDE5LjZoLTQuNUwxMywyMmgtM2w0LjQtMTJoMy4yTDIyLDIyaC0zTDE4LjMsMTkuNnogTTE0LjUsMTcuMmgzTDE2LDEyLjNMMTQuNSwxNy4yeiIgZmlsbD0iI0QxMUMxQyIgY2xhc3M9IlJlZCIgLz4NCiAgICA8cG9seWdvbiBwb2ludHM9IjQsMjIgMTAsMTYgNCwxMCA0LDE0IDAsMTQgMCwxOCA0LDE4ICAiIGZpbGw9IiM3MjcyNzIiIGNsYXNzPSJCbGFjayIgLz4NCiAgICA8cG9seWdvbiBwb2ludHM9IjI4LDEwIDIyLDE2IDI4LDIyIDI4LDE4IDMyLDE4IDMyLDE0IDI4LDE0ICAiIGZpbGw9IiM3MjcyNzIiIGNsYXNzPSJCbGFjayIgLz4NCiAgPC9nPg0KICA8ZyBpZD0iTGF5ZXJfMiIgLz4NCjwvc3ZnPg=="));
                Add(ImageAlias.ClearGrouping, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxnIG9wYWNpdHk9IjAuNSIgY2xhc3M9InN0MiI+DQogICAgICA8cGF0aCBkPSJNMjgsMTZIMThjLTEuMSwwLTItMC45LTItMlY0YzAtMS4xLDAuOS0yLDItMmgxMGMxLjEsMCwyLDAuOSwyLDJ2MTBDMzAsMTUuMSwyOS4xLDE2LDI4LDE2eiIgZmlsbD0iIzcyNzI3MiIgb3BhY2l0eT0iMC41IiBjbGFzcz0iQmxhY2siIC8+DQogICAgPC9nPg0KICAgIDxwYXRoIGQ9Ik0xNCwzMEg0Yy0xLjEsMC0yLTAuOS0yLTJWMThjMC0xLjEsMC45LTIsMi0yaDEwYzEuMSwwLDIsMC45LDIsMnYxMEMxNiwyOS4xLDE1LjEsMzAsMTQsMzB6IiBmaWxsPSIjMTE3N0Q3IiBjbGFzcz0iQmx1ZSIgLz4NCiAgPC9nPg0KICA8ZyBpZD0iTGF5ZXJfMiIgLz4NCjwvc3ZnPg=="));
                Add(ImageAlias.ClearFilters, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxwYXRoIGQ9Ik0yNC4xLDJIMmMwLDAsNi44LDcuOSwxMC4yLDExLjlMMjQuMSwyeiIgZmlsbD0iIzExNzdENyIgY2xhc3M9IkJsdWUiIC8+DQogICAgPHBvbHlnb24gcG9pbnRzPSIxNCwyNC4xIDE0LDI4IDE4LDI4IDE4LDIwLjEgICIgZmlsbD0iIzExNzdENyIgY2xhc3M9IkJsdWUiIC8+DQogICAgPHBhdGggZD0iTTIuNCwyOS42TDIuNCwyOS42YzAuNiwwLjYsMS40LDAuNiwyLDBMMzEuNiwyLjRjMC42LTAuNiwwLjYtMS40LDAtMmwwLDBjLTAuNi0wLjYtMS40LTAuNi0yLDBMMi40LDI3LjYgICBDMS45LDI4LjEsMS45LDI5LDIuNCwyOS42eiIgZmlsbD0iI0QxMUMxQyIgY2xhc3M9IlJlZCIgLz4NCiAgPC9nPg0KICA8ZyBpZD0iTGF5ZXJfMiIgLz4NCjwvc3ZnPg=="));

                AddSvgImage(ImageAlias.GroupColumn);
                AddSvgImage(ImageAlias.ColumnsVisibility);
                AddSvgImage(ImageAlias.UngroupAll);
                AddSvgImage(ImageAlias.BestFitColumn);
                AddSvgImage(ImageAlias.ClearGrouping);
                AddSvgImage(ImageAlias.ClearFilters);

                // WARN: same images
                Add(ImageAlias.ShowGroupPanel, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxwYXRoIGQ9Ik0yNiwzMEgyYy0xLjEsMC0yLTAuOS0yLTJWMTZjMC0xLjEsMC45LTIsMi0yaDI0YzEuMSwwLDIsMC45LDIsMnYxMkMyOCwyOS4xLDI3LjEsMzAsMjYsMzB6IiBmaWxsPSIjNzI3MjcyIiBvcGFjaXR5PSIwLjM1IiBjbGFzcz0ic3QwIiAvPg0KICAgIDxwYXRoIGQ9Ik0yOCwxMGgtOGMtMS4xLDAtMi0wLjktMi0yVjRjMC0xLjEsMC45LTIsMi0yaDhjMS4xLDAsMiwwLjksMiwydjRDMzAsOS4xLDI5LjEsMTAsMjgsMTB6IiBmaWxsPSIjMTE3N0Q3IiBjbGFzcz0iQmx1ZSIgLz4NCiAgICA8cGF0aCBkPSJNMTYsMThsLTYsOGwtNi04aDRjMC00LjYsMS44LTguNiwzLjktMTEuNWMwLjktMS4yLDIuOC0xLjEsMy41LDAuM2wwLDBjMC40LDAuNywwLjIsMS42LTAuMywyLjIgICBjLTEuOCwyLjItMy4xLDUuNi0zLjEsOUgxNnoiIGZpbGw9IiM3MjcyNzIiIGNsYXNzPSJCbGFjayIgLz4NCiAgPC9nPg0KICA8ZyBpZD0iTGF5ZXJfMiIgLz4NCjwvc3ZnPg=="));
                Add(ImageAlias.HideGroupPanel, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxwYXRoIGQ9Ik0yNiwzMEgyYy0xLjEsMC0yLTAuOS0yLTJWMTZjMC0xLjEsMC45LTIsMi0yaDI0YzEuMSwwLDIsMC45LDIsMnYxMkMyOCwyOS4xLDI3LjEsMzAsMjYsMzB6IiBmaWxsPSIjNzI3MjcyIiBvcGFjaXR5PSIwLjM1IiBjbGFzcz0ic3QwIiAvPg0KICAgIDxwYXRoIGQ9Ik0yOCwxMGgtOGMtMS4xLDAtMi0wLjktMi0yVjRjMC0xLjEsMC45LTIsMi0yaDhjMS4xLDAsMiwwLjksMiwydjRDMzAsOS4xLDI5LjEsMTAsMjgsMTB6IiBmaWxsPSIjMTE3N0Q3IiBjbGFzcz0iQmx1ZSIgLz4NCiAgICA8cGF0aCBkPSJNMTYsMThsLTYsOGwtNi04aDRjMC00LjYsMS44LTguNiwzLjktMTEuNWMwLjktMS4yLDIuOC0xLjEsMy41LDAuM2wwLDBjMC40LDAuNywwLjIsMS42LTAuMywyLjIgICBjLTEuOCwyLjItMy4xLDUuNi0zLjEsOUgxNnoiIGZpbGw9IiM3MjcyNzIiIGNsYXNzPSJCbGFjayIgLz4NCiAgPC9nPg0KICA8ZyBpZD0iTGF5ZXJfMiIgLz4NCjwvc3ZnPg=="));

                AddSvgImage(ImageAlias.ShowGroupPanel);
                AddSvgImage(ImageAlias.HideGroupPanel);

                Add(ImageAlias.CollapseAllGroupLevels, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxwb2x5Z29uIHBvaW50cz0iMTAsOCA2LDggNiw0IDQsNCA0LDggMCw4IDAsMTAgNCwxMCA0LDE0IDYsMTQgNiwxMCAxMCwxMCAgIiBmaWxsPSIjMTE3N0Q3IiBjbGFzcz0iQmx1ZSIgLz4NCiAgICA8cG9seWdvbiBwb2ludHM9IjEwLDIwIDYsMjAgNiwxNiA0LDE2IDQsMjAgMCwyMCAwLDIyIDQsMjIgNCwyNiA2LDI2IDYsMjIgMTAsMjIgICIgZmlsbD0iIzExNzdENyIgY2xhc3M9IkJsdWUiIC8+DQogICAgPGcgb3BhY2l0eT0iMC41IiBjbGFzcz0ic3QzIj4NCiAgICAgIDxyZWN0IHg9IjE0IiB5PSI4IiB3aWR0aD0iMTgiIGhlaWdodD0iMiIgcng9IjAiIHJ5PSIwIiBmaWxsPSIjNzI3MjcyIiBvcGFjaXR5PSIwLjUiIGNsYXNzPSJzdDEiIC8+DQogICAgICA8cmVjdCB4PSIxNCIgeT0iMjAiIHdpZHRoPSIxOCIgaGVpZ2h0PSIyIiByeD0iMCIgcnk9IjAiIGZpbGw9IiM3MjcyNzIiIG9wYWNpdHk9IjAuNSIgY2xhc3M9InN0MSIgLz4NCiAgICA8L2c+DQogIDwvZz4NCiAgPGcgaWQ9IkxheWVyXzIiIC8+DQo8L3N2Zz4="));
                Add(ImageAlias.ExpandAllGroupLevels, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxyZWN0IHg9IjAiIHk9IjgiIHdpZHRoPSIxMCIgaGVpZ2h0PSIyIiByeD0iMCIgcnk9IjAiIGZpbGw9IiMxMTc3RDciIGNsYXNzPSJCbHVlIiAvPg0KICAgIDxyZWN0IHg9IjYiIHk9IjIwIiB3aWR0aD0iMTAiIGhlaWdodD0iMiIgcng9IjAiIHJ5PSIwIiBmaWxsPSIjMTE3N0Q3IiBjbGFzcz0iQmx1ZSIgLz4NCiAgICA8ZyBvcGFjaXR5PSIwLjUiIGNsYXNzPSJzdDMiPg0KICAgICAgPHJlY3QgeD0iMTQiIHk9IjgiIHdpZHRoPSIxOCIgaGVpZ2h0PSIyIiByeD0iMCIgcnk9IjAiIGZpbGw9IiM3MjcyNzIiIG9wYWNpdHk9IjAuNSIgY2xhc3M9InN0MSIgLz4NCiAgICAgIDxyZWN0IHg9IjIwIiB5PSIyMCIgd2lkdGg9IjEyIiBoZWlnaHQ9IjIiIHJ4PSIwIiByeT0iMCIgZmlsbD0iIzcyNzI3MiIgb3BhY2l0eT0iMC41IiBjbGFzcz0ic3QxIiAvPg0KICAgIDwvZz4NCiAgPC9nPg0KICA8ZyBpZD0iTGF5ZXJfMiIgLz4NCjwvc3ZnPg=="));
                Add(ImageAlias.FilterEditor, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxwYXRoIGQ9Ik0yLDJoMjhMMTgsMTZ2MTJoLTRWMTZDMTQsMTYsMiwxLjksMiwyeiIgZmlsbD0iIzExNzdENyIgY2xhc3M9IkJsdWUiIC8+DQogIDwvZz4NCiAgPGcgaWQ9IkxheWVyXzIiIC8+DQo8L3N2Zz4="));
                Add(ImageAlias.SortByAsc, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxwb2x5Z29uIHBvaW50cz0iMTgsMjIgMjQsMzAgMzAsMjIgMjYsMjIgMjYsMiAyMiwyIDIyLDIyICAiIGZpbGw9IiM3MjcyNzIiIGNsYXNzPSJCbGFjayIgLz4NCiAgICA8cGF0aCBkPSJNMTAuMywxMS42SDUuN0w1LDE0SDJMNi40LDJoMy4yTDE0LDE0aC0zTDEwLjMsMTEuNnogTTYuNSw5LjJoM0w4LDQuM0w2LjUsOS4yeiIgZmlsbD0iI0QxMUMxQyIgY2xhc3M9IlJlZCIgLz4NCiAgICA8cGF0aCBkPSJNMiwxOGgxMnYyLjRsLTcuMiw3LjJIMTRWMzBIMnYtMi40bDcuMi03LjJIMlYxOHoiIGZpbGw9IiMxMTc3RDciIGNsYXNzPSJCbHVlIiAvPg0KICA8L2c+DQogIDxnIGlkPSJMYXllcl8yIiAvPg0KPC9zdmc+"));
                Add(ImageAlias.SortByDesc, () => new SvgImageLoader(false, @"PD94bWwgdmVyc2lvbj0nMS4wJyBlbmNvZGluZz0nVVRGLTgnPz4NCjxzdmcgeD0iMHB4IiB5PSIwcHgiIHZpZXdCb3g9IjAgMCAzMiAzMiIgdmVyc2lvbj0iMS4xIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB4bWw6c3BhY2U9InByZXNlcnZlIj4NCiAgPGcgaWQ9IkJhY2tncm91bmQiPg0KICAgIDxwb2x5Z29uIHBvaW50cz0iMTgsMjIgMjQsMzAgMzAsMjIgMjYsMjIgMjYsMiAyMiwyIDIyLDIyICAiIGZpbGw9IiM3MjcyNzIiIGNsYXNzPSJCbGFjayIgLz4NCiAgICA8cGF0aCBkPSJNMTAuMywyNy42SDUuN0w1LDMwSDJsNC40LTEyaDMuMkwxNCwzMGgtM0wxMC4zLDI3LjZ6IE02LjUsMjUuMmgzTDgsMjAuM0w2LjUsMjUuMnoiIGZpbGw9IiMxMTc3RDciIGNsYXNzPSJCbHVlIiAvPg0KICAgIDxwYXRoIGQ9Ik0yLDJoMTJ2Mi40bC03LjIsNy4ySDE0VjE0SDJ2LTIuNGw3LjItNy4ySDJWMnoiIGZpbGw9IiNEMTFDMUMiIGNsYXNzPSJSZWQiIC8+DQogIDwvZz4NCiAgPGcgaWQ9IkxheWVyXzIiIC8+DQo8L3N2Zz4="));

                AddSvgImage(ImageAlias.CollapseAllGroupLevels);
                AddSvgImage(ImageAlias.ExpandAllGroupLevels);
                AddSvgImage(ImageAlias.FilterEditor);
                AddSvgImage(ImageAlias.SortByAsc);
                AddSvgImage(ImageAlias.SortByDesc);

                Add(ImageAlias.ToExcel, () => new ImageLoader(false, @"iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAA7EAAAOxAGVKw4bAAACCUlEQVQ4jY2SX0hTURzHv3ezvdVDIBWshwpfxPVYD/3BIT0JlkgkbU5lbhT5IuiYFgp7EcbGkC2kQUJR+FiEDHpwq5EPRUKU0Es7FlzRZAmbeO+5u797jy/eG9Yc+8KXc/id8/udz+/wk3AgZ8SZAnARzemrETfGDkVuz8yERZNyjjsLVp50sIZLpdITxhgYYwiFQljf5Hi3lsZ25UcjkqK9uzU5ec8wTUGGYTu2ONSQJLY4JFosguTIyHx+eRmMMQwHg/i1qYEThyEMTL/0AwCGu6ZQWHuFn7+/I+Z7AU4cVoHsg0zG9SaRSF/r7IQpBM6edoETBwkCJ47gjUeYfzsNADh3qt2OOyyCx6Oj6ff5PJ4tLMA0TazLHKquggTB740gk3sIv3cCqq7iSns3SBBUXbULZIPx+NhVrxd3BwdhCAH3mWNQSIFu6lhafQ6FFMzlogh0RTGXi0I3dSik2C2En0YiqQ+FAhhjuBMIQN7QbQJ3axvcrW0AgNTSBADYBLYuDwyM79ZqYmdvz3ZPskOUqSy2+NZ/LlNZ9CQ7RN056PP5IMuE7Mp9sD+rRw6BBOnvHHh6e6MVTRPb1WpDVzRNOPocK1ae/QevE4nZT8UiGGO42d+PDdmo++r5Cy6gVufguMczWyUSO5w3dJVISNelj/8SYNf17eSJSy2fj2z4cPNfmrrXjPYBYu6kpqPyPGIAAAAASUVORK5CYII="));

                // File icons
                Add(ImageAlias.Archive, () => new ImageLoader(false, @"iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAA7EAAAOxAGVKw4bAAACB0lEQVQ4jXWST0gVURSHv7lvpoSkeqAQgYumZRChVrQKYva1ioKEINCdILhIqEUR2EIQgha2ClwkSVCbNs4ygshAiBYK3YLcpVH5tPHNPee2mHnznv8OHIb7m3t+850zJ6CMxcfcBk61zp8s2yPPWJsZpmcg5jA749vgXZ4DhC1FPafPj765h0hh8OInCwsx1loGbva2S2s1Pj65+qh1rAy8x+ByfJ4DEMcx9Xoday2+ebSqD6KouLvbQBSjIviSAKBerxd0HVpgDKL7GKjH+AMMOjWMQQ8i8CKouOru/Px8SeA66ncSmA6DUCVHXJHWWm5d+Fq8c23de0W0/eFOA6NOUOdQ54jjmIdzAXEcV5qJDrG6vIIoZuE+/XsIvDi0zIt9X0iSBGstKg5Tq7G6vIKKcOWBjopyZx+CnFZ++HGGNE2ZGOqptI3ffzjRd5KtZrOaQ2XghFCdQ8rs710kSRImZ9eQsgWRosWNLMNJMYdde+BQKRZp6dcl0jRl/PoRVHJUajhXtLe5vb2XoJiBoGWePfaOJEmYerlJa8GkfA5/763+REWQNeny2l6kz43LpGnK2LUIL4JXIcsyvArr61tkTbp2EPz9R7cJDFEYEYUR/cffkyQJ069zojDCBIZGo4EJDCOvut82Mp4CBC2DqSGmveccHTExFy5O3nCD7IogYGl8ljGA/+OnUA5erE9BAAAAAElFTkSuQmCC"));
                Add(ImageAlias.Audio, () => new ImageLoader(false, @"iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAB4UlEQVQ4jY2Tv2tTURTHP+e895oqxsEfQ6OgiC7t4D/gYuqmgyKCf4AgCK4OgRLaLk5OgnS0WKFU51YLboJbImZrEEJAJEhKyQsNefeHQ/KeLyQUDxzu4XL48P2ec68wjkfV6j2gxMnxZWd1tZm/CNPCOVfarFY3htaSWEviHNZ7vPcA1Ot1Nvf3nz5cWeHj+noG0QxgrTrnsEmCTRKGgwHH/T5xHNPr9YjjmA9raxvO2jsPKpXrUwBrLQKoCFEQMB9FzIUhoSoiQrfb5fXWFuXFxTeNWu3xtAVrVUVQEQCiIPhn1BjK5TLee4rFIp/q9e7B7u6kAmOtyhiQZhQEvKid58gWMiVp77QFYzIL+QR4+/MsXqMMYo1hFmBkQXUC0G4bOq1fvPtxmrkwJBgBphUYYzRTkIM0m4ZXd8/x/ivMRxFREGBygHACIIKIjKiq4BwLpQJDN+Dq5QJRYAlUZwNskmQKHNDoeL534OIFQ6OrVG4PUZGRhSQ5QYH3qAjPtg9YuHaJ3602y/dv4LzHeUHGvbNmkG1BgCv6h297n3lyM5warMltIVNw3O+HQDaD7ee38IBL/8N4JpLrnQAcHR6eSuv0QXnvkfGpgFNFVCd6M0A/jltnlpZe8h8hqq20/gtX0eMRZyybMQAAAABJRU5ErkJggg=="));
                Add(ImageAlias.Image, () => new ImageLoader(false, @"iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAA7EAAAOxAGVKw4bAAABvUlEQVQ4jcWSO2tUURSFvzlz7mPuMzdmDDOTKBlBixETLewEQQsLGxEFext/REiRVmytBFsL9Q/YamljF2Kp2DiZGCfX3PPYFhcUOxkEV7c3fIu12Qv+tzr3dnZuA8MF+c/aObf2cnf36SL03e3tR9pZixfh2pMOqgNBF3QXGgsiUPSgiEErGK/AZADWw/0rgrMWba1VToTNtQ6hhrIHKyl8+Qa9oJ0FmB3DxSFsjuC4ASeCtVZpa4xyIgzTt5wfbFClQ0Tgwiqs5q1R3cD4VJvkxLamTgRrTGtgvefN+2e8ixNuXb7D9ckNtIJQt7FrA8sJGNfukgCs962BMUZZEfb291gql3hx9JzDo0+s988wWF7ndFHRC2PoJC0ctqf98IIxRmnTNMp5z8nhiK/fNfNpyKvZB3rxPmmSkqUZWZaRZzllXrDeH7LRH3H1bIRpmt8JxoObhGFIFEUkSUKeZ5RlTlUVVFVOWaYURUyaBqhIYaX+lUB77//4r4hgraOuG5SqsVYxnwsHB444jomiiEsTj2karWfTadgNAl4/fvCX9fFATTcImU2nYafc2nrovT+3SBOVUh8X4f6tfgJGLLpjAOYoMwAAAABJRU5ErkJggg=="));
                Add(ImageAlias.Excel, () => new ImageLoader(false, @"iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAA7EAAAOxAGVKw4bAAACJElEQVQ4jX2Sz0tUURTHP+e+N1GQtAgJxEUt3IztpBYGgiMtrFZFMKREEhH9AYpk42CGtSoMQVoF0kAU4UYUMYraByMELXKRyzbaD33vNffed1rMzGMkprO5Xy6Hz/2ce6/QqKvl8iWgi//Xu9czM1utG2EzpGnatVQuP6t5j/Uem6Z4VVQVgGq1ytLGxu0rpRJvZmcziDx8desDMNDuyOPHTtJ/+gabm5tcGx7m8tTUHeDt8tzcFkBovR0oFZ+3dZ59OYaIsLOzw0KlQiGfX3xaqZSABwBh4hK8eqYrowCMDd3l/edlvn3/wv2RFyQuITSGQqGAqtLR0cF6tbrzdW2NDODUkbiEm+fvsbg+DcCpE/ls/1AYgnO4NAXAeW+ahia2MU4do4MTLKxOMTo4TmxjzuUv4tQR25jDuRyHwpDQGEQE71w2oolchE0tK5+WiFzE/Ook14cmmV+dxKaWyEXkgiCDBHVAZhA2Dbo7e+ju7AHgycp4XbVhYETIBQEAQRDgWgGJS5BAuHB2JNNqZgmExCUYEQByQUBgzEGA9fbj6OP+tv+gt/tMBgDqI1hr/mnsKxYnUlV13qv1Xv84p4m1GtVqupck+iuO9UcU6e9aTfuKxYnMoBmccwhgREhV6ytgAIyBxhMaEVzLK2SAeH8/BBARDLSFSEvvAcDP3d0jzSwiGBFUFWmsBkiNQYw50JsB9vf2to/29j5qd5mtJcZsN/Nf2FgZnVwEUE8AAAAASUVORK5CYII="));
                Add(ImageAlias.PowerPoint, () => new ImageLoader(false, @"iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAA7EAAAOxAGVKw4bAAACK0lEQVQ4jX2SP2hTURSHv3vfe5oWsxSh0hZEdJBWK1gRHOyQQSg4Ka3BSVxEcLVLDSFUqrMI4lqJCKKORSqCIrg4JN3EVik6aqo1zXt5989xSBMTJJ7l/Lic+53fPecqdmO2WDwPjPD/eP20VFrvPgjbwns/slwsPkydwziH8R4ngogAUKlUWF5dvXaxUODZ4mIHomp3Tr4Bpvu1bB44zo9zS1SrVS7PzHBhYeE68OrF0tI6QOiMnd5fqPb1/H3xBEoparUa98tlcuPjD+6VywXgNkBoE4N4h619Y/v5LVQUEQ4fZeD0JcKhMWxiCLUml8shImSzWV5WKrVPKysAaJukiLdsPZkn/viezJkrZM5epf7uEeItNknZE4aEWqOUAsA6p9sOtYmbiLU0NqqYuIkeO0HyuUL8ZQ2xFhM3yURRD8RZ+3cLtpHgnUEPHaK58YGvN44AsPfwKbwz2EZCFASdC0EL0O0gRZwlO1ciGJ3ESUgwOkl2roQ4i4lTtFJEQUAmioiCANsFCG2SEmjF4PBBBm8+7tlAoBU2aQEAoiAg0LoX4Ix9uzY73PcfDByb7gA6TzBG/1M4lc/PexGxzolxTprWSmKMNNJU6kki23EsPxsN+Z2mMpXPz3cctIW1FgVopfAirQxoAK3B+9bQlMJ2b6Et4p2dEEAphYa+ENVV2wP4tbU10NZKKbRSiAhqN2vAa43Suqe2A9ip1zf3TUzc7TfM7lBab7b1H1gHFkWGOE0kAAAAAElFTkSuQmCC"));
                Add(ImageAlias.Word, () => new ImageLoader(false, @"iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAA7EAAAOxAGVKw4bAAACMElEQVQ4jX2ST0hUURTGf+e+98KiWShCYBBKBKMuMiQQAhdDEPYHogwGV+2iXS2SQGQQRaRNEVFEi0CYICJqEyJGpAtxVeMiCHTjQlc19mdm3nPuve+2mJnXSExn831cvvvd75xzhXpdy+UuAl38vz68mprabD7wGySO4675XO5p1Vq0teg4xjqHcw6AQqHA/NLSjauTk7yenk5MZPTO0jLCcKsnu48e4vpoN+vr64yNjHBlYuIm8P7N7OwmgK+rdvjtg3MtM1++tYiIUCwWeZTPk+nre/Iwn58EZgD8KDQUvn7n7v01xi6c4MW7jQTnbg8RhQZfKTKZDM45UqkUi4VCcWNhAQAVlQ29Pe1UQ8ul4WM44xLs7WknKhsO+D6+UogIAMZalQwxLGuMcagYPn/5xvGuFJ/qaIwjLGvagqCmNgYRwRrzdwuVkkHrmEAJH1e3OZXuZHl1m4F0J1rHVEqGwPOSC17NIEmgGgkCUays7XAy3cHK2g4D6Y4kgRIh8DzagoDA8zBNBn5UMYh4PLt3Nnnl5ePztR2LR1QxqHrvgefhKbXfQO/Zld4zz1v+g6HTRxKDpAWt1T/CwWx2PHbOGWudttbtGeMirV2lWnWlKHK/wtD9qFTc72rVDWaz40mCBjHGIIASIXauhoACUAriuDY0EUzzFhokLJf9Wt+CgpYm0qTdZ/Bzd/dgg4sISgTnHFJHBcRKIUrt0yYG5VJp63B//1yrYTaXKLXV4H8AQqIUi10+Wo8AAAAASUVORK5CYII="));
                Add(ImageAlias.Pdf, () => new ImageLoader(false, @"iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAA7EAAAOxAGVKw4bAAACHUlEQVQ4jX2SPUtcQRSGn5l7VyOJIUUCQQIGSWGwtAlWKpZplhDY3iIpLWSLBV3EFGkiki6tIAghhRYuUdgfEIhskYWEtdCAEhNYNbof996ZOSmcvdklHweG83KYec57Zkbh42mx+BgY4v9Rfru0tN9dCDvCOTe0Viy+ia0lsZbEOawIIgJApVJhbXf32ZOFBd4tL/+G7MGrzyA1kH2QLyBVkArIB5CPi4vy6ehI1re3RUQkWyg8zxYKDzrnVRXkJtAPaCAB2kALaABN4PbxMeVyOW36en19oVYqvQAIDWABA+hMhrtRxNeREdzBAc7XQ62Znp5GRBgcHOR9pVKvlUoAaAPEvqsdHSWpVumbmSHx9QToC0NCrVFKAWCs1R03OuqyHExM8GNlhVuzs/RPThJ5yLVMpgdijUnH0W0/60A2y518nutTU+AcDzc2uDc/TwRkgiCFBFeA1EHY8jYfra7ybXOT7zs71Pf2uDg5IbGWNqCVIhMEAARBgOkGREAEbA0PAyD+Up3XsQfgnQRa9wB0DCtN4Bw4Beo+n/l1f24OrVS6AqWwSZIC0hjP5fJORIy1klgrkTHSThJpxrFcttvys9WSs2ZTLuJYxnO5fDpCRxhjUN6uE7nK/nOhNTh3ZVkpTNcrpIBWoxECKKXQ8E+I6trbAzg/PR3oaOXnFRGUzxpwWqO07tmbAhqXl4c3xsZe/nE5fwml9WFH/wJXfgp0DRVBLwAAAABJRU5ErkJggg=="));
                Add(ImageAlias.Text, () => new ImageLoader(false, @"iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAB20lEQVQ4jY2SwUocQRCGvyp3o4HIzp4WTAhB9xAUT77B4jGXEAJ5g7yDIIt6yQMEglfBU5KzBCGnnHUPuSW4+gCOBGdnxemuysGZYZfVkIKmi+bvr/6uaqGMt/3+K2CJf8f3zzs7vycPGlViZksH/f7+bYwUMVKYEd1xdwAGgwEHx8fv32xv83Vvr4ZoDYhRzYxYFMSi4PbmhvFoRJZlXF9fk2UZX3Z39y3GzddbW90ZQIwRAVSE5twcC80mjxoNGqqICGma8vHwkN7q6qefp6fvZp8Qo6oIP05OcHeSJCHPc152uxACvV4Pd2dxcZFvg0H66+hoGhBiVBGh2+1iZrg77XabhWazFASCWa2dcRBDQIDhcEiSJFxeXuLunJWNbLVajMdj1tfXiSHMTiGGoCrCyspKXb1yUi0zY06EGMKsgxCCCnB+fl5frJwkSYK7k+c5TzsdwoOACQeTfZjMm6r3A2JR6GQP0jTFzKaqz8/P86LTIRbFww6Wl5cBaLfbuDvPWi3MDHPH3JFSex+gnkJlN89znicJqEI5QhUh3DeF8WjUANjc2LjrgTsOmPvddy0hMqGdAvy5unpc5SKCiuDuSLkrYKqI6pS2Boyy7OLJ2toH/iNE9aLK/wKaGjMmY1HvJgAAAABJRU5ErkJggg=="));
                Add(ImageAlias.Undefined, () => new ImageLoader(false, @"iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAA7EAAAOxAGVKw4bAAABYklEQVQ4jX2SMY7bMBBF34xoyVp3glt3Bnwhn2EB1z6DT5CtUuUI8UHcBDqDjQW2MgIj0EozKVaUGFnIAIRG5Mzj/yQF4HA4hOPx+Jplmbo77q7uDoC7IyKk/2bGfr9/q+vaAsBms3mpquqbiGgsiA0AIoKIoKqICGbGbrf7Xtf1HwUoy1Jjca9g2CmFxa+IsN1uA4ACFEUxAKagaR5jvV6PgBCCxnwuRORprqqqABAA8jwPaVH0PDcX89VqNQLKstR0EUBVhxuYgxVFMQJ6C09yU+h05HmuA0BEdOozvftpc79p+C9gTk0KyLJsVBC9zu0+tfJVP66HvmbWwnMIDnTJ4wq9nDDXkJ4BgDl0bYe70TTNaEFVQypz2ghg9tUMRggZl8vl9wBYLBazhxijM6dtW8BYLpdcr9dfp9PpYwA8Ho/77Xb74e4G/DPcsa4za9tPK4q8ud/v7+fz+WfTNA3AX0Q0tuyOmYugAAAAAElFTkSuQmCC"));
                Add(ImageAlias.Video, () => new ImageLoader(false, @"iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAA7EAAAOxAGVKw4bAAAB3ElEQVQ4jYWRPU9UQRSGn1mtbeigk9XKBBPtDAlhC3VRkESJYSsRExO1Nf4EYylKUBErP7IJrJCwSiHRwk4TSexc7KCjsePO+bCYe+9KwXqaeWcy5zkfbwAYm274sb4+AK7MPqRXvF98AMCfvT3W37wORwEky3g79zhcv3PXW8/v9wQAvHv6JJy/es0BAkBtfML7+gcAmLj1qHgG/B+d7qsvUoG93R0+ra2GCkDMMpoL88FUUAM1x9STVkfN06lgKjQX5kPMMgAqaYR9Jmduuoqgkj6LeqmL5ARSJm/MuGT7XYDGSGvpZXDV9LlItAIEooaKYyK0Xi0FjbELiBK5PN1wUSmri4CUHVh6zwGXGg2PkgABYOjsGd/69j3Up6b8vxYA7WYzFDkB4NTpIR84XgXg4u3Fg79zI4In+eHZLAA7vzv8/LGVXDBVNlaWg0pMbUuaV6KXO4n5Mk0iGyvLwVS7OzBVRutjbpLvQCgTRCyHen4qo/W6HwSYsdleDyraXVws3OAAwCSy2W4HM+sC3Izh0ZqbSreSOmKe7LOk0zjKcK3mngMCQPVE1Tu/OuHcyEjZ2mFROVLh6+cvocgJAIPVwdK+C/fWegI+zo2XeruzHQJwEujvmXV47P4FHEmG5zZCUJIAAAAASUVORK5CYII="));

                BindFileExtensions(ImageAlias.Word, "doc", "docx");
                BindFileExtension("txt", ImageAlias.Text);
                BindFileExtensions(ImageAlias.Image, "bmp", "png", "jpg", "jpeg", "ico", "tiff", "svg", "gif");
                BindFileExtensions(ImageAlias.Video, "mp4", "mkv");
                BindFileExtensions(ImageAlias.Audio, "mp3", "wav");
                BindFileExtensions(ImageAlias.Archive, "rar", "tar", "7z", "7zip", "zip");
                BindFileExtensions(ImageAlias.Excel, "xls", "xlsx", "csv");
                BindFileExtensions(ImageAlias.PowerPoint, "ppt", "pptx");
                BindFileExtension("pdf", ImageAlias.Pdf);
            }

            public void Dispose(ImageAlias alias)
            {
                object loader;
                if (_loaders.TryGetValue(alias, out loader))
                {
                    if (loader is IDisposable)
                    {
                        IDisposable disposable = loader as IDisposable;
                        disposable.Dispose();
                    }

                    _loaders.Remove(alias);
                }
            }

            public void Dispose()
            {
                foreach (var loader in _loaders.Values)
                {
                    if (loader is IDisposable)
                    {
                        IDisposable disposable = loader as IDisposable;
                        disposable.Dispose();
                    }
                }

                _loaders.Clear();

                _svgImages.Clear();
                _svgImages.Dispose();

                _fileExtensionIconCollection.Clear();
            }
        }

        public class SimpleLogger
        {
            public string filePath;
            public string name;

            private bool _enableLogging;

            //private bool _hasAccessToWriteToFile;
            private readonly object _fileLock = new object();

            public bool EnableLogging
            {
                get { return _enableLogging; }
                set
                {
                    if (value == _enableLogging) return;

                    if (!value)
                    {
                        _enableLogging = false;
                        return;
                    }

                    if (EnsureDirectoryExists() && CheckWriteAccessOnExistingFile()) _enableLogging = value;
                }
            }

            public bool AbleToLog
            {
                get { return _enableLogging && EnsureDirectoryExists() && CheckWriteAccessOnExistingFile(); }
            }

            public enum LogLevel { Info, Debug, Trace, Warn, Error, Fatal }

            public SimpleLogger()
            {
                _enableLogging = false;
            }

            public SimpleLogger(string name, string filePath, bool enableLogging)
            {
                this.name = name;
                this.filePath = filePath;
                this._enableLogging = enableLogging;

                EnsureDirectoryExists();
                //_hasAccessToWriteToFile = CheckWriteAccessOnExistingFile();
            }

            private bool EnsureDirectoryExists()
            {
                try
                {
                    string directory = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    return true;
                }
                catch
                {
                    return false;
                }
            }

            private bool CheckWriteAccessOnExistingFile()
            {
                try
                {
                    using (var fs = new FileStream(filePath, FileMode.Append, FileAccess.Write))
                    {
                        return true;
                    }
                }
                catch
                {
                    return false;
                }
            }

            private void LogBase(string content)
            {
                if (!AbleToLog)
                {
                    return;
                }

                string logEntry = string.Format("{0}|{1}|{2}{3}",
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff"),
                    name,
                    content,
                    Environment.NewLine);

                lock (_fileLock)
                {
                    try
                    {
                        System.IO.File.AppendAllText(filePath, logEntry);
                    }
                    catch (IOException ex)
                    {
                        _enableLogging = false;

                        if (IsFileLocked(ex))
                        {

                        }
                        //Console.Error.WriteLine($"Файл {filePath} заблокирован другим процессом, лог пропущен");
                    }
                    catch (UnauthorizedAccessException)
                    {
                        _enableLogging = false;
                        //Console.Error.WriteLine($"Нет прав на запись в {filePath}, логирование отключено");
                    }
                    catch (Exception ex)
                    {
                        _enableLogging = false;
                        //Console.Error.WriteLine($"Ошибка логирования: {ex.Message}");
                    }
                }
            }

            private static bool IsFileLocked(IOException exception)
            {
                int errorCode = Marshal.GetHRForException(exception) & 0xFFFF;
                return errorCode == 32;
            }


            private static string GetLogLevelDisplayString(LogLevel level)
            {
                switch (level)
                {
                    case LogLevel.Info: return "INFO";
                    case LogLevel.Debug: return "DEBUG";
                    case LogLevel.Trace: return "TRACE";
                    case LogLevel.Warn: return "WARN";
                    case LogLevel.Error: return "ERROR";
                    case LogLevel.Fatal: return "FATAL";
                    default: return "";
                }
            }

            private void LogBase(LogLevel level, string content)
            {
                if (!AbleToLog)
                {
                    return;
                }

                string logEntry = string.Format("{0}|{1}|{2}|{3}{4}",
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.ffff"),
                    GetLogLevelDisplayString(level),
                    name,
                    content,
                    Environment.NewLine);

                lock (_fileLock)
                {
                    System.IO.File.AppendAllText(filePath, logEntry);
                }
            }

            public void Log(string content)
            {
                LogBase(content);
            }

            public void Log(LogLevel level, string content)
            {
                LogBase(level, content);
            }

            public void Info(string content)
            {
                LogBase(LogLevel.Info, content);
            }

            public void Debug(string content)
            {
                LogBase(LogLevel.Debug, content);
            }

            public void Trace(string content)
            {
                LogBase(LogLevel.Trace, content);
            }

            public void Warn(string content)
            {
                LogBase(LogLevel.Warn, content);
            }

            public void Error(string content)
            {
                LogBase(LogLevel.Error, content);
            }

            public void Fatal(string content)
            {
                LogBase(LogLevel.Fatal, content);
            }
        }

        private enum DataPersistenceContext
        {
            None, User, Group
        }

        private struct DbSelectDocsOptions
        {
            public Guid userId;
            public string nameContains;
            public int offset;
            public int limit;
            public int creationLastDays;
            public int optionsMask;
            public DateTime createdFrom, createdTo;
            public DateTime regDateFrom, regDateTo;
        }

        /// <summary>
        /// Содержит псевдонимы элементов разметки данной карточки
        /// </summary>
        private enum LayoutItemAlias
        {
            /// <summary>
            /// Элемент для предпросмотра файлов
            /// </summary>
            FilePreview,

            /// <summary>
            /// Таблица с файлами
            /// </summary>
            TableFiles,

            /// <summary>
            /// Таблица с документами
            /// </summary>
            TableDocuments,

            /// <summary>
            /// Поле "Дата создания с"
            /// </summary>
            CreationDateFrom,

            /// <summary>
            /// Поле "Дата создания по"
            /// </summary>
            CreationDateTo,

            /// <summary>
            /// Поле с целым числом. Хранит лимит строк для запроса к БД по документам.
            /// </summary>
            DocumentsLimit,

            /// <summary>
            /// Поле с сотрудником, по которому осуществляется поиск документов
            /// </summary>
            AuthorSelector,

            /// <summary>
            /// Кнопка, выполняющая подстановку текущего сотрудника в поле с автором
            /// </summary>
            MeButton,

            /// <summary>
            /// Поле с заголовком элемента предпросмотра файла
            /// </summary>
            PreviewableFileHeader,

            /// <summary>
            /// Элемент для получения доступа к группе, содержащей таблицу с файлами. НЕ предназначен для пользователя
            /// </summary>
            FilesGroupWrapper,

            /// <summary>
            /// Элемент
            /// </summary>
            FilePreviewGroupWrapper,

            /// <summary>
            /// Кнопка, инициирующая экспорт данных из таблицы с документами в Excel
            /// </summary>
            TableDocumentsExportButton,

            /// <summary>
            /// Кнопка, инициирующая выполнение запроса к БД и загрузку данных в таблицу с документами
            /// </summary>
            SearchButton,

            /// <summary>
            /// Кнопка, настраивающая видимость панели группировки для таблицы с документами
            /// </summary>
            GroupPanelVisibilitySwitch,

            /// <summary>
            /// Кнопка для сброса группировки в таблице (представлении) с документами
            /// </summary>
            ClearGroupingButton,

            /// <summary>
            /// Кнопка, настраивающая видимость колонок таблицы (представления) с документами
            /// </summary>
            ColumnVisibilitySwitch,

            /// <summary>
            /// Кнопка для с(раз-)сворачивания всех групп в таблице (представлении) с документами
            /// </summary>
            CollapseExpandGroupsButton,

            /// <summary>
            /// Кнопки сброса фильтра для всех колонок в таблице (представлении) с документами
            /// </summary>
            ClearFiltersButton,
        }

        public interface ILayoutItemRepository<TKey>
        {
            Dictionary<TKey, string> Items { get; }
            string this[TKey key] { get; }
            void RegisterNew(TKey key, string name);
        }

        public class LayoutItemRepository<TItemKey> : ILayoutItemRepository<TItemKey>
        {
            public Dictionary<TItemKey, string> Items { get; private set; }

            public string this[TItemKey key]
            {
                get
                {
                    string name;
                    if (Items.TryGetValue(key, out name))
                    {
                        return name;
                    }

                    throw new NullReferenceException("Items collection doesn't has item with key '" + key + "'");
                }
            }

            public LayoutItemRepository()
            {
                Items = new Dictionary<TItemKey, string>();
            }

            public void RegisterNew(TItemKey key, string name)
            {
                Items.Add(key, name);
            }
        }

        public interface ILayoutItemFinder
        {
            ICustomizableControl CustomizableControl { get; }
            DevExpress.XtraLayout.BaseLayoutItem GetItem(string name);
            TItem GetItem<TItem>(string name) where TItem : DevExpress.XtraLayout.BaseLayoutItem;
            System.Windows.Forms.Control GetControl(string name);
            TControl GetControl<TControl>(string name) where TControl : System.Windows.Forms.Control;
            TProperty FindPropertyItem<TProperty>(string name);
        }

        public interface ISafeLayoutItemFinder : ILayoutItemFinder
        {
            bool TryGetItem(string name, out DevExpress.XtraLayout.BaseLayoutItem item);
            bool TryGetItem<TItem>(string name, out DevExpress.XtraLayout.BaseLayoutItem item) where TItem : DevExpress.XtraLayout.BaseLayoutItem;
            bool TryGetControl(string name, out System.Windows.Forms.Control control);
            bool TryGetControl<TControl>(string name, out System.Windows.Forms.Control control) where TControl : System.Windows.Forms.Control;
        }

        public class BaseLayoutItemFinder : ISafeLayoutItemFinder
        {
            public ICustomizableControl CustomizableControl { get; private set; }

            public BaseLayoutItemFinder(ICustomizableControl customizableControl)
            {
                if (customizableControl == null)
                {
                    throw new ArgumentNullException(typeof(ICustomizableControl).Name + " reference is not set");
                }

                CustomizableControl = customizableControl;
            }

            public TProperty FindPropertyItem<TProperty>(string name)
            {
                return CustomizableControl.FindPropertyItem<TProperty>(name);
            }

            public Control GetControl(string name)
            {
                return CustomizableControl.LayoutControl.GetControlByName(name);
            }

            public TControl GetControl<TControl>(string name) where TControl : Control
            {
                return CustomizableControl.LayoutControl.GetControlByName(name) as TControl;
            }

            public BaseLayoutItem GetItem(string name)
            {
                return CustomizableControl.FindLayoutItem(name);
            }

            public TItem GetItem<TItem>(string name) where TItem : BaseLayoutItem
            {
                return CustomizableControl.FindPropertyItem<TItem>(name);
            }

            public bool TryGetItem(string name, out BaseLayoutItem item)
            {
                item = GetItem(name);
                return item != null;
            }

            public bool TryGetItem<TItem>(string name, out BaseLayoutItem item) where TItem : BaseLayoutItem
            {
                item = GetItem<TItem>(name);
                return item != null;
            }

            public bool TryGetControl(string name, out Control control)
            {
                control = GetControl(name);
                return control != null;
            }

            public bool TryGetControl<TControl>(string name, out Control control) where TControl : Control
            {
                control = GetControl<TControl>(name);
                return control != null;
            }
        }

        public interface ILayoutItemHelper<TItemKey>
        {
            ILayoutItemRepository<TItemKey> Repository { get; }
            BaseLayoutItemFinder ItemFinder { get; }
        }

        public static class LayoutItemValidator
        {
            public static bool IsItemExist(ICustomizableControl customizableControl, string itemName)
            {
                DevExpress.XtraLayout.BaseLayoutItem item = customizableControl.FindLayoutItem(itemName);
                return item != null;
            }

            public static bool IsControlExist(ICustomizableControl customizableControl, string controlName)
            {
                System.Windows.Forms.Control control = customizableControl.LayoutControl.GetControlByName(controlName);
                return control != null;
            }

            public static bool IsItemExist(LayoutGroup groupItem, string name)
            {
                return groupItem.Items.Where((item) => { return item.Name == name; }).Any();
            }
        }

        public class BaseLayoutItemHelper<TItemKey> : ILayoutItemHelper<TItemKey>
        {
            public ILayoutItemRepository<TItemKey> Repository { get; protected set; }
            public BaseLayoutItemFinder ItemFinder { get; protected set; }

            public BaseLayoutItemHelper(ILayoutItemRepository<TItemKey> itemsRepository, BaseLayoutItemFinder layoutItemFinder)
            {
                if (itemsRepository == null)
                {
                    throw new ArgumentNullException(typeof(ILayoutItemRepository<TItemKey>).Name + " reference is not set");
                }
                if (layoutItemFinder == null)
                {
                    throw new ArgumentNullException(typeof(BaseLayoutItemFinder).Name + " reference is not set");
                }

                Repository = itemsRepository;
                ItemFinder = layoutItemFinder;
            }

            public TProperty FindPropertyItem<TProperty>(TItemKey key)
            {
                return ItemFinder.FindPropertyItem<TProperty>(Repository[key]);
            }
        }

        /// <summary>
        /// Предоставляет интерфейс для получения элементов разметки данной карточки
        /// </summary>
        private sealed class LayoutItemHelper : BaseLayoutItemHelper<LayoutItemAlias>
        {
            public LayoutItemHelper(ILayoutItemRepository<LayoutItemAlias> itemsRepository, BaseLayoutItemFinder layoutItemFinder) : base(itemsRepository, layoutItemFinder) { }

            public static LayoutItemHelper BuildDefault(ICustomizableControl customizableControl)
            {
                if (customizableControl == null)
                {
                    throw new ArgumentNullException("ICustomizableControl reference is not set");
                }

                ILayoutItemRepository<LayoutItemAlias> itemsRepo = new LayoutItemRepository<LayoutItemAlias>();
                itemsRepo.RegisterNew(LayoutItemAlias.SearchButton, "Search");
                itemsRepo.RegisterNew(LayoutItemAlias.FilePreview, "FilePreview");
                itemsRepo.RegisterNew(LayoutItemAlias.TableFiles, "FilesTable");
                itemsRepo.RegisterNew(LayoutItemAlias.TableDocuments, "DocumentsTable");
                itemsRepo.RegisterNew(LayoutItemAlias.CreationDateFrom, "CreationDateFrom");
                itemsRepo.RegisterNew(LayoutItemAlias.CreationDateTo, "CreationDateTo");
                itemsRepo.RegisterNew(LayoutItemAlias.DocumentsLimit, "Limit");
                itemsRepo.RegisterNew(LayoutItemAlias.AuthorSelector, "Author");
                itemsRepo.RegisterNew(LayoutItemAlias.PreviewableFileHeader, "SelectedFile");

                // wrappers
                itemsRepo.RegisterNew(LayoutItemAlias.FilesGroupWrapper, "_filesGroupWrapper");
                itemsRepo.RegisterNew(LayoutItemAlias.FilePreviewGroupWrapper, "_filePreviewGroupWrapper");

                // table with documents control panel
                itemsRepo.RegisterNew(LayoutItemAlias.TableDocumentsExportButton, "ExportDocumentsToExcel");
                itemsRepo.RegisterNew(LayoutItemAlias.GroupPanelVisibilitySwitch, "GroupPanelVisibilitySwitch");
                itemsRepo.RegisterNew(LayoutItemAlias.ClearGroupingButton, "ClearGrouping");
                itemsRepo.RegisterNew(LayoutItemAlias.ColumnVisibilitySwitch, "ColumnVisibilitySwitch");
                itemsRepo.RegisterNew(LayoutItemAlias.CollapseExpandGroupsButton, "CollapseExpandGroups");
                itemsRepo.RegisterNew(LayoutItemAlias.ClearFiltersButton, "ClearFilters");

                BaseLayoutItemFinder itemFinder = new BaseLayoutItemFinder(customizableControl);

                return new LayoutItemHelper(itemsRepo, itemFinder);
            }

            public Control GetControl(LayoutItemAlias alias)
            {
                return ItemFinder.GetControl(Repository[alias]);
            }

            public TControl GetControl<TControl>(LayoutItemAlias alias) where TControl : Control
            {
                return ItemFinder.GetControl<TControl>(Repository[alias]);
            }

            public BaseLayoutItem GetItem(LayoutItemAlias alias)
            {
                return ItemFinder.GetItem(Repository[alias]);
            }

            public TItem GetItem<TItem>(LayoutItemAlias alias) where TItem : BaseLayoutItem
            {
                return ItemFinder.GetItem<TItem>(Repository[alias]);
            }

            public DevExpress.XtraLayout.BaseLayoutItem DocumentsTableControlPanelGroupXGroupWrapper
            {
                get { return ItemFinder.GetItem(Repository[LayoutItemAlias.TableDocumentsExportButton]); }
            }

            public DevExpress.XtraLayout.LayoutControlGroup DocumentsTableControlPanelGroup
            {
                get { return Utils.GetXLayoutGroup(DocumentsTableControlPanelGroupXGroupWrapper) as DevExpress.XtraLayout.LayoutControlGroup; }
            }

            public DevExpress.XtraEditors.SimpleButton ExportToExcelButton
            {
                get { return ItemFinder.GetControl<DevExpress.XtraEditors.SimpleButton>(Repository[LayoutItemAlias.TableDocumentsExportButton]); }
            }

            public ITableControl TableDocsDVControl
            {
                get { return FindPropertyItem<ITableControl>(LayoutItemAlias.TableDocuments); }
            }

            public ITableControl TableFilesDVControl
            {
                get { return FindPropertyItem<ITableControl>(LayoutItemAlias.TableFiles); }
            }

            public DevExpress.XtraGrid.GridControl TableDocsGridControl
            {
                get { return TableDocsDVControl.GetControl(); }
            }

            public DevExpress.XtraGrid.GridControl TableFilesGridControl
            {
                get { return TableFilesDVControl.GetControl(); }
            }

            public System.Windows.Forms.Control AuthorControl
            {
                get { return ItemFinder.GetControl(Repository[LayoutItemAlias.AuthorSelector]); }
            }

            public DevExpress.XtraLayout.BaseLayoutItem MeButtonItem
            {
                get { return FindPropertyItem<DevExpress.XtraLayout.BaseLayoutItem>(LayoutItemAlias.MeButton); }
            }

            public DevExpress.XtraLayout.BaseLayoutItem LimitItem
            {
                get { return ItemFinder.GetItem(Repository[LayoutItemAlias.DocumentsLimit]); }
            }

            public DevExpress.XtraLayout.BaseLayoutItem FilePreviewXGroupWrapper
            {
                get { return FindPropertyItem<DevExpress.XtraLayout.BaseLayoutItem>(LayoutItemAlias.FilePreviewGroupWrapper); }
            }

            public DevExpress.XtraLayout.BaseLayoutItem FilesXGroupWrapper
            {
                get { return FindPropertyItem<DevExpress.XtraLayout.BaseLayoutItem>(LayoutItemAlias.FilesGroupWrapper); }
            }

            public DevExpress.XtraLayout.LayoutControlGroup FilePreviewLayoutGroup
            {
                get { return Utils.GetXLayoutGroup(FilePreviewXGroupWrapper) as DevExpress.XtraLayout.LayoutControlGroup; }
            }

            public DevExpress.XtraLayout.LayoutControlGroup FilesLayoutGroup
            {
                get { return Utils.GetXLayoutGroup(FilesXGroupWrapper) as DevExpress.XtraLayout.LayoutControlGroup; }
            }

            public DevExpress.XtraLayout.BaseLayoutItem PreviewableFileHeaderXItem
            {
                get { return FindPropertyItem<DevExpress.XtraLayout.BaseLayoutItem>(LayoutItemAlias.PreviewableFileHeader); }
            }

            public DevExpress.XtraLayout.BaseLayoutItem TableFilesItem
            {
                get { return FindPropertyItem<DevExpress.XtraLayout.BaseLayoutItem>(LayoutItemAlias.TableFiles); }
            }

            public DevExpress.XtraEditors.SimpleButton ShowGroupPanelXButton
            {
                get { return ItemFinder.GetControl<DevExpress.XtraEditors.SimpleButton>(Repository[LayoutItemAlias.GroupPanelVisibilitySwitch]); }
            }

            public DevExpress.XtraEditors.SimpleButton ClearGroupingXButton
            {
                get { return ItemFinder.GetControl<DevExpress.XtraEditors.SimpleButton>(Repository[LayoutItemAlias.ClearGroupingButton]); }
            }

            public DevExpress.XtraEditors.SimpleButton ColumnVisibilitySwitchButton
            {
                get { return ItemFinder.GetControl<DevExpress.XtraEditors.SimpleButton>(Repository[LayoutItemAlias.ColumnVisibilitySwitch]); }
            }

            public System.Windows.Forms.Control SearchButton
            {
                get { return ItemFinder.GetControl(Repository[LayoutItemAlias.SearchButton]); }
            }

            public ILayoutPropertyItem PreviewableFileHeaderPropertyItem
            {
                get { return FindPropertyItem<ILayoutPropertyItem>(LayoutItemAlias.PreviewableFileHeader); }
            }

            public IPreviewFileControl FilePreviewDVControl
            {
                get { return FindPropertyItem<IPreviewFileControl>(LayoutItemAlias.FilePreview); }
            }

            public Control FilePreviewControl
            {
                get { return ItemFinder.GetControl(Repository[LayoutItemAlias.SearchButton]); }
            }

            public DevExpress.XtraEditors.DateEdit CreationDateFromDateEdit
            {
                get { return ItemFinder.GetControl<DevExpress.XtraEditors.DateEdit>(Repository[LayoutItemAlias.CreationDateFrom]); }
            }

            public DevExpress.XtraEditors.DateEdit CreationDateToDateEdit
            {
                get { return ItemFinder.GetControl<DevExpress.XtraEditors.DateEdit>(Repository[LayoutItemAlias.CreationDateTo]); }
            }

            public DevExpress.XtraLayout.BaseLayoutItem CreationDateFromItem
            {
                get { return ItemFinder.GetItem(Repository[LayoutItemAlias.CreationDateFrom]); }
            }

            public DevExpress.XtraLayout.BaseLayoutItem CreationDateToItem
            {
                get { return ItemFinder.GetItem(Repository[LayoutItemAlias.CreationDateTo]); }
            }

            public ILayoutPropertyItem DocumentsLimitPropertyItem
            {
                get { return FindPropertyItem<ILayoutPropertyItem>(LayoutItemAlias.DocumentsLimit); }
            }

            public ILayoutPropertyItem CreationDateFromPropertyItem
            {
                get { return FindPropertyItem<ILayoutPropertyItem>(LayoutItemAlias.CreationDateFrom); }
            }

            public ILayoutPropertyItem CreationDateToPropertyItem
            {
                get { return FindPropertyItem<ILayoutPropertyItem>(LayoutItemAlias.CreationDateTo); }
            }

            public ILayoutPropertyItem AuthorPropertyItem
            {
                get { return FindPropertyItem<ILayoutPropertyItem>(LayoutItemAlias.AuthorSelector); }
            }
        }

        private class CardUrlHelper
        {
            public const string urlSuffix = @"&ShowPanels=2048&";
            public const string cardIdPrefix = @"?CardID=";

            public readonly string baseUrl;
            public readonly string cardUrlTemplate;

            public CardUrlHelper(string baseUrl)
            {
                this.baseUrl = baseUrl;

                cardUrlTemplate = string.Concat(baseUrl, cardIdPrefix);
            }

            public string GetUrl(string cardId)
            {
                if (string.IsNullOrEmpty(cardId))
                {
                    return string.Empty;
                }

                Guid cardGuid;
                bool parseResult = Guid.TryParse(cardId, out cardGuid);
                if (!parseResult)
                {
                    return string.Empty;
                }

                return cardUrlTemplate + cardGuid.ToString("B") + urlSuffix;
            }
        }

        private class DateEditorButtonsHandler
        {
            private Dictionary<DevExpress.XtraEditors.DateEdit, List<EditorButtonEntry>> _buttonsDictionary;

            private class EditorButtonEntry
            {
                public EditorButton button;
                public EventHandler clickHandler;

                public EditorButtonEntry(EditorButton button, EventHandler clickHandler)
                {
                    this.button = button;
                    this.clickHandler = clickHandler;

                    button.Click += clickHandler;
                }
            }

            public void AddTo(DevExpress.XtraEditors.DateEdit dateEdit, EditorButton button, EventHandler clickHandler)
            {
                if (_buttonsDictionary == null)
                {
                    _buttonsDictionary = new Dictionary<DateEdit, List<EditorButtonEntry>>();
                }

                if (_buttonsDictionary.ContainsKey(dateEdit))
                {
                    if (_buttonsDictionary[dateEdit] == null)
                    {
                        _buttonsDictionary[dateEdit] = new List<EditorButtonEntry>();
                    }
                }
                else
                {
                    _buttonsDictionary.Add(dateEdit, new List<EditorButtonEntry>());
                }

                _buttonsDictionary[dateEdit].Add(new EditorButtonEntry(button, clickHandler));
                dateEdit.Properties.Buttons.Add(button);
            }

            public void Dispose()
            {
                if (_buttonsDictionary == null || _buttonsDictionary.Count == 0)
                {
                    return;
                }

                foreach (var buttonKvp in _buttonsDictionary)
                {
                    List<EditorButtonEntry> entries = buttonKvp.Value;

                    if (entries == null || entries.Count == 0) continue;

                    foreach (var entry in entries)
                    {
                        entry.button.Click -= entry.clickHandler;
                    }
                }

                _buttonsDictionary.Clear();
                _buttonsDictionary = null;
            }
        }

        internal class DbCommand
        {
            private System.Data.CommandType _cmdType;
            private string _cmdTypeEntry;
            private string _cmdText = null;
            private DbParamCollection _paramCollection;

            private DbCommand() { }

            public static DbCommand CreateAsProcedureExec(string schemeName, string procedureName)
            {
                if (string.IsNullOrWhiteSpace(schemeName) || string.IsNullOrEmpty(procedureName))
                {
                    throw new ArgumentNullException("Some of the arguments was empty/null");
                }

                DbCommand cmd = new DbCommand();
                cmd._cmdType = CommandType.StoredProcedure;
                cmd._cmdTypeEntry = string.Format("EXEC [{0}].[{1}]", schemeName, procedureName);

                cmd._paramCollection = new DbParamCollection();

                return cmd;
            }

            public static DbCommand CreateAsProcedureExec(string schemeName, string procedureName, DbParamCollection paramCollection)
            {
                if (string.IsNullOrWhiteSpace(schemeName) || string.IsNullOrEmpty(procedureName))
                {
                    throw new ArgumentNullException("Some of the arguments was empty/null");
                }

                DbCommand cmd = new DbCommand();
                cmd._cmdType = CommandType.StoredProcedure;
                cmd._cmdTypeEntry = string.Format("EXEC [{0}].[{1}]", schemeName, procedureName);

                if (paramCollection == null)
                {
                    cmd._paramCollection = new DbParamCollection();
                }
                else
                {
                    cmd._paramCollection = paramCollection;
                }

                return cmd;
            }

            public DbParamCollection Params
            {
                get { return _paramCollection; }
            }

            public string CommandText
            {
                get
                {
                    if (_cmdText != null)
                        return _cmdText;

                    StringBuilder sB = new StringBuilder();
                    sB.AppendLine(_cmdTypeEntry);

                    if (_paramCollection == null || _paramCollection.Params == null)
                    {
                        _cmdText = sB.ToString();
                        return _cmdText;
                    }

                    int n = _paramCollection.Params.Count;
                    if (n == 0)
                    {
                        _cmdText = sB.ToString();
                        return _cmdText;
                    }

                    sB.Append(_paramCollection.Params[0].ToString());

                    for (int i = 1; i < n; i++)
                    {
                        sB.Append("," + _paramCollection.Params[i].ToString());
                    }

                    _cmdText = sB.ToString();
                    return _cmdText;
                }
            }

            public override string ToString()
            {
                return CommandText;
            }
        }

        internal class DbParamCollection
        {
            private List<DbParameter> _params;

            public DbParamCollection()
            {
                _params = new List<DbParameter>();
            }

            public List<DbParameter> Params
            {
                get { return _params; }
            }

            public void Add<T>(string name, T value, SqlDbType dbType)
            {
                if (Utils.CanBeNull<T>() && value == null)
                {
                    _params.Add(new DbParameter(name, "NULL", dbType));
                }
                else
                {
                    _params.Add(new DbParameter(name, value.ToString(), dbType));
                }
            }

            //public void Add(string name, string value, SqlDbType dbType)
            //{
            //    _params.Add(new DbParameter(name, value, dbType));
            //}

            public void Add(DbParameter parameter)
            {
                _params.Add(parameter);
            }
        }

        internal class DbParameter
        {
            public string Name { get; private set; }
            public string Value { get; private set; }
            public System.Data.SqlDbType Type { get; private set; }

            public DbParameter(string name, string value, SqlDbType dbType)
            {
                Name = name;
                Value = value;
                Type = dbType;
            }

            public override string ToString()
            {
                if (Value.ToUpper() == "NULL")
                    return string.Format("@{0}={1}", Name, Value);

                switch (this.Type)
                {
                    case SqlDbType.Char:
                    case SqlDbType.NVarChar:
                    case SqlDbType.VarChar:
                    case SqlDbType.NChar:
                    case SqlDbType.NText:
                    case SqlDbType.Text:
                    case SqlDbType.Date:
                    case SqlDbType.DateTime:
                    case SqlDbType.SmallDateTime:
                    case SqlDbType.DateTime2:
                    case SqlDbType.UniqueIdentifier:
                    case SqlDbType.Xml:
                        return string.Format("@{0}='{1}'", Name, Value);

                    default:
                        return string.Format("@{0}={1}", Name, Value);
                }
            }
        }

        private static class GridDataManager
        {
            private static Dictionary<GridControl, GridViewData> _data;

            public static void Register(GridControl control, GridViewData viewData)
            {
                if (_data == null)
                {
                    _data = new Dictionary<GridControl, GridViewData>();
                }

                if (_data.ContainsKey(control))
                {
                    _data[control] = viewData;
                }
                else
                {
                    _data.Add(control, viewData);
                }
            }

            public static GridViewData GetData(GridControl control)
            {
                if (_data == null || !_data.ContainsKey(control))
                {
                    throw new NullReferenceException("Data is null or doesn't contains value with key '" + control + "'");
                }

                return _data[control];
            }

            public static void ProcessInitialize(GridControl control, System.Action firstInitAction)
            {
                if (_data == null || !_data.ContainsKey(control))
                {
                    throw new NullReferenceException("Data is null or doesn't contains value with key '" + control + "'");
                }

                if (_data[control].persistenceOptions == GridViewData.PersistenceOptions.SaveToXml)
                {
                    if (System.IO.File.Exists(_data[control].DataPathXml))
                    {
                        _data[control].view.RestoreLayoutFromXml(_data[control].DataPathXml);
                    }
                    else
                    {
                        firstInitAction();
                    }
                    return;
                }

                firstInitAction();
            }

            public static void ProcessDataSourceChange(GridControl control)
            {
                if (_data.ContainsKey(control))
                {
                    if (_data[control].OnDataSourceChanged != null)
                    {
                        _data[control].OnDataSourceChanged();
                    }
                }
            }

            private static DocsVision.BackOffice.WinForms.Controls.GridExView Save(GridControl control)
            {
                if (_data == null || !_data.ContainsKey(control))
                {
                    throw new NullReferenceException("Data is null or doesn't contains value with key '" + control + "'");
                }

                if (_data[control].persistenceOptions == GridViewData.PersistenceOptions.SaveToXml)
                {
                    if (!System.IO.File.Exists(_data[control].DataPathXml))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(_data[control].DataPathXml));
                    }

                    _data[control].view.SaveLayoutToXml(_data[control].DataPathXml);
                }

                return _data[control].view;
            }

            private static void ProcessInitialize()
            {
                if (_data == null || _data.Count == 0)
                {
                    return;
                }

                foreach (var data in _data)
                {
                    if (data.Value.persistenceOptions == GridViewData.PersistenceOptions.SaveToXml)
                    {
                        data.Value.view.RestoreLayoutFromXml(data.Value.DataPathXml);
                    }
                }
            }

            public static void ProcessDeinitialize()
            {
                if (_data == null || _data.Count == 0)
                {
                    return;
                }

                foreach (var data in _data)
                {
                    if (data.Value.persistenceOptions == GridViewData.PersistenceOptions.SaveToXml)
                    {
                        if (!System.IO.File.Exists(data.Value.DataPathXml))
                        {
                            Directory.CreateDirectory(Path.GetDirectoryName(data.Value.DataPathXml));
                        }

                        data.Value.view.SaveLayoutToXml(data.Value.DataPathXml);
                    }
                }
            }
        }

        private static class ContextGridViewMenuItemHelper
        {
            public static class CustomMenuItems
            {
                public const string OpenCardItemCaption = "Открыть карточку";
                public const string CopyUrlItemCaption = "Скопировать ссылку на карточку";

                public const string ClearSortingCaption = "Очистить сортировку"; // for rename built-in option "Clear All Sorting"
            }

            // TO DO: receive default captions or items from DevExpress library if possible
            public static class BuiltInMenuItems
            {
                // Меню по колонке
                public const string ClearSortingCaption = "Clear All Sorting";
                public const string FilterEditorCaption = "Редактор фильтра";
                public const string ShowSearchPanelCaption = "Показать панель поиска";
                public const string ShowAutoFilterRowCaption = "Показать строку автофильтрации";
                public const string ShowGroupPanelCaption = "Показать область группировки";
                public const string HideGroupPanelCaption = "Скрыть область группировки";
                public const string SetColumnsVisibilityCaption = "Настройка колонки";
                public const string HideColumnCaption = "Удалить столбец";

                //Меню по панели группировки
                public const string CollapseAllGroupsCaption = "Свернуть всё";
                public const string ExpandAllGroupsCaption = "Развернуть всё";
                public const string ClearGroupingCaption = "Отменить группировку";
            }

            public static DXMenuItem GetDXMenuItem(GridViewMenu menu, string caption)
            {
                foreach (DXMenuItem item in menu.Items)
                {
                    if (item.Caption == caption)
                    {
                        return item;
                    }
                }

                return null;
            }

            public static void RenameMenuItem(GridViewMenu menu, string currentCaption, string newCaption)
            {
                foreach (DXMenuItem item in menu.Items)
                {
                    if (item.Caption == currentCaption)
                    {
                        item.Caption = newCaption;
                        return;
                    }
                }
            }

            public static void HideDocsGridViewColumnMenuStandardItems(GridViewMenu menu)
            {
                foreach (DXMenuItem item in menu.Items)
                {
                    switch (item.Caption)
                    {
                        case BuiltInMenuItems.FilterEditorCaption:
                        case BuiltInMenuItems.ShowSearchPanelCaption:
                        case BuiltInMenuItems.ShowGroupPanelCaption:
                        case BuiltInMenuItems.HideGroupPanelCaption:
                        case BuiltInMenuItems.SetColumnsVisibilityCaption:
                        case BuiltInMenuItems.HideColumnCaption:
                        case BuiltInMenuItems.ShowAutoFilterRowCaption:
                            item.Visible = false;
                            break;
                    }
                }
            }

            public static void HideDocsGridViewGroupPanelMenuStandardItems(GridViewMenu menu)
            {
                foreach (DXMenuItem item in menu.Items)
                {
                    switch (item.Caption)
                    {
                        case BuiltInMenuItems.CollapseAllGroupsCaption:
                        case BuiltInMenuItems.ExpandAllGroupsCaption:
                        //case BuiltInMenuItems.ShowGroupPanelCaption:
                        case BuiltInMenuItems.HideGroupPanelCaption:
                        case BuiltInMenuItems.ClearGroupingCaption:
                            item.Visible = false;
                            break;
                    }
                }
            }
        }

        private static class ContextMenuHelper
        {
            public static DXPopupMenu CreateColumnSelectorPopupMenu(DocsVision.BackOffice.WinForms.Controls.GridExView view)
            {
                DXPopupMenu popupMenu = null;

                if (view != null)
                {
                    popupMenu = new DXPopupMenu();

                    var columns = view.Columns;
                    foreach (DevExpress.XtraGrid.Columns.GridColumn column in columns)
                    {
                        var item = AddColumnCheckItem(popupMenu, column);
                        item.Appearance.Font = new Font("Segoe UI", 9F);
                        item.Appearance.TextOptions.HAlignment = HorzAlignment.Near;
                    }
                }
                else
                {

                }

                return popupMenu;
            }

            public static DXMenuCheckItem AddColumnCheckItem(DXPopupMenu menu, DevExpress.XtraGrid.Columns.GridColumn column)
            {
                if (menu == null || column == null)
                {
                    return null;
                }

                DXMenuCheckItem checkItem = new DXMenuCheckItem(
                    column.Caption, column.Visible);
                checkItem.Click += (s, e) =>
                {
                    if (s is DXMenuCheckItem)
                    {
                        DXMenuCheckItem item = s as DXMenuCheckItem;
                        if (item != null)
                        {
                            column.Visible = item.Checked;
                        }
                    }
                };

                if (!string.IsNullOrEmpty(column.ToolTip))
                {
                    //checkItem.Hint = column.ToolTip;
                }

                menu.Items.Add(checkItem);

                return checkItem;
            }

            //private static void CheckItem_CheckedChanged(object sender, EventArgs e)
            //{
            //    if (sender is DXMenuCheckItem)
            //    {
            //        DXMenuCheckItem item = sender as DXMenuCheckItem;
            //        if (item != null)
            //        {
            //            bool newState = !item.Checked;
            //            item.Checked = newState;
            //            column.Visible = newState;
            //        }
            //    }
            //}
        }

        private class GridViewMenuItemRepository
        {
            private Dictionary<DXMenuItem, EventHandler> _menuItems = new Dictionary<DXMenuItem, EventHandler>();

            public DXMenuItem AddOrGetItem(string caption, EventHandler handler)
            {
                if (handler == null || string.IsNullOrEmpty(caption))
                {
                    return null;
                }

                DXMenuItem existingItem = GetItem(caption);
                if (existingItem == null)
                {
                    //if (_menuItems[existingItem] != null) existingItem.Click -= _menuItems[existingItem];
                    existingItem = new DXMenuItem(caption, handler);
                    _menuItems.Add(existingItem, handler);
                }
                else
                {
                    if (_menuItems[existingItem] != null) existingItem.Click -= _menuItems[existingItem];
                    _menuItems[existingItem] = handler;
                    existingItem.Click += handler;
                }

                return existingItem;
            }

            public void RemoveHandlers()
            {
                try
                {
                    foreach (var kvp in _menuItems)
                    {
                        if (kvp.Value != null)
                        {
                            kvp.Key.Click -= kvp.Value;
                        }

                        _menuItems[kvp.Key] = null;
                    }
                }
                catch { }
            }

            public void Dispose()
            {
                RemoveHandlers();
                _menuItems.Clear();
            }

            public DXMenuItem GetItem(string caption)
            {
                if (string.IsNullOrEmpty(caption))
                {
                    return null;
                }

                foreach (DXMenuItem item in _menuItems.Keys)
                {
                    if (item.Caption == caption)
                    {
                        return item;
                    }
                }

                return null;
            }
        }

        private class GridViewData
        {
            [Flags]
            public enum PersistenceOptions
            {
                None = 0,
                Session = 1,
                SaveToXml = 2,
                SaveToRegistry = 4,
                SaveToStream = 8,
            }

            public PersistenceOptions persistenceOptions = PersistenceOptions.None;

            public DocsVision.BackOffice.WinForms.Controls.GridExView view;

            public string DataPathXml { get; private set; }

            public System.Action OnDataSourceChanged;

            public GridViewData(DocsVision.BackOffice.WinForms.Controls.GridExView view)
            {
                this.view = view;
            }

            public GridViewData(DocsVision.BackOffice.WinForms.Controls.GridExView view, PersistenceOptions persistenceOptions)
            {
                this.view = view;
                this.persistenceOptions = persistenceOptions;
            }

            public GridViewData(string dataPathXml, DocsVision.BackOffice.WinForms.Controls.GridExView view, PersistenceOptions persistenceOptions)
            {
                DataPathXml = dataPathXml;
                this.view = view;
                this.persistenceOptions = persistenceOptions;
            }
        }

        private static class DbProcedureRegistry
        {
            public enum Procedure
            {
                GetDocumentsInfo,
                GetReceiversInternalInfo,
                GetReceiversCounterpartyInfo,
                GetFilesInfo,
            }

            private static Dictionary<Procedure, Tuple<string, string>> _procsRegistry;

            public static Tuple<string, string> GetProcedureInfo(Procedure procedure)
            {
                if (_procsRegistry == null || !_procsRegistry.ContainsKey(procedure))
                {
                    throw new NullReferenceException("Registry is null or doesn't contains data with key '" + procedure + "'");
                }

                return _procsRegistry[procedure];
            }

            public static void Add(Procedure procedure, string schemeName, string procName)
            {
                if (_procsRegistry == null)
                {
                    _procsRegistry = new Dictionary<Procedure, Tuple<string, string>>();
                }

                if (!_procsRegistry.ContainsKey(procedure))
                {
                    _procsRegistry.Add(procedure, new Tuple<string, string>(schemeName, procName));
                }
                else
                {
                    _procsRegistry[procedure] = new Tuple<string, string>(schemeName, procName);
                }
            }

            static DbProcedureRegistry()
            {
                Add(Procedure.GetDocumentsInfo, "dbo", "GAZ_GetDocumentsCreatedByUser");
                Add(Procedure.GetReceiversInternalInfo, "dbo", "GAZ_GetInternalReceiversFromDocument");
                Add(Procedure.GetReceiversCounterpartyInfo, "dbo", "GAZ_GetCounterpartyReceiversFromDocument");
                Add(Procedure.GetFilesInfo, "dbo", "GAZ_GetDocumentFilesInfo");
            }
        }

        private class ColumnDefinition
        {
            public string name;
            public Type type;
            public string displayName;

            public ColumnDefinition(string name, Type type, string displayName)
            {
                this.name = name;
                this.type = type;
                this.displayName = displayName;
            }
        }

        private static class ColumnDefinitionsForTableDocs
        {
            public enum Column
            {
                KindName, Type, RegNumber, ProjectNumber, Name, CreationDate, RegDate, SignerFIO, ReceiversInternal, ReceiversCounterparty, SenderCounterpartyOrg,
                LegalEntity, StateName, InitiatorFIO, CardId
            }

            private static Dictionary<Column, ColumnDefinition> _defs;
            private static Dictionary<Column, ColumnDefinition> _defaultDefs;

            private static void Bind(Column column, ColumnDefinition definition)
            {
                if (_defs == null)
                {
                    _defs = new Dictionary<Column, ColumnDefinition>();
                }

                if (_defs.ContainsKey(column))
                {
                    _defs[column] = definition;
                }
                else
                {
                    _defs.Add(column, definition);
                }
            }

            public static Dictionary<Column, ColumnDefinition> DefaultDefs
            {
                get
                {
                    if (_defaultDefs == null)
                    {
                        SetDefaultBindings();
                    }

                    return _defaultDefs;
                }
            }

            public static void SetDefaultBindings()
            {
                if (_defaultDefs != null) return;

                _defaultDefs = new Dictionary<Column, ColumnDefinition>();

                _defaultDefs.Add(Column.KindName, new ColumnDefinition("KindName", typeof(string), "Вид"));
                _defaultDefs.Add(Column.Type, new ColumnDefinition("Type", typeof(string), "Тип"));
                _defaultDefs.Add(Column.RegNumber, new ColumnDefinition("RegNumber", typeof(string), "Рег. номер"));
                _defaultDefs.Add(Column.ProjectNumber, new ColumnDefinition("ProjectNumber", typeof(string), "Проектный номер"));
                _defaultDefs.Add(Column.Name, new ColumnDefinition("Name", typeof(string), "Заголовок"));
                _defaultDefs.Add(Column.CreationDate, new ColumnDefinition("CreationDate", typeof(System.DateTime), "Дата создания"));
                _defaultDefs.Add(Column.RegDate, new ColumnDefinition("RegDate", typeof(System.DateTime), "Дата регистрации"));
                _defaultDefs.Add(Column.SignerFIO, new ColumnDefinition("SignerFIO", typeof(string), "Подписант"));
                _defaultDefs.Add(Column.ReceiversInternal, new ColumnDefinition("ReceiversInternal", typeof(string), "Получатели внутренние"));
                _defaultDefs.Add(Column.ReceiversCounterparty, new ColumnDefinition("ReceiversCounterparty", typeof(string), "Получатели контрагенты"));
                _defaultDefs.Add(Column.SenderCounterpartyOrg, new ColumnDefinition("SenderCounterpartyOrg", typeof(string), "Отправитель контрагент"));
                _defaultDefs.Add(Column.LegalEntity, new ColumnDefinition("LegalEntity", typeof(string), "Юридическое лицо"));
                _defaultDefs.Add(Column.StateName, new ColumnDefinition("StateName", typeof(string), "Состояние"));
                _defaultDefs.Add(Column.InitiatorFIO, new ColumnDefinition("InitiatorFIO", typeof(string), "Инициатор"));
                _defaultDefs.Add(Column.CardId, new ColumnDefinition("CardId", typeof(System.Guid), "ID карточки"));
            }
        }

        private static class ColumnDefinitionsForTableFiles
        {
            public enum Column
            {
                FileId, FileType, FileName, CurrentVersionId, CheckinDate, Version, VersionNumber, AuthorFIO
            }

            private static Dictionary<Column, ColumnDefinition> _defs;
            private static Dictionary<Column, ColumnDefinition> _defaultDefs;

            private static void Bind(Column column, ColumnDefinition definition)
            {
                if (_defs == null)
                {
                    _defs = new Dictionary<Column, ColumnDefinition>();
                }

                if (_defs.ContainsKey(column))
                {
                    _defs[column] = definition;
                }
                else
                {
                    _defs.Add(column, definition);
                }
            }

            public static Dictionary<Column, ColumnDefinition> DefaultDefs
            {
                get
                {
                    if (_defaultDefs == null)
                    {
                        SetDefaultBindings();
                    }

                    return _defaultDefs;
                }
            }

            public static void SetDefaultBindings()
            {
                if (_defaultDefs != null) return;

                _defaultDefs = new Dictionary<Column, ColumnDefinition>();

                _defaultDefs.Add(Column.FileId, new ColumnDefinition("FileId", typeof(Guid), "ID файла"));
                _defaultDefs.Add(Column.FileType, new ColumnDefinition("FileType", typeof(string), "Тип"));
                _defaultDefs.Add(Column.FileName, new ColumnDefinition("FileName", typeof(string), "Название"));
                _defaultDefs.Add(Column.CurrentVersionId, new ColumnDefinition("CurrentVersionId", typeof(Guid), "ID текущей версии файла"));
                _defaultDefs.Add(Column.CheckinDate, new ColumnDefinition("CheckinDate", typeof(System.DateTime), "Дата изменения"));
                _defaultDefs.Add(Column.Version, new ColumnDefinition("Version", typeof(int), "Версия"));
                _defaultDefs.Add(Column.VersionNumber, new ColumnDefinition("VersionNumber", typeof(int), "Номер версии"));
                _defaultDefs.Add(Column.AuthorFIO, new ColumnDefinition("AuthorFIO", typeof(string), "Автор"));
            }
        }

        public class UserProfileCardData
        {
            public int documentsLimit;
        }

        private interface IPersistantDataController
        {
            bool TrySave<TData>(TData data);
            bool TryLoad<TData>(out TData data);
        }

        private class UserProfileCardDataController : IPersistantDataController
        {
            public readonly ISerializationStrategy serializationStrategy;
            public readonly DocsVision.Platform.ObjectModel.ObjectContext objectContext;

            public UserProfileCardDataController(ISerializationStrategy serializationStrategy, DocsVision.Platform.ObjectModel.ObjectContext objectContext)
            {
                this.serializationStrategy = serializationStrategy;
                this.objectContext = objectContext;
            }

            public bool TryLoad<TData>(out TData data)
            {
                IUserProfileCardService svc = objectContext.GetService<IUserProfileCardService>();

                object setting = svc.GetSetting(SettingId, null, SettingObjectId);
                data = default(TData);
                bool result = false;

                if (setting != null)
                {
                    try
                    {
                        data = serializationStrategy.Deserialize<TData>(setting as string);
                        result = true;
                    }
                    catch { }
                }

                return result;
            }

            public bool TrySave<TData>(TData data)
            {
                bool result = false;

                IUserProfileCardService svc = objectContext.GetService<IUserProfileCardService>();

                string serializedData = serializationStrategy.Serialize(data);

                try
                {
                    svc.SetSetting(SettingId, SettingObjectId, serializedData);
                    result = true;
                }
                catch { }

                return result;
            }
        }

        private interface ISerializationStrategy
        {
            string Serialize<T>(T obj);
            T Deserialize<T>(string data);
        }

        private class XmlSerializer : ISerializationStrategy
        {
            public string Serialize<T>(T obj)
            {
                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(T));
                var namespaces = new XmlSerializerNamespaces();
                namespaces.Add("", ""); // remove xmlns

                using (var writer = new StringWriter(new StringBuilder()))
                {
                    serializer.Serialize(writer, obj, namespaces);
                    return writer.ToString();
                }
            }

            public T Deserialize<T>(string xml)
            {
                if (string.IsNullOrEmpty(xml))
                {
                    return default(T);
                }

                var serializer = new System.Xml.Serialization.XmlSerializer(typeof(T));

                using (var reader = new StringReader(xml))
                {
                    return (T)serializer.Deserialize(reader);
                }
            }
        }

        private static class Utils
        {
            public static DevExpress.XtraLayout.LayoutGroup GetXLayoutGroup(DevExpress.XtraLayout.BaseLayoutItem childItem)
            {
                if (childItem == null)
                {
                    throw new NullReferenceException("Child layout item reference is not set");
                }

                if (childItem.Parent == null)
                {
                    throw new NullReferenceException("Parent of the child layout item is null");
                }

                return childItem.Parent;
            }

            public static bool IsAllGroupRowsExpanded(DevExpress.XtraGrid.Views.Grid.GridView gridView)
            {
                int handle = -1; // Групповые строки начинаются с -1

                while (gridView.GetRow(handle) != null)
                {
                    if (!gridView.GetRowExpanded(handle))
                    {
                        return false;
                    }
                    handle--;
                }

                return true;
            }

            public static bool IsAllGroupRowsCollapsed(DevExpress.XtraGrid.Views.Grid.GridView gridView)
            {
                int handle = -1; // Групповые строки начинаются с -1

                while (gridView.GetRow(handle) != null)
                {
                    if (gridView.GetRowExpanded(handle))
                    {
                        return false;
                    }
                    handle--;
                }

                return true;
            }

            public static object GetPrivateFieldValue(object obj, string fieldName)
            {
                if (obj == null)
                    throw new ArgumentNullException("obj");

                Type type = obj.GetType();

                // Ищем приватное нестатическое поле
                FieldInfo field = type.GetField(fieldName, BindingFlags.NonPublic | BindingFlags.Instance);

                if (field == null)
                    throw new InvalidOperationException(string.Format("Поле '{0}' не найдено в типе {1}", fieldName, type.Name));

                return field.GetValue(obj);
            }

            public static void CloseExcelAndReleaseComObjects(Excel.Workbook workbook, Excel.Application excelApp, List<object> objects)
            {
                if (workbook != null) workbook.Close(false);
                if (excelApp != null) excelApp.Quit();

                int n = objects.Count;
                for (int i = 0; i < n; i++)
                {
                    if (objects[i] != null) Marshal.FinalReleaseComObject(objects[i]);
                }

                for (int i = 0; i < n; i++)
                {
                    objects[i] = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            public static Image StringToImage(string base64String, bool compressed)
            {
                if (compressed)
                {
                    base64String = BaseResolver.DecompressString(base64String);
                }

                byte[] bytes = Convert.FromBase64String(base64String);
                using (MemoryStream ms = new MemoryStream(bytes))
                {
                    return Image.FromStream(ms);
                }
            }

            public static Image StringToImage(string base64String)
            {
                byte[] bytes = Convert.FromBase64String(base64String);
                using (MemoryStream ms = new MemoryStream(bytes))
                {
                    return Image.FromStream(ms);
                }
            }

            public static DevExpress.Utils.Svg.SvgImage StringToSvgImage(string base64String)
            {
                byte[] svgBytes = Convert.FromBase64String(base64String);
                using (MemoryStream stream = new MemoryStream(svgBytes))
                {
                    return DevExpress.Utils.Svg.SvgImage.FromStream(stream);
                }
            }

            public static DevExpress.Utils.Svg.SvgImage StringToSvgImage(string base64String, bool compressed)
            {
                if (compressed)
                {
                    base64String = BaseResolver.DecompressString(base64String);
                }

                byte[] svgBytes = Convert.FromBase64String(base64String);
                using (MemoryStream stream = new MemoryStream(svgBytes))
                {
                    return DevExpress.Utils.Svg.SvgImage.FromStream(stream);
                }
            }

            public static string ImageToBase64String(string imagePath)
            {
                string ext = Path.GetExtension(imagePath).ToLower();
                if (ext == ".svg")
                {
                    return Convert.ToBase64String(Encoding.UTF8.GetBytes(System.IO.File.ReadAllText(imagePath, Encoding.UTF8)));
                }

                return Convert.ToBase64String(System.IO.File.ReadAllBytes(imagePath));
            }

            public static string ImageToCompressedBase64String(string imagePath)
            {
                string ext = Path.GetExtension(imagePath).ToLower();
                if (ext == ".svg")
                {
                    return BaseResolver.CompressString(Convert.ToBase64String(Encoding.UTF8.GetBytes(System.IO.File.ReadAllText(imagePath, Encoding.UTF8))));
                }

                return BaseResolver.CompressString(Convert.ToBase64String(System.IO.File.ReadAllBytes(imagePath)));
            }

            public static bool CanBeNull<T>()
            {
                Type type = typeof(T);

                if (!type.IsValueType)
                    return true;

                if (IsNullableType(type))
                    return true;

                return false;
            }

            /// <summary>
            /// Проверяет, является ли тип Nullable<T>
            /// </summary>
            /// <param name="type"></param>
            /// <returns></returns>
            private static bool IsNullableType(Type type)
            {
                return type.IsGenericType
                    && type.GetGenericTypeDefinition() == typeof(Nullable<>);
            }

            public static DocumentFile GetFile(ObjectCollection<DocumentFile> files, Guid fileId)
            {
                if (files == null)
                {
                    return null;
                }

                if (fileId == Guid.Empty)
                {
                    return null;
                }

                int n = files.Count;
                for (int i = 0; i < n; i++)
                {
                    if (files[i].FileId == fileId)
                    {
                        return files[i];
                    }
                }

                return null;
            }
        }

        #endregion

        #region Properties

        public string SEName
        {
            get
            {
                try
                {
                    string SEname = (string)Utils.GetPrivateFieldValue(GAZSE, "SEName");
                    if (string.IsNullOrEmpty(SEname))
                    {
                        return _SEDefaultName;
                    }
                    return SEname;
                }
                catch
                {
                    return _SEDefaultName;
                }
            }
        }

        private ICustomizableControl CustomizableControl { get { return CardControl; } }

        private BarManager MainBarManager { get { return CustomizableControl.BarManager; } }

        private Document SelectedDocument
        {
            get
            {
                Document document = null;
                var gridWrapper = _xtraGridRepository[Table.UserDocuments] as UserDocsGridWrapper;
                try
                {
                    document = CardControl.ObjectContext.GetObject<Document>(gridWrapper.SelectedDocumentId);
                }
                catch (Exception ex)
                {
                    _logger.Error("get.SelectedDocument" + Environment.NewLine + ex.ToString());
                }

                return document;
            }
        }

        private string ServerUrl
        {
            get
            {
                return Session.Properties[SessionPropertyNames.ServerUrlSessionPropertyName].Value + "";
            }
        }

        private new Guid CurrentUserId
        {
            get { return CardControl.ObjectContext.GetService<IStaffService>().GetCurrentEmployee().GetObjectId(); }
        }

        #endregion

        public ReportOnDocumentsScript() : base()
        {
            _logger = new SimpleLogger(typeof(ReportOnDocumentsScript).Name, _loggerFilePath, _enableLogging);
            _logger.Info("Script instance created");
        }

        #region Methods

        #region Logging

        private static double ToMegabytes(long bytes)
        {
            return Math.Round(bytes / 1024.0 / 1024.0, 2);
        }

        private void LogMemoryUsage(string context)
        {
            if (!_logger.EnableLogging) return;

            try
            {
                long workingSet = _currentProcess.WorkingSet64;
                long privateMemory = _currentProcess.PrivateMemorySize64;
                long virtualMemory = _currentProcess.VirtualMemorySize64;
                long gcMemory = GC.GetTotalMemory(false);

                string logEntry = string.Format
                (
                    "Context: {0}\n" +
                    "  Working Set: {1} MB\n" +
                    "  Private Memory: {2} MB\n" +
                    "  Virtual Memory: {3} MB\n" +
                    "  GC Total Memory: {4} MB\n",
                    context,
                    ToMegabytes(workingSet),
                    ToMegabytes(privateMemory),
                    ToMegabytes(virtualMemory),
                    ToMegabytes(gcMemory)
                );

                _logger.Trace(logEntry);
            }
            catch (Exception ex)
            {
                _logger.Error("Error logging memory: " + ex.Message);
            }
        }

        private static void LogMemoryUsage(string context, SimpleLogger logger)
        {
            if (logger == null || !logger.EnableLogging) return;

            try
            {
                long workingSet = _currentProcess.WorkingSet64;
                long privateMemory = _currentProcess.PrivateMemorySize64;
                long virtualMemory = _currentProcess.VirtualMemorySize64;
                long gcMemory = GC.GetTotalMemory(false);

                string logEntry = string.Format
                (
                    "Context: {0}\n" +
                    "  Working Set: {1} MB\n" +
                    "  Private Memory: {2} MB\n" +
                    "  Virtual Memory: {3} MB\n" +
                    "  GC Total Memory: {4} MB\n",
                    context,
                    ToMegabytes(workingSet),
                    ToMegabytes(privateMemory),
                    ToMegabytes(virtualMemory),
                    ToMegabytes(gcMemory)
                );

                logger.Trace(logEntry);
            }
            catch (Exception ex)
            {
                logger.Error("Error logging memory: " + ex.Message);
            }
        }

        private void LogSessionProperties()
        {
            StringBuilder sB = new StringBuilder();
            sB.AppendLine("Параметры пользовательской сессии:");
            foreach (var x in Session.Properties)
            {
                sB.AppendLine(x.Name + "=" + x.Value);
            }
            _logger.Debug(sB.ToString());
        }

        private void LogUI()
        {
            _logger.Trace("LogUI begin");

            try
            {
                var items = CustomizableControl.LayoutControl.Items;
                if (items == null)
                {
                    _logger.Warn("Layout items collection is null");
                    return;
                }

                int foreachCount = 0;
                foreach (var item in items)
                {
                    foreachCount++;
                }
                _logger.Trace("foreach count:" + foreachCount);

                if (items.Count != items.ItemCount)
                {
                    _logger.Trace(string.Format("Items collection's counts are not equal. items.Count={0} items.ItemCount={1}", items.Count, items.ItemCount));
                }

                _logger.Trace("Items.GroupCount=" + items.GroupCount);

                int n = items.Count;
                for (int i = 0; i < n; i++)
                {
                    _logger.Trace(string.Format("item[{0}] type: {1}", i, items[i].GetType().FullName));

                    if (items[i] is BaseLayoutItem)
                    {
                        BaseLayoutItem baseLayoutItem = (BaseLayoutItem)items[i];
                        _logger.Trace(string.Format("item[{0}] name:{1}; parentName:{2}; typeName:{3}, text:{4}, visible:{5}",
                            i, baseLayoutItem.Name, baseLayoutItem.ParentName, baseLayoutItem.TypeName, baseLayoutItem.Text, baseLayoutItem.Visible));
                    }
                    else
                    {
                        _logger.Warn(string.Format("item[{0}] is not assignable to type {1}", i, typeof(BaseLayoutItem)));
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("LogUI(): " + ex.ToString());
            }

            _logger.Trace("LogUI end");
        }

        private void LogUI(LayoutControlGroup parent)
        {
            if (parent == null)
            {
                throw new ArgumentNullException("Parent is null");
            }

            _logger.Trace("LogUI begin. Parent.Name: " + parent.Name);

            try
            {
                var items = parent.Items;
                if (items == null)
                {
                    _logger.Warn("Layout items collection is null");
                    return;
                }

                int foreachCount = 0;
                foreach (var item in items)
                {
                    foreachCount++;
                }
                _logger.Trace("foreach count:" + foreachCount);

                if (items.Count != items.ItemCount)
                {
                    _logger.Trace(string.Format("Items collection's counts are not equal. items.Count={0} items.ItemCount={1}", items.Count, items.ItemCount));
                }

                _logger.Trace("Items.GroupCount=" + items.GroupCount);

                int n = items.Count;
                for (int i = 0; i < n; i++)
                {
                    _logger.Trace(string.Format("item[{0}] type: {1}", i, items[i].GetType().FullName));

                    if (items[i] is BaseLayoutItem)
                    {
                        BaseLayoutItem baseLayoutItem = (BaseLayoutItem)items[i];
                        _logger.Trace(string.Format("item[{0}] name:{1}; parentName:{2}; typeName:{3}, text:{4}, visible:{5}",
                            i, baseLayoutItem.Name, baseLayoutItem.ParentName, baseLayoutItem.TypeName, baseLayoutItem.Text, baseLayoutItem.Visible));
                    }
                    else
                    {
                        _logger.Warn(string.Format("item[{0}] is not assignable to type {1}", i, typeof(BaseLayoutItem)));
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("LogUI(): " + ex.ToString());
            }

            _logger.Trace("LogUI end");
        }

        private void TraceDataTable(System.Data.DataTable dataTable, int rowLimit = 10)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("TraceDataTable begin");

            if (dataTable == null)
            {
                sb.AppendLine("DataTable reference is not set");
            }
            else if (dataTable.Rows.Count == 0)
            {
                sb.AppendLine("DataTable rowCount: " + dataTable.Rows.Count);
            }
            else
            {
                int rowCount = dataTable.Rows.Count;
                sb.AppendLine("DataTable rowCount: " + rowCount);

                sb.Append("Columns: ");
                for (int colIndex = 0; colIndex < dataTable.Columns.Count; colIndex++)
                {
                    sb.Append(dataTable.Columns[colIndex].ColumnName);
                    if (colIndex < dataTable.Columns.Count - 1)
                        sb.Append(", ");
                }
                sb.AppendLine();

                int displayRowCount = Math.Min(rowCount, rowLimit);
                for (int i = 0; i < displayRowCount; i++)
                {
                    sb.Append(string.Format("Row {0}: ", i));
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        var cellValue = dataTable.Rows[i][j];
                        string cellValueStr = "NULL";
                        if (cellValue != null) cellValueStr = cellValue.ToString();
                        sb.Append(string.Format("{0}={1}", dataTable.Columns[j].ColumnName, cellValueStr));
                        if (j < dataTable.Columns.Count - 1)
                            sb.Append(", ");
                    }
                    sb.AppendLine();
                }
            }

            sb.AppendLine("TraceDataTable end");
            _logger.Trace(sb.ToString());
        }

        #endregion

        private void OpenCard(Guid cardId)
        {
            _logger.Trace("OpenCard by Guid");

            string url = _cardUrlHelper.GetUrl(cardId.ToString());
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url)
            {
                UseShellExecute = true
            });
        }

        internal static void OpenCard(string url)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url)
            {
                UseShellExecute = true
            });
        }

        private DbSelectDocsOptions GetDefaultSelectOptions()
        {
            DbSelectDocsOptions selectOptions = new DbSelectDocsOptions();
            selectOptions.userId = _userId;
            selectOptions.limit = _defaultDocsLimit;
            selectOptions.creationLastDays = 30;
            selectOptions.optionsMask = 0;
            selectOptions.nameContains = null;
            selectOptions.offset = 0;
            selectOptions.createdFrom = DateTime.Now.AddDays(-30);
            selectOptions.createdTo = DateTime.Now;

            return selectOptions;
        }

        private int DefineOptionsMaskByControls()
        {
            int mask = 0;

            if (_layoutItemHelper.CreationDateFromPropertyItem.ControlValue != null)
            {
                mask |= 1;
            }

            if (_layoutItemHelper.CreationDateToPropertyItem.ControlValue != null)
            {
                mask |= 2;
            }

            //if ((bool)CreationLastDaysFlagItem.ControlValue)
            //{
            //    mask = 4;
            //}

            _logger.Debug("Defined optionsMask: " + mask);
            return mask;
        }

        private static string GetExcelSavePath()
        {
            string fileName = "Созданные мной документы " + DateTime.Now.ToString("HH-mm-ss_dd.MM.yyyy") + ".xlsx";
            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Personal),
                fileName
            );

            if (System.IO.File.Exists(path))
                System.IO.File.Delete(path);

            return path;
        }

        private static string GetImageSavePath(string name, string format)
        {
            if (name == null)
            {
                name = "Image";
            }

            if (format == null)
            {
                format = "";
            }

            string fileName = name + " " + DateTime.Now.ToString("HH-mm-ss_dd.MM.yyyy") + "." + format;
            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Personal),
                fileName
            );

            if (System.IO.File.Exists(path))
                System.IO.File.Delete(path);

            return path;
        }


        private class SqlServerExtensionResolver
        {
            private SimpleLogger _logger = new SimpleLogger();
            private string _providerName;
            private DocsVision.Platform.ObjectManager.ExtensionManager _extManager;

            public SqlServerExtensionResolver(string providerName, DocsVision.Platform.ObjectManager.ExtensionManager extensionManager)
            {
                _providerName = providerName;

                if (extensionManager == null)
                {
                    throw new NullReferenceException("ExtensionManager reference is not set");
                }
                _extManager = extensionManager;
            }

            public void AttachLogger(SimpleLogger logger)
            {
                if (logger != null) _logger = logger;
            }

            public string GetUserDocuments(DbSelectDocsOptions selectOptions)
            {
                _logger.Trace("GetDocumentsFromSelectResponse(DbSelectOptions selectOptions)");
                //_logger.Trace(string.Format("UserId={0}, OptionsMask={1}, Offset={2}, Limit={3}, CreatedFrom={4}, CreatedTo={5}, LastDays={6}, NameContains={7}",
                //    selectOptions.userId, selectOptions.optionsMask, selectOptions.offset, selectOptions.limit, selectOptions.createdFrom, selectOptions.createdTo, selectOptions.creationLastDays, selectOptions.nameContains));

                var procData = DbProcedureRegistry.GetProcedureInfo(DbProcedureRegistry.Procedure.GetDocumentsInfo);
                DbCommand cmd = DbCommand.CreateAsProcedureExec(procData.Item1, procData.Item2);
                cmd.Params.Add("UserId", selectOptions.userId, SqlDbType.UniqueIdentifier);
                cmd.Params.Add("Offset", selectOptions.offset, SqlDbType.Int);
                cmd.Params.Add("Limit", selectOptions.limit, SqlDbType.Int);
                cmd.Params.Add<string>("CreatedFrom", selectOptions.createdFrom.ToString("yyyy-MM-dd HH:mm:ss"), SqlDbType.DateTime);
                cmd.Params.Add<string>("CreatedTo", selectOptions.createdTo.ToString("yyyy-MM-dd HH:mm:ss"), SqlDbType.DateTime);
                cmd.Params.Add("NameContains", selectOptions.nameContains, SqlDbType.NVarChar);
                cmd.Params.Add("LastDays", selectOptions.creationLastDays, SqlDbType.Int);
                cmd.Params.Add("OptionsMask", selectOptions.optionsMask, SqlDbType.Int);

                return ExecuteSQLCommandCore(cmd);
            }

            public string GetFiles(Guid documentId)
            {
                _logger.Trace("GetFiles(Guid documentId); id = " + documentId.ToString());

                var procData = DbProcedureRegistry.GetProcedureInfo(DbProcedureRegistry.Procedure.GetFilesInfo);
                DbCommand cmd = DbCommand.CreateAsProcedureExec(procData.Item1, procData.Item2);
                cmd.Params.Add("DocId", documentId, SqlDbType.UniqueIdentifier);

                return ExecuteSQLCommandCore(cmd);
            }

            public string GetReceiversInternal(Guid documentId)
            {
                var procData = DbProcedureRegistry.GetProcedureInfo(DbProcedureRegistry.Procedure.GetReceiversInternalInfo);
                DbCommand cmd = DbCommand.CreateAsProcedureExec(procData.Item1, procData.Item2);
                cmd.Params.Add("DocId", documentId, SqlDbType.UniqueIdentifier);

                return ExecuteSQLCommandCore(cmd);
            }

            public string GetReceiversCounterparty(Guid documentId)
            {
                var procData = DbProcedureRegistry.GetProcedureInfo(DbProcedureRegistry.Procedure.GetReceiversCounterpartyInfo);
                DbCommand cmd = DbCommand.CreateAsProcedureExec(procData.Item1, procData.Item2);
                cmd.Params.Add("DocId", documentId, SqlDbType.UniqueIdentifier);

                return ExecuteSQLCommandCore(cmd);
            }

            public string ExecuteSQLCommandCore(DbCommand command)
            {
                _logger.Trace("ExecuteSQLCommandCore begin");

                if (command == null)
                {
                    throw new ArgumentNullException("Command reference is not set");
                }

                _logger.Debug("Command text (from new line):" + Environment.NewLine + command.CommandText);

                DocsVision.Platform.ObjectManager.ExtensionMethod extM = _extManager.GetExtensionMethod(_providerName, _extMethodName);
                extM.Parameters.AddNew("commandtext", ParameterValueType.String, command.CommandText);

                _logger.Trace("ExecuteSQLCommandCore end");
                return BaseResolver.DecompressString(extM.Execute() as string);
            }
        }

        private static string GetLocalizedFileTypeName(DocumentFileType fileType)
        {
            switch (fileType)
            {
                case DocumentFileType.Main: return "Основной";
                case DocumentFileType.Additional: return "Дополнительный";
                default: return "";
            }
        }

        private System.Windows.Forms.Label AddLoadingOverlay(DevExpress.XtraGrid.Views.Grid.GridView gridView)
        {
            _logger.Trace("AddLoadingOverlay begin");

            // Создаем Label для сообщения
            var loadingLabel = new System.Windows.Forms.Label
            {
                Text = "Идет загрузка, пожалуйста подождите...",
                Font = new System.Drawing.Font("Segoe UI", 12, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.DarkGray,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240, 240), // Полупрозрачный фон
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Dock = System.Windows.Forms.DockStyle.Fill,
                Visible = true
            };

            // Находим родительский контейнер GridControl
            var gridControl = gridView.GridControl;
            var parentControl = gridControl.Parent;

            // Добавляем Label поверх GridControl
            parentControl.Controls.Add(loadingLabel);
            loadingLabel.BringToFront();

            // Привязываем размеры Label к GridControl
            loadingLabel.Location = gridControl.Location;
            loadingLabel.Size = gridControl.Size;

            _logger.Trace("AddLoadingOverlay end");
            return loadingLabel;
        }

        private void EnableSelectCommands(bool enable)
        {
            _logger.Trace("EnableSelectCommands begin. Enable: " + enable);

            //CustomizableControl.LayoutControl.BeginUpdate();

            _exportToExcelBtn.Enabled = enable;
            _searchBtn.Enabled = enable;

            //CustomizableControl.LayoutControl.EndUpdate();
            _logger.Trace("EnableSelectCommands end");
        }

        private void SetDefaultValues()
        {
            _logger.Trace("SetDefaultValues begin");

            _userId = CurrentUserId;
            _layoutItemHelper.AuthorPropertyItem.ControlValue = _userId;
            _layoutItemHelper.CreationDateFromPropertyItem.ControlValue = DateTime.Now.AddDays(-30).Date;
            _layoutItemHelper.CreationDateToPropertyItem.ControlValue = DateTime.Now;
            _layoutItemHelper.DocumentsLimitPropertyItem.ControlValue = _defaultDocsLimit;

            _logger.Trace("SetDefaultValues end");
        }

        private bool ValidateCreationDates()
        {
            if (_layoutItemHelper.CreationDateFromPropertyItem.ControlValue != null && _layoutItemHelper.CreationDateToPropertyItem.ControlValue != null)
            {
                if ((DateTime)_layoutItemHelper.CreationDateFromPropertyItem.ControlValue > (DateTime)_layoutItemHelper.CreationDateToPropertyItem.ControlValue)
                {
                    return false;
                }
            }

            return true;
        }

        private void SetupUserRoleBasedControls()
        {
            bool isUserAdmin = (bool)Session.Properties[SessionPropertyNames.IsAdminSessionPropertyName].Value;
            _logger.Trace(SessionPropertyNames.IsAdminSessionPropertyName + "=" + isUserAdmin);
            if (isUserAdmin)
            {
                var authorItem = _layoutItemHelper.AuthorControl;
                authorItem.Enabled = true;

                var meButton = _layoutItemHelper.MeButtonItem;
                //meButton.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                var limitItem = _layoutItemHelper.LimitItem;
                limitItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }
            else
            {
                var authorItem = _layoutItemHelper.AuthorControl;
                authorItem.Enabled = false;

                var meButton = _layoutItemHelper.MeButtonItem;
                meButton.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;

                var limitItem = _layoutItemHelper.LimitItem;
                limitItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;
            }
        }

        [Obsolete]
        private void SetFormatForDateTimeColumns(DevExpress.XtraGrid.Views.Grid.GridView gridView, string format, DevExpress.Utils.FormatType formatType)
        {

            foreach (DevExpress.XtraGrid.Columns.GridColumn column in gridView.Columns)
            {
                if (column.ColumnType == typeof(DateTime))
                {
                    column.DisplayFormat.FormatString = format;
                    column.DisplayFormat.FormatType = formatType;
                }
            }
        }

        private TLayoutItem GetXLayoutItem<TLayoutItem>(string name) where TLayoutItem : DevExpress.XtraLayout.BaseLayoutItem
        {
            var items = CustomizableControl.LayoutControl.Items;
            if (items == null)
            {
                throw new Exception("ICustomizableControl.LayoutControl.Items is null");
            }

            int n = items.Count;
            TLayoutItem layoutItem = null;
            for (int i = 0; i < n; i++)
            {
                if (items[i] is TLayoutItem)
                {
                    TLayoutItem castedItem = (TLayoutItem)items[i];
                    if (castedItem.Name == name)
                    {
                        layoutItem = castedItem;
                        break;
                    }
                }
            }

            if (layoutItem == null)
            {
                throw new Exception((typeof(TLayoutItem)) + " wasn't found with name=" + name);
            }

            return layoutItem;
        }

        private bool ValidateFilters(StringBuilder messageBuilder)
        {
            if (_currentSelectOptions.optionsMask > 4)
            {
                _logger.Error("optionsMask is greater than 4. The value = " + _currentSelectOptions.optionsMask);
                throw new Exception("optionsMask is greater than 4. The value = " + _currentSelectOptions.optionsMask);
            }

            if ((_currentSelectOptions.optionsMask & 3) == 3 && !ValidateCreationDates())
            {
                messageBuilder.AppendLine("Дата в поле \"Дата создания с\" должна быть раньше или совпадать с датой в поле \"Дата создания по\".");
                return false;
            }

            return true;
        }

        [Obsolete]
        private bool TryGetSettingFromUserProfileCard(out UserProfileCardData data)
        {
            _logger.Trace("TryGetSettingsDataFromUserProfileCard begin. SettingId=" + SettingId + " ObjectId=" + SettingObjectId);

            IUserProfileCardService svc = CardControl.ObjectContext.GetService<IUserProfileCardService>();

            object setting = svc.GetSetting(SettingId, null, SettingObjectId);
            data = null;
            bool result = false;

            if (setting != null)
            {
                try
                {
                    data = _serializationStrategy.Deserialize<UserProfileCardData>(setting as string);
                    result = true;

                    _logger.Trace("Setting was found and deserialized");
                }
                catch (Exception ex)
                {
                    _logger.Error("Setting found, but error occured on deserialization." + Environment.NewLine + ex.ToString());
                }
            }
            else
            {
                _logger.Trace("Setting not found");
            }

            _logger.Trace("TryGetSettingsDataFromUserProfileCard end");

            return result;
        }

        [Obsolete]
        private bool TrySetSettingToUserProfileCard(UserProfileCardData data)
        {
            _logger.Trace("TrySetSettingsDataToUserProfileCard begin. SettingId=" + SettingId + " ObjectId=" + SettingObjectId);

            bool result = false;

            if (data != null)
            {
                IUserProfileCardService svc = CardControl.ObjectContext.GetService<IUserProfileCardService>();

                string serializedData = _serializationStrategy.Serialize(data);

                try
                {
                    svc.SetSetting(SettingId, SettingObjectId, serializedData);
                    result = true;
                }
                catch (Exception ex)
                {
                    _logger.Error("Error occured on SetSetting." + " Serialized data (from new line):" + Environment.NewLine + serializedData + Environment.NewLine + ex.ToString());
                }
            }
            else
            {
                _logger.Trace("CardSettingsData reference is not set");
            }

            _logger.Trace("TrySetSettingsDataToUserProfileCard end");

            return result;
        }

        private int AddFilesLinksToTable(Guid documentId)
        {
            int rowCount = 0;

            if (documentId != Guid.Empty)
            {
                var gridWrapper = _xtraGridRepository[Table.FilesFromUserDocument] as FilesFromUserDocsGridWrapper;
                gridWrapper.LoadTableData(out rowCount, documentId);
            }

            return rowCount;
        }

        #endregion

        #region Event Handlers Own

        private void OnSelectedDocumentIdChanged(Guid documentId)
        {
            _logger.Trace("OnSelectedDocumentIdChanged begin. Id:" + documentId);

            var filesGroup = _layoutItemHelper.FilesLayoutGroup;

            filesGroup.BeginUpdate();
            if (documentId != Guid.Empty)
            {
                filesGroup.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                int addedLinks = AddFilesLinksToTable(documentId);
                _logger.Trace("Files count: " + addedLinks);

                var filesTableLayout = _layoutItemHelper.TableFilesItem;

                if (addedLinks > 0)
                {
                    filesGroup.Text = "Файлы (кол-во: " + addedLinks + ")";
                    filesGroup.Expanded = true;
                    filesGroup.ExpandButtonVisible = true;

                    var filesView = _xtraGridRepository[Table.FilesFromUserDocument].MainView;
                    if (filesView != null)
                    {
                        filesView.FocusedRowHandle = 0;
                        filesView.MakeRowVisible(0);
                    }
                    else
                    {
                        _logger.Error("FilesView reference is not set");
                    }

                    filesTableLayout.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                }
                else
                {
                    filesGroup.Text = "Нет файлов в выбранном документе";
                    filesGroup.Expanded = false;
                    filesGroup.ExpandButtonVisible = false;

                    filesTableLayout.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;
                }
            }
            else
            {
                filesGroup.Text = "Файлы";
                filesGroup.Expanded = false;
                filesGroup.ExpandButtonVisible = false;
                filesGroup.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }

            filesGroup.EndUpdate();

            CustomizableControl.LayoutControl.PerformLayout();

            var docsView = _xtraGridRepository[Table.UserDocuments].MainView;
            docsView.MakeRowVisible(docsView.FocusedRowHandle);

            _logger.Trace("OnSelectedDocumentIdChanged end");
        }

        private void HandleFilePreview(Guid fileId)
        {
            _logger.Trace("HandleFilePreview begin. FileId:" + fileId);

            if (fileId == Guid.Empty)
            {
                _layoutItemHelper.PreviewableFileHeaderPropertyItem.ControlValue = "";

                _layoutItemHelper.FilePreviewLayoutGroup.Expanded = false;
                //_layoutItemHelper.FilePreviewDVControl.Preview(Guid.Empty); //do nothing
                _layoutItemHelper.FilePreviewLayoutGroup.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;
            }
            else
            {
                var filesViewHandler = (_xtraGridRepository[Table.FilesFromUserDocument] as FilesFromUserDocsGridWrapper).ViewHandler;
                _layoutItemHelper.PreviewableFileHeaderPropertyItem.ControlValue = filesViewHandler.GetSelectedFileInfo().name;

                _layoutItemHelper.FilePreviewLayoutGroup.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                _layoutItemHelper.FilePreviewLayoutGroup.Expanded = true;
                _layoutItemHelper.FilePreviewDVControl.Preview(fileId);
            }

            _logger.Trace("HandleFilePreview end");
        }

        private void OnFilePreviewPreviewCompleted(object sender, EventArgs e)
        {
            _logger.Trace("Preview completed");

            //FilePreviewXGroup.Update(); // ?
            _filePreviewCtrl.Update();
            _filePreviewCtrl.Refresh();
        }

        #endregion

        #region Event Handlers From Designer

        private void CardDocument_CardAddHandlers(System.Object sender, System.EventArgs e)
        {
            Session.LockManager.GetLockableObject(CardData.Id).ForceUnlock();

            CardControl.Caption = frameCaption;
        }

        private ImageManager _imgManager = null;

        private void CardDocument_CardActivated(object sender, DocsVision.Platform.WinForms.CardActivatedEventArgs e)
        {
            _logger.Info("CardDocument_CardActivated begin");

            if (!_cardActivated)
            {
                int i = 0;

                if (_layoutItemHelper != null)
                {
                    _logger.Warn("LayoutItemHelper reference already set");
                }
                else
                {
                    _layoutItemHelper = LayoutItemHelper.BuildDefault(CustomizableControl);
                    _logger.Trace("LayoutItemHelper instantiated");
                }

                if (_imgManager != null)
                {
                    _logger.Warn("ImageManager reference wasn't null");
                    _imgManager.Dispose();
                }
                else
                {
                    _logger.Trace("ImageManager reference is not set");
                }
                _imgManager = new ImageManager();

                _cardUrlHelper = new CardUrlHelper(ServerUrl);

                _sqlServerExtResolver = new SqlServerExtensionResolver(SEName, ExtManager);
                _sqlServerExtResolver.AttachLogger(new SimpleLogger(typeof(SqlServerExtensionResolver).Name, _loggerFilePath, _enableLogging));

                FilesFromUserDocsGridViewHandler filesGridViewHandler = new FilesFromUserDocsGridViewHandler(_layoutItemHelper.TableFilesGridControl, _imgManager, _tableFilesRowHeight, _tableFilesFileIconSize);
                FilesFromUserDocsGridWrapper filesFromUserDocGridWrapper = new FilesFromUserDocsGridWrapper(_layoutItemHelper.TableFilesGridControl, filesGridViewHandler, _sqlServerExtResolver);
                filesGridViewHandler.AttachLogger(new SimpleLogger(typeof(FilesFromUserDocsGridViewHandler).Name, _loggerFilePath, _enableLogging));
                filesFromUserDocGridWrapper.AttachLogger(new SimpleLogger(typeof(FilesFromUserDocsGridWrapper).Name, _loggerFilePath, _enableLogging));

                GridViewMenuItemRepository docsGridMenuItemRepo = new GridViewMenuItemRepository();
                UserDocsGridViewHandler userDocsGridViewHandler = new UserDocsGridViewHandler(_layoutItemHelper.TableDocsGridControl, _cardUrlHelper, docsGridMenuItemRepo, filesGridViewHandler, _imgManager, _layoutItemHelper);
                UserDocsGridWrapper userDocsGridWrapper = new UserDocsGridWrapper(_layoutItemHelper.TableDocsGridControl, userDocsGridViewHandler, _sqlServerExtResolver, _cardUrlHelper, filesFromUserDocGridWrapper, _layoutItemHelper);
                userDocsGridWrapper.AttachLogger(new SimpleLogger("UserDocumentsGridWrapper", _loggerFilePath, _enableLogging));
                userDocsGridViewHandler.AttachLogger(new SimpleLogger("UserDocumentsGridViewHandler", _loggerFilePath, _enableLogging));

                _xtraGridRepository = new XtraGridRepository();
                _xtraGridRepository.AddGrid(Table.UserDocuments, userDocsGridWrapper);
                _xtraGridRepository.AddGrid(Table.FilesFromUserDocument, filesFromUserDocGridWrapper);
                _xtraGridRepository.InitializeAll();

                filesFromUserDocGridWrapper.ViewHandler.SelectedFileIdChanged += HandleFilePreview;
                userDocsGridWrapper.ViewHandler.SelectedDocumentIdChanged += OnSelectedDocumentIdChanged;
                userDocsGridWrapper.TableLoaded += OnUserDocsGrid_TableLoaded;
                userDocsGridWrapper.UnloadedToExcel += OnUserDocsGridWrapper_UnloadedToExcel;

                _layoutItemHelper.FilePreviewDVControl.PreviewCompleted += OnFilePreviewPreviewCompleted;

                _logger.Debug(i++ + "");

                _userId = CurrentUserId;
                _logger.Trace("Current user id=" + CurrentUserId);

                _currentSelectOptions = GetDefaultSelectOptions();

                //_logger.Debug(i++ + "");

                SetDefaultValues();

                //_logger.Debug(i++ + "");

                SetupUserRoleBasedControls();

                //_logger.Debug(i++ + "");

                //_logger.Debug(i++ + "");

                //_logger.Trace("Before SE");
                //_logger.Trace("SEName:" + SEName);

                //_logger.Debug(i++ + "");

                _layoutItemHelper.FilePreviewLayoutGroup.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;
                _layoutItemHelper.FilesLayoutGroup.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;

                _exportToExcelBtn = _layoutItemHelper.ExportToExcelButton;
                _searchBtn = _layoutItemHelper.SearchButton;

                EnableSelectCommands(true);

                ClearExtraButtonsOnDateEdits();
                AddExtraButtonsToDateEdits();

                _dataController = new UserProfileCardDataController(_serializationStrategy, CardControl.ObjectContext);
                if (!_dataController.TryLoad(out _userProfileCardData))
                {
                    _userProfileCardData = new UserProfileCardData()
                    {
                        documentsLimit = _currentSelectOptions.limit
                    };
                    _logger.Trace("Data not loaded");
                }
                else
                {
                    _layoutItemHelper.DocumentsLimitPropertyItem.ControlValue = _userProfileCardData.documentsLimit;
                    _logger.Trace("Data loaded");
                }

                _filePreviewCtrl = _layoutItemHelper.FilePreviewControl;
                if (_filePreviewCtrl == null)
                {
                    _logger.Error("_filePreviewCtrl reference is not set");
                }

                TEST();

                _cardActivated = true;
            }
            else
            {

            }

            LogMemoryUsage("CardDocument_CardActivated end");
        }

        private void OnUserDocsGridWrapper_UnloadedToExcel(UnloadToExcelResultInfo resultInfo)
        {
            _logger.Trace("OnUserDocsGridWrapper_UnloadedToExcel begin. Rows:" + resultInfo.rowCount + " Result:" + resultInfo.result);

            EnableSelectCommands(true);

            switch (resultInfo.result)
            {
                case UnloadToExcelResultInfo.Result.Success:
                    var messageResult = CardControl.ShowMessage("Файл с документами сохранен в пути: " + resultInfo.pathToFile + "\nОткрыть файл?", "Результат выгрузки",
                            DocsVision.Platform.CardHost.MessageType.Question, DocsVision.Platform.CardHost.MessageButtons.YesNo);

                    if (messageResult == DocsVision.Platform.CardHost.MessageResult.Yes)
                    {
                        Process.Start(resultInfo.pathToFile);
                    }

                    break;

                case UnloadToExcelResultInfo.Result.NoData:
                    CardControl.ShowMessage("По указанному периоду создания не найдены документы.", "Результат выгрузки",
                            DocsVision.Platform.CardHost.MessageType.Question, DocsVision.Platform.CardHost.MessageButtons.Ok);

                    break;

                case UnloadToExcelResultInfo.Result.Cancelled:
                    break;

                case UnloadToExcelResultInfo.Result.Fail:
                    _logger.Error("Unload result = fail. RowCount=" + resultInfo.rowCount);

                    break;
            }

            _logger.Trace("OnUserDocsGridWrapper_UnloadedToExcel end");
        }

        private void OnUserDocsGrid_TableLoaded(DataTable table)
        {
            _logger.Trace("OnUserDocsGrid_TableLoaded begin");

            EnableSelectCommands(true);

            if (table == null || table.Rows.Count == 0)
            {
                CardControl.ShowMessage("Не найдены документы.");
            }

            _logger.Trace("OnUserDocsGrid_TableLoaded end");
        }

        private void ClearExtraButtonsOnDateEdits()
        {
            for (int i = _layoutItemHelper.CreationDateFromDateEdit.Properties.Buttons.Count - 1; i > 0; i--)
                _layoutItemHelper.CreationDateFromDateEdit.Properties.Buttons.RemoveAt(i);
            for (int i = _layoutItemHelper.CreationDateToDateEdit.Properties.Buttons.Count - 1; i > 0; i--)
                _layoutItemHelper.CreationDateToDateEdit.Properties.Buttons.RemoveAt(i);
        }

        private void AddExtraButtonsToDateEdits()
        {
            _logger.Trace("AddExtraButtonsToDateEdits begin");

            _dateEditorButtonsHandler.AddTo(_layoutItemHelper.CreationDateFromDateEdit, new EditorButton(ButtonPredefines.Undo), (sender, e) => { _layoutItemHelper.CreationDateFromPropertyItem.ControlValue = DateTime.Now.AddDays(-30).Date; });
            _dateEditorButtonsHandler.AddTo(_layoutItemHelper.CreationDateFromDateEdit, new EditorButton(ButtonPredefines.Clear), (sender, e) => { _layoutItemHelper.CreationDateFromPropertyItem.ControlValue = null; });

            _dateEditorButtonsHandler.AddTo(_layoutItemHelper.CreationDateToDateEdit, new EditorButton(ButtonPredefines.Undo), (sender, e) => { _layoutItemHelper.CreationDateToPropertyItem.ControlValue = DateTime.Now; });
            _dateEditorButtonsHandler.AddTo(_layoutItemHelper.CreationDateToDateEdit, new EditorButton(ButtonPredefines.Clear), (sender, e) => { _layoutItemHelper.CreationDateToPropertyItem.ControlValue = null; });

            _logger.Trace("AddExtraButtonsToDateEdits end");
        }

        private void CardDocument_CardClosing(System.Object sender, DocsVision.Platform.WinForms.CardClosingEventArgs e)
        {
            _logger.Trace("CardDocument_CardClosing begin");

            _cardClosing = true;

            _layoutItemHelper.FilePreviewDVControl.PreviewCompleted -= OnFilePreviewPreviewCompleted;

            var filesGridWrapper = _xtraGridRepository[Table.FilesFromUserDocument] as FilesFromUserDocsGridWrapper;
            var userDocsGridWrapper = _xtraGridRepository[Table.UserDocuments] as UserDocsGridWrapper;

            filesGridWrapper.ViewHandler.SelectedFileIdChanged -= HandleFilePreview;
            userDocsGridWrapper.ViewHandler.SelectedDocumentIdChanged -= OnSelectedDocumentIdChanged;

            _xtraGridRepository.Dispose();

            if (_imgManager != null)
            {
                _imgManager.Dispose();
                _imgManager = null;
            }
            else
            {
                _logger.Error("ImageManager reference is not set");
            }

            _dateEditorButtonsHandler.Dispose();
            ClearExtraButtonsOnDateEdits();

            _logger.Trace("CardDocument_CardClosing end");
        }

        private void CreationDateFrom_EditValueChanged(System.Object sender, System.EventArgs e)
        {
            _logger.Trace("CreationDateFrom_EditValueChanged");

            _currentSelectOptions.optionsMask = DefineOptionsMaskByControls();

            ILayoutPropertyItem fromItem = _layoutItemHelper.CreationDateFromPropertyItem;

            if (fromItem.ControlValue != null)
            {
                _currentSelectOptions.createdFrom = (DateTime)fromItem.ControlValue;
                //if (CreationDateToItem.ControlValue != null && (DateTime)CreationDateFromItem.ControlValue > (DateTime)CreationDateToItem.ControlValue)
                //{
                //    if ((bool)CreationDateFromFlagItem.ControlValue)
                //        EnableSelectCommands(false);

                //    CardControl.ShowMessage("Укажите начальную дату (\"Дата создания с\") так, чтобы она была не позже конечной даты (\"Дата создания по\")");
                //}
            }
            else
            {

            }
        }

        private void CreationDateTo_EditValueChanged(System.Object sender, System.EventArgs e)
        {
            _logger.Trace("CreationDateTo_EditValueChanged");

            _currentSelectOptions.optionsMask = DefineOptionsMaskByControls();

            ILayoutPropertyItem toItem = _layoutItemHelper.CreationDateToPropertyItem;

            if (toItem.ControlValue != null)
            {
                _currentSelectOptions.createdTo = (DateTime)toItem.ControlValue;
                //if (CreationDateFromItem.ControlValue != null && (DateTime)CreationDateFromItem.ControlValue > (DateTime)CreationDateToItem.ControlValue)
                //{
                //    if ((bool)CreationDateToFlagItem.ControlValue)
                //        EnableSelectCommands(false);

                //    CardControl.ShowMessage("Укажите конечную дату (\"Дата создания по\") так, чтобы она была не раньше начальной даты (\"Дата создания с\")");
                //}
            }
            else
            {

            }
        }

        private void MeButton_Click(System.Object sender, System.EventArgs e)
        {
            _logger.Trace("MeButton_Click");

            ILayoutPropertyItem author = _layoutItemHelper.AuthorPropertyItem;
            author.ControlValue = CurrentUserId;
            _currentSelectOptions.userId = (Guid)author.ControlValue;
        }

        private void Author_ControlValueChanged(System.Object sender, System.EventArgs e)
        {
            _logger.Trace("Author_ControlValueChanged");

            ILayoutPropertyItem author = _layoutItemHelper.AuthorPropertyItem;

            if (author.ControlValue != null && (Guid)author.ControlValue != _userId)
            {
                _userId = (Guid)author.ControlValue;
                _currentSelectOptions.userId = _userId;
            }

            _logger.Trace("Author_ControlValueChanged UserId=" + _userId.ToString("B"));
        }

        private void Limit_ValueChanged(System.Object sender, System.EventArgs e)
        {
            ILayoutPropertyItem limitItem = _layoutItemHelper.DocumentsLimitPropertyItem;

            if (limitItem.ControlValue == null)
            {
                _logger.Warn("LimitItem.ControlValue is null");
                return;
            }

            int limit = (int)limitItem.ControlValue;
            if (limit < 0) limit = 0;
            if (limit == 0)
            {
                _currentSelectOptions.limit = int.MaxValue;
            }
            else
            {
                _currentSelectOptions.limit = limit;
            }

            // Сохранение данных
            _userProfileCardData.documentsLimit = _currentSelectOptions.limit;
            if (_dataController.TrySave(_userProfileCardData))
            {
                _logger.Trace("Data saved");
            }
            else
            {
                _logger.Error("Data wasn't saved");
            }
        }

        private void Search_Click(System.Object sender, System.EventArgs e)
        {
            _logger.Info("Search_Click begin");

            _currentSelectOptions.optionsMask = DefineOptionsMaskByControls();

            var gridWrapper = _xtraGridRepository[Table.UserDocuments] as UserDocsGridWrapper;

            try
            {
                if (gridWrapper.LoadTableDataWorkerWrapper.Worker.IsBusy)
                {
                    _logger.Warn("Поиск уже выполняется.");
                    return;
                }

                if (!ValidateCreationDates())
                {
                    CardControl.ShowMessage("Укажите период создания так, чтобы 'Дата создания с' было не позднее 'по'");
                    return;
                }

                EnableSelectCommands(false);

                _layoutItemHelper.FilesLayoutGroup.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;

                gridWrapper.LoadTableDataAsync(_currentSelectOptions);
            }
            catch (Exception ex)
            {
                _logger.Error(ex.ToString());
            }
            finally
            {
                _logger.Info("Search_Click end");
            }
        }

        private void ExportDocumentsToExcel_Click(System.Object sender, System.EventArgs e)
        {
            _logger.Info("ExportDocumentsToExcel_Click");

            var gridWrapper = _xtraGridRepository[Table.UserDocuments] as UserDocsGridWrapper;

            if (gridWrapper.UnloadToExcelWorkerWrapper.Worker.IsBusy)
                return;

            if (!ValidateCreationDates())
            {
                CardControl.ShowMessage("Укажите период создания так, чтобы 'Дата создания с' было не позднее 'по'");
                return;
            }
            else
            {
                EnableSelectCommands(false);

                gridWrapper.UnloadToExcelAsync(_currentSelectOptions);
            }
        }

        // TO DO
        private void ColumnVisibilitySwitch_Click(System.Object sender, System.EventArgs e)
        {
            _logger.Trace("ColumnVisibilitySwitch_Click begin");

            var view = _xtraGridRepository[Table.UserDocuments].MainView;
            if (view != null)
            {
                DevExpress.XtraEditors.SimpleButton switchButton = _layoutItemHelper.ColumnVisibilitySwitchButton;

                DXPopupMenu menu = ContextMenuHelper.CreateColumnSelectorPopupMenu(view);

                //Point location = new Point(switchButton.Bounds.Right, switchButton.Bounds.Top);
                //Point location2 = new Point(switchButton.Width, 0);
                Point location3 = new Point(0, 0);
                menu.ShowPopup(switchButton, location3);
            }
            else
            {
                _logger.Error("view reference is not set");
            }

            _logger.Trace("ColumnVisibilitySwitch_Click end");
        }

        private void GroupPanelVisibilitySwitch_Click(System.Object sender, System.EventArgs e)
        {
            _logger.Trace("GroupPanelVisibilitySwitch_Click begin");

            var view = _xtraGridRepository[Table.UserDocuments].MainView;
            if (view != null)
            {
                view.OptionsView.ShowGroupPanel = !view.OptionsView.ShowGroupPanel;

                DevExpress.XtraEditors.SimpleButton showGroupPanelButton = _layoutItemHelper.GetControl<DevExpress.XtraEditors.SimpleButton>(LayoutItemAlias.GroupPanelVisibilitySwitch);
                if (view.OptionsView.ShowGroupPanel)
                {
                    showGroupPanelButton.Text = "Скрыть панель группировки";
                }
                else
                {
                    showGroupPanelButton.Text = "Показать панель группировки";
                }
            }
            else
            {
                _logger.Error("view reference is not set");
            }

            _logger.Trace("GroupPanelVisibilitySwitch_Click end");
        }

        private void ClearGrouping_Click(System.Object sender, System.EventArgs e)
        {
            _logger.Trace("ClearGrouping_Click begin");

            var view = _xtraGridRepository[Table.UserDocuments].MainView;
            if (view != null)
            {
                view.ClearGrouping();
            }
            else
            {
                _logger.Error("view reference is not set");
            }

            _logger.Trace("ClearGrouping_Click end");
        }

        #endregion

        #region Event Handlers Test

        private void ExportLayoutXML_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //CustomizableControl.LayoutControl.SaveLayoutToXml(_pathToLayoutXML);
        }

        private void LogUI_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("LogUI_ItemClick begin");

            LogUI(_layoutItemHelper.FilesLayoutGroup);

            _logger.Trace("LogUI_ItemClick end");
        }

        private void LogUIRoot_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LogUI();
        }

        private void TraceFilesInfo_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("TraceFilesInfo_ItemClick begin");

            try
            {
                Document document = SelectedDocument;
                var files = document.Files;

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("TRACE FILES");

                foreach (var file in files)
                {
                    //FileData data = CardControl.Session.FileManager.GetFile(file.FileId);
                    sb.AppendLine(string.Format("Name:{0} FileId:{1} FileVId:{2} FileVRowId:{3}", file.FileName, file.FileId, file.FileVersionId, file.FileVersionRowId));
                }

                _logger.Debug(sb.ToString());
            }
            catch (Exception ex)
            {
                _logger.Error(ex.ToString());
            }

            _logger.Trace("TraceFilesInfo_ItemClick end");
        }

        private void GetDataFromUserProfileCard_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("GetDataFromUserProfileCard_ItemClick begin");



            _logger.Trace("GetDataFromUserProfileCard_ItemClick end");
        }

        private void SetDataToUserProfileCard_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("SetDataToUserProfileCard_ItemClick begin");



            _logger.Trace("SetDataToUserProfileCard_ItemClick end");
        }

        private void testRef_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //Хренов
            // Card ID F25B53F8-714E-4075-BD38-EABE52AC31ED
            //ReferenceList_FIELD_REFCARDID E16D5C4D-E9ED-4D6A-8891-F27030E920A2

            Guid refListId = new Guid("E16D5C4D-E9ED-4D6A-8891-F27030E920A2");
            Guid cardId = new Guid("F25B53F8-714E-4075-BD38-EABE52AC31ED");
            var refSvc = CardControl.ObjectContext.GetService<IReferenceListService>();
            ReferenceList refList;
            bool refFromCardResult = refSvc.TryGetReferenceListFromCard(cardId, false, out refList);

            _logger.Trace("refFromCardResult=" + refFromCardResult);

            //var xz = CardControl.ObjectContext.GetService<DocsVision.BackOffice.ObjectModel.Services.IBaseUniversalService>();
            //[dbo].[dd_Get_ExecutionTaskList]
            /*
             * <ExecutionTaskList>
  <TaskInfo TaskID="F911FD9A-FB0A-4722-8426-B383F989A864" Level="1" StartTaskDate="2024-05-20T13:30:11" Author="Хренов С.Н./ООО &quot;ПАЗ&quot;/Начальник отдела" Performer="Васильев А.В./ПАО &quot;Павловский автобус&quot;/Генеральный директор" CurPerformers="Огурцов А. А./ООО &quot;Павловский автобусный завод&quot;/Директор по операционной деятельности" CompletedUser="Огурцов А. А./ООО &quot;Павловский автобусный завод&quot;/Директор по операционной деятельности" KindName="На ознакомление" TaskContent="Вам необходимо ознакомиться с СТО" ExpectedEndDate="2024-05-23T13:00:00" ActualEndDate="2024-05-24T15:21:24" StateTask="Завершено" />
  <TaskInfo TaskID="AF127EF4-5A4A-4CCF-9246-CD5EB4ABB729" Level="2" StartTaskDate="2024-05-24T15:23:46" Author="Хренов С.Н./ООО &quot;ПАЗ&quot;/Начальник отдела" Performer="Хорьков А.И./ООО &quot;ЦИТРБ&quot;/Ведущий специалист по информационной безопасности, Зуев А.О./ООО &quot;ЦИТРБ&quot;/Главный специалист по технической защите информации, Бусарин Д.В./ООО &quot;ЦИТРБ&quot;/Специалист по информационной безопасности" CurPerformers="Бусарин Д.В./ООО &quot;ЦИТРБ&quot;/Специалист по информационной безопасности" CompletedUser="Бусарин Д.В./ООО &quot;ЦИТРБ&quot;/Специалист по информационной безопасности" KindName="По поручению" TaskContent="Вам поступила заявка на выдачу ЭЦП:&#xA;Сотрудник: Васильев А. В.&#xA;Имя ПК: W18-001086&#xA;Номер телефона: +79100589292&#xA;Тип ЭЦП: НЭП&#xA;Цели использования: Договоры и первичные документы" ActualEndDate="2024-05-24T15:26:21" Report="Ок" StateTask="Завершено" FileAuthor="00B50DCB-E827-4D20-BC1A-BAD5F7BEFA75" Description="Файл Инструкция (НЭП)_.docx" Card="3AA73C04-D8A1-4489-A5F4-1D8700D85172" />
  <TaskInfo TaskID="AF127EF4-5A4A-4CCF-9246-CD5EB4ABB729" Level="2" StartTaskDate="2024-05-24T15:23:46" Author="Хренов С.Н./ООО &quot;ПАЗ&quot;/Начальник отдела" Performer="Хорьков А.И./ООО &quot;ЦИТРБ&quot;/Ведущий специалист по информационной безопасности, Зуев А.О./ООО &quot;ЦИТРБ&quot;/Главный специалист по технической защите информации, Бусарин Д.В./ООО &quot;ЦИТРБ&quot;/Специалист по информационной безопасности" CurPerformers="Бусарин Д.В./ООО &quot;ЦИТРБ&quot;/Специалист по информационной безопасности" CompletedUser="Бусарин Д.В./ООО &quot;ЦИТРБ&quot;/Специалист по информационной безопасности" KindName="По поручению" TaskContent="Вам поступила заявка на выдачу ЭЦП:&#xA;Сотрудник: Васильев А. В.&#xA;Имя ПК: W18-001086&#xA;Номер телефона: +79100589292&#xA;Тип ЭЦП: НЭП&#xA;Цели использования: Договоры и первичные документы" ActualEndDate="2024-05-24T15:26:21" Report="Ок" StateTask="Завершено" FileAuthor="00B50DCB-E827-4D20-BC1A-BAD5F7BEFA75" Description="Файл Васильев Андрей Владимирович.7z" Card="4E22B375-9146-47C8-ACC8-55FE7A3240EE" />
  <TaskInfo TaskID="7B639F4D-43E7-4693-B90F-062A59C871CE" Level="3" StartTaskDate="2024-05-24T15:28:29" Author="Хренов С.Н./ООО &quot;ПАЗ&quot;/Начальник отдела" Performer="Васильев А.В./ПАО &quot;Павловский автобус&quot;/Генеральный директор" CurPerformers="Васильев А.В.//ООО &quot;ПАЗ&quot;/Генеральный директор" CompletedUser="Васильев А.В.//ООО &quot;ПАЗ&quot;/Генеральный директор" KindName="По поручению" TaskContent="Необходимо установить электронную подпись:&#xA;   1. Открыть лист исполнения. В нем вложены архив с файлом установки сертификата (в формате ФИО.7z) и инструкция по установке.&#xA;   2. Установить электронную подпись следуя шагам, описанным в инструкции.&#xA;   3. После успешной установки подтвердить получение подписи, заполнив отчет и нажав кнопку «Получен»." ExpectedEndDate="2024-05-29T15:28:00" ActualEndDate="2024-05-28T14:41:44" StateTask="Завершено" />
</ExecutionTaskList>
             */

            //_logger.Trace();

            if (refFromCardResult)
            {
                var refs = refList.References;
                _logger.Trace("ref count: " + refs.Count);
                foreach (var @ref in refs)
                {
                    _logger.Trace(@ref.CreationDate.ToString());
                }
            }
        }

        #endregion

        #region TEST

        private void TEST()
        {
            _logger.Trace("TEST SANDBOX begin");



            _logger.Trace("TEST SANDBOX end");
        }

        private struct DocumentFileInfo
        {
            public Guid fileId;
            public string name;

            public DocumentFileInfo(Guid fileId, string name)
            {
                this.fileId = fileId;
                this.name = name;
            }
        }

        private void testFilter_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            var view = _layoutItemHelper.TableDocsGridControl.MainView as DocsVision.BackOffice.WinForms.Controls.GridExView;

            var criteria = view.ActiveFilterCriteria;

            if (!view.ActiveFilterEnabled)
            {
                _logger.Warn("ActiveFilterEnabled = false");
            }

            List<DevExpress.XtraGrid.Views.Base.ViewColumnFilterInfo> filters = new List<DevExpress.XtraGrid.Views.Base.ViewColumnFilterInfo>();
            foreach (DevExpress.XtraGrid.Columns.GridColumn col in view.Columns)
            {
                var filter = view.ActiveFilter[col];
                if (filter != null && !filter.IsEmpty)
                {
                    filters.Add(filter);
                }
            }

            _logger.Log("Applied filters count: " + filters.Count);
            StringBuilder sB = new StringBuilder("FilterInfo from new line:");
            sB.AppendLine();
            foreach (var filterInfo in filters)
            {
                DevExpress.XtraGrid.Columns.ColumnFilterInfo info = filterInfo.Filter;
                if (info.Value != null)
                {
                    sB.AppendLine(string.Format("FieldName:{0} Filter.DisplayText:{1} Filter.Value:{2}, Filter.FilterString:{3}, Filter.Type:{4}, Filter.Kind:{5}, Filter.CriteriaType:{6}",
                        filterInfo.Column.FieldName, info.DisplayText, info.Value, info.FilterString,
                        info.Type, info.Kind, info.FilterCriteria.GetType().Name));
                }
                else
                {
                    sB.AppendLine(string.Format("FieldName:{0} Filter.DisplayText:{1}, Filter.Value:NULL, Filter.FilterString:{2}, Filter.Type:{3}, Filter.Kind:{4}, Filter.CriteriaType:{5}",
                        filterInfo.Column.FieldName, info.DisplayText, info.FilterString,
                        info.Type, info.Kind, info.FilterCriteria.GetType().Name));
                }
            }
            _logger.Log(sB.ToString());

            //BinaryOperator binaryOperator = new BinaryOperator();

            //_logger.Trace();
            //CardControl.ShowMessage(criteria.ToString());

            //var filters = ExtractFilters(criteria);
            //StringBuilder sB = new StringBuilder();
            //sB.AppendLine("Applied filters");
            //foreach (var f in filters)
            //{
            //    if (f.Value != null)
            //    {
            //        sB.AppendLine(f.FieldName + "|" + f.Operator + "|" + f.Value);
            //    }
            //    else
            //    {
            //        sB.AppendLine(f.FieldName + "|" + f.Operator);
            //    }
            //}

            //string log = sB.ToString();
            //_logger.Debug(log);
            //CardControl.ShowMessage(log);
        }

        private void LoadLayoutXML_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //_logger.Trace("LoadLayoutXML_ItemClick. Path: " + _pathToLayoutXML);

            //CustomizableControl.LayoutControl.RestoreLayoutFromXml(_pathToLayoutXML);

            //_logger.Trace("LoadLayoutXML_ItemClick end");
        }

        private void exportUserDocsView_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("exportUserDocsView_ItemClick begin");

            var gridWrapper = _xtraGridRepository[Table.UserDocuments] as UserDocsGridWrapper;
            gridWrapper.Control.ExportToXlsx(GetExcelSavePath());

            _logger.Trace("exportUserDocsView_ItemClick end");
        }

        private void ClearFilters_Click(System.Object sender, System.EventArgs e)
        {
            _logger.Trace("ClearFilters_Click begin");

            var gridWrapper = _xtraGridRepository[Table.UserDocuments] as UserDocsGridWrapper;
            var view = gridWrapper.ViewHandler.View;

            view.ClearColumnsFilter();

            _logger.Trace("ClearFilters_Click end");
        }

        private void CollapseExpandGroups_Click(System.Object sender, System.EventArgs e)
        {
            var gridWrapper = _xtraGridRepository[Table.UserDocuments] as UserDocsGridWrapper;
            var viewHandler = gridWrapper.ViewHandler;
            var view = viewHandler.View;

            bool isAllGroupRowsExpanded = Utils.IsAllGroupRowsExpanded(view);
            _logger.Trace("CollapseExpandGroups_Click begin. Action: " + (isAllGroupRowsExpanded ? "Collapse" : "Expand"));

            if (view.GroupCount != 0)
            {
                if (!isAllGroupRowsExpanded)
                {
                    view.ExpandAllGroups();
                }
                else
                {
                    view.CollapseAllGroups();
                }
            }
            else
            {
                _logger.Error("GroupCount == 0");
            }

            _logger.Trace("CollapseExpandGroups_Click end");
        }

        private void selectAndConvertImageToCompressedBase64String_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("selectAndConvertImageToBase64String_ItemClick begin");

            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Images only
            openFileDialog.Filter =
                "Изображения (*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.tiff;*.ico;*.svg)|" +
                "*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.tiff;*.ico;*.svg";

            openFileDialog.Title = "Выберите изображения";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);

            openFileDialog.CheckFileExists = true;
            openFileDialog.CheckPathExists = true;
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string[] imagePaths = openFileDialog.FileNames;
                int n = imagePaths.Length;
                if (n > 0)
                {
                    string textFile = imagePaths[0].Substring(0, imagePaths[0].LastIndexOf(System.IO.Path.DirectorySeparatorChar) + 1);
                    textFile = textFile + "Images data.txt";

                    if (System.IO.File.Exists(textFile))
                    {
                        System.IO.File.Delete(textFile);
                    }

                    // ! Если файл содержит множественные расширения, то это может сработать некорректно
                    //string textFile = imagePaths[0].Substring(0, imagePaths[0].LastIndexOf('.') + 1) + "txt";

                    _logger.Trace("Images data will be in file: " + textFile);

                    for (int i = 0; i < imagePaths.Length; i++)
                    {
                        System.IO.File.AppendAllText(textFile, "<File path=\"" + imagePaths[i] +
                            "\" base64String(Compressed)=\"" + Utils.ImageToCompressedBase64String(imagePaths[i]) + "\"" + Environment.NewLine + "/>" + Environment.NewLine);
                    }
                }
            }

            _logger.Trace("selectAndConvertImageToBase64String_ItemClick end");
        }

        private void selectAndConvertImageToBase64String_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("selectAndConvertImageToBase64String_ItemClick begin");

            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Images only
            openFileDialog.Filter =
                "Изображения (*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.tiff;*.ico;*.svg)|" +
                "*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.tiff;*.ico;*.svg";

            openFileDialog.Title = "Выберите изображения";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);

            openFileDialog.CheckFileExists = true;
            openFileDialog.CheckPathExists = true;
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string[] imagePaths = openFileDialog.FileNames;
                int n = imagePaths.Length;
                if (n > 0)
                {
                    string textFile = imagePaths[0].Substring(0, imagePaths[0].LastIndexOf(System.IO.Path.DirectorySeparatorChar) + 1);
                    textFile = textFile + "Images data.txt";

                    if (System.IO.File.Exists(textFile))
                    {
                        System.IO.File.Delete(textFile);
                    }

                    // ! Если файл содержит множественные расширения, то это может сработать некорректно
                    //string textFile = imagePaths[0].Substring(0, imagePaths[0].LastIndexOf('.') + 1) + "txt";

                    _logger.Trace("Images data will be in file: " + textFile);

                    for (int i = 0; i < imagePaths.Length; i++)
                    {
                        System.IO.File.AppendAllText(textFile, "<File path=\"" + imagePaths[i] +
                            "\" base64String=\"" + Utils.ImageToBase64String(imagePaths[i]) + "\"" + Environment.NewLine + "/>" + Environment.NewLine);
                    }
                }
            }

            _logger.Trace("selectAndConvertImageToBase64String_ItemClick end");
        }

        private static void TraceAndSaveItemsImagesFromGridViewMenu(DevExpress.XtraGrid.Menu.GridViewMenu menu, SimpleLogger logger)
        {
            if (logger == null)
            {
                logger = new SimpleLogger();
            }

            logger.Trace("TraceGridViewMenu begin");

            if (menu == null)
            {
                logger.Warn("GridViewMenu reference is not set");
            }
            else
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Caption: " + menu.Caption);

                var items = menu.Items;
                if (items != null)
                {
                    int n = items.Count;
                    logger.Trace(n + " items in menu");
                    for (int i = 0; i < n; i++)
                    {
                        if (items[i].ImageOptions != null)
                        {
                            Image img = items[i].ImageOptions.Image;
                            if (img != null)
                            {
                                img.Save(GetImageSavePath(items[i].Caption, "png"));
                            }

                            DevExpress.Utils.Svg.SvgImage svgImg = items[i].ImageOptions.SvgImage;
                            if (svgImg != null)
                            {
                                svgImg.Save(GetImageSavePath(items[i].Caption, "svg"));
                            }
                        }
                        else
                        {
                            sb.AppendLine("Image[" + i + "] ImageOptions referene is not set");
                        }
                    }
                }
                else
                {
                    logger.Warn("No items in menu");
                }

                logger.Trace(sb.ToString());
            }

            logger.Trace("TraceGridViewMenu end");
        }

        private enum Table { UserDocuments, FilesFromUserDocument, AllDocuments };

        private class XtraGridRepository
        {
            private Dictionary<Table, XtraGridWrapper> _tables;
            private bool _disposed = false;

            public XtraGridRepository()
            {
                _tables = new Dictionary<Table, XtraGridWrapper>();
            }

            public XtraGridWrapper this[Table table]
            {
                get
                {
                    if (!_tables.ContainsKey(table))
                    {
                        return null;
                    }

                    return _tables[table];
                }
            }

            public void AddGrid(Table table, XtraGridWrapper gridWrapper)
            {
                if (gridWrapper == null)
                {
                    return;
                }

                if (_tables.ContainsKey(table))
                {
                    if (_tables[table] == null)
                    {
                        _tables[table] = gridWrapper;
                    }
                }
                else
                {
                    _tables.Add(table, gridWrapper);
                }
            }

            public void AddGrid(Table table, XtraGridWrapper gridWrapper, bool initialize)
            {
                if (gridWrapper == null)
                {
                    return;
                }

                if (_tables.ContainsKey(table))
                {
                    if (_tables[table] == null)
                    {
                        _tables[table] = gridWrapper;
                    }
                }
                else
                {
                    _tables.Add(table, gridWrapper);
                }

                if (initialize) gridWrapper.Initilize();
            }

            public void InitializeAll()
            {
                foreach (XtraGridWrapper gridWrapper in _tables.Values)
                {
                    InitializeCore(gridWrapper);
                }
            }

            public void Initialize(Table table)
            {
                if (_tables.ContainsKey(table) && _tables[table] != null)
                {
                    InitializeCore(_tables[table]);
                }
            }

            private void InitializeCore(XtraGridWrapper gridWrapper)
            {
                gridWrapper.Initilize();
            }

            protected virtual void Dispose(bool disposing)
            {
                if (_disposed) return;

                if (disposing)
                {
                    // Освобождаем управляемые ресурсы
                    XtraGridWrapper[] tables = _tables.Values.ToArray();
                    int n = tables.Length;
                    for (int i = 0; i < n; i++)
                    {
                        if (tables[i] != null)
                        {
                            tables[i].Dispose();
                            tables[i] = null;
                        }
                    }
                }

                // Освобождаем неуправляемые ресурсы

                _disposed = true;
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            public void Dispose(Table table)
            {
                if (_tables.ContainsKey(table) && _tables[table] != null)
                {
                    _tables[table].Dispose();
                }
            }
        }

        protected abstract class XtraGridViewHandler<TView> : IDisposable where TView : GridView
        {
            protected bool _disposed = false;
            protected GridControl _gridControl;
            protected TView _view;

            public TView View
            {
                get
                {
                    return _view;
                }
                set
                {
                    _view = value;
                }
            }

            public XtraGridViewHandler(GridControl gridControl)
            {
                if (gridControl == null)
                {
                    throw new NullReferenceException("GridControl reference is not set");
                }

                _gridControl = gridControl;
            }

            public abstract void Initialize();

            protected abstract void Dispose(bool disposing);

            public void Dispose()
            {
                // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        protected abstract class SaveableGridViewHandler<TView> : XtraGridViewHandler<TView> where TView : GridView
        {
            public SaveableGridViewHandler(GridControl gridControl) : base(gridControl)
            {

            }
        }

        private class FilesFromUserDocsGridViewHandler : XtraGridViewHandler<DocsVision.BackOffice.WinForms.Controls.GridExView>
        {
            private SimpleLogger _logger = new SimpleLogger();
            private Guid _selectedFileId = Guid.Empty;
            public Size fileIconSize;
            public int rowHeight;
            public ImageManager ImageManager;

            public event Action<Guid> SelectedFileIdChanged;

            public Guid SelectedFileId
            {
                get
                {
                    return _selectedFileId;
                }
                set
                {
                    OnFileSelected(value);
                }
            }

            public FilesFromUserDocsGridViewHandler(GridControl gridControl, ImageManager imageManager, int rowHeight, Size fileIconSize) : base(gridControl)
            {
                if (gridControl.MainView == null || !(gridControl.MainView is DocsVision.BackOffice.WinForms.Controls.GridExView))
                {
                    View = new DocsVision.BackOffice.WinForms.Controls.GridExView();
                    gridControl.MainView = View;
                }
                else
                {
                    View = gridControl.MainView as DocsVision.BackOffice.WinForms.Controls.GridExView;
                }

                if (View == null)
                {
                    throw new NullReferenceException("View reference is not set");
                }

                if (imageManager == null)
                {
                    throw new NullReferenceException("ImageManager reference is not set");
                }

                this.rowHeight = rowHeight;
                this.fileIconSize = fileIconSize;
                this.fileIconSize.Height = Math.Min(rowHeight, fileIconSize.Height);
                this.ImageManager = imageManager;
            }

            public override void Initialize()
            {
                RestoreView();

                View.RowClick += GridView_RowClick;
                //View.FocusedRowChanged += GridView_FocusedRowChanged;
                View.CustomDrawCell += GridView_CustomDrawCell;

                //var controlGroup = LayoutHelper.FilesX
            }

            private void GridView_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
            {
                //_logger.Trace("GridView_CustomDrawCell begin");

                GridView view = sender as GridView;
                GridColumn fileNameCol = View.Columns[ColumnDefinitionsForTableFiles.DefaultDefs[ColumnDefinitionsForTableFiles.Column.FileName].name];
                Image imageIcon = GetImageForRow(e.RowHandle);
                if (view == null || e.RowHandle < 0 || e.Column.FieldName != fileNameCol.FieldName || imageIcon == null || e.Cache == null || e.Appearance == null || e.Graphics == null)
                {
                    e.DefaultDraw();
                    e.Handled = true;
                    //_logger.Trace("DefaultDraw for row " + e.RowHandle + " col " + e.Column.Caption);
                    //return;
                }
                else
                {
                    int iconMargin = 4;

                    e.Appearance.DrawBackground(e.Graphics, e.Cache, e.Bounds, useZeroOffset: false);

                    //using (Brush backBrush = e.Appearance.GetBackBrush(e.Cache))
                    //{
                    //    if (backBrush == null)
                    //    {
                    //        _logger.Error("Brush backBrush reference is not set");
                    //        e.DefaultDraw();
                    //        e.Handled = true;
                    //        return;
                    //    }
                    //    else
                    //    {
                    //        try
                    //        {
                    //            e.Graphics.FillRectangle(backBrush, e.Bounds);
                    //        }
                    //        catch (Exception ex)
                    //        {
                    //            _logger.Error(ex.ToString());
                    //            e.Appearance.DrawBackground(e.Graphics, e.Cache, e.Bounds, false);
                    //        }
                    //    }
                    //}

                    //_logger.Trace("After background draw");

                    int x = e.Bounds.Left + iconMargin;
                    int y = e.Bounds.Top + (e.Bounds.Height - fileIconSize.Height) / 2;

                    e.Graphics.DrawImage(imageIcon, x, y, fileIconSize.Width, fileIconSize.Height);

                    //_logger.Trace("After image draw");

                    int textX = e.Bounds.Left + iconMargin + fileIconSize.Width + iconMargin;
                    Rectangle textRect = new Rectangle(
                        textX,
                        e.Bounds.Top,
                        e.Bounds.Width - (textX - e.Bounds.Left),
                        e.Bounds.Height
                    );

                    using (Brush foreBrush = e.Appearance.GetForeBrush(e.Cache))
                    {
                        if (foreBrush == null)
                        {
                            //_logger.Error("Brush foreBrush reference is not set");
                            e.DefaultDraw();
                            e.Handled = true;
                            return;
                        }
                        else
                        {
                            StringFormat sf = new StringFormat(StringFormatFlags.NoWrap)
                            {
                                LineAlignment = StringAlignment.Center,
                                Trimming = StringTrimming.EllipsisCharacter
                            };

                            try
                            {
                                //e.Graphics.DrawString(
                                //    e.DisplayText,
                                //    e.Appearance.Font,
                                //    foreBrush,
                                //    textRect,
                                //    sf
                                //);

                                e.Appearance.DrawString(
                                    e.Cache,
                                    e.DisplayText,
                                    textRect
                                );
                            }
                            catch (Exception ex)
                            {
                                _logger.Error("Error on Graphics.DrawString" + Environment.NewLine + ex.ToString());
                                _logger.Debug("Text rect. Top=" + textRect.Top + " Bottom=" + textRect.Bottom + " Left=" + textRect.Left + " Right=" + textRect.Right);
                                _logger.Debug(string.Format("DisplayText:{0} | Appearance.Font:{1} | foreBrush:{2} | e.Graphics:{3}", e.DisplayText, e.Appearance.Font, foreBrush == null, e.Graphics == null));
                            }
                        }
                    }

                    e.Handled = true;

                    //_logger.Trace("CustomDraw for row " + e.RowHandle + " col " + e.Column.Caption);
                }

                if (e.RowHandle == view.DataRowCount - 1)
                {
                    view.BestFitColumns();
                }
            }

            private Image GetImageForRow(int rowHandle)
            {
                //_logger.Trace("GetImageForRow begin");

                GridColumn fileNameCol = View.Columns[ColumnDefinitionsForTableFiles.DefaultDefs[ColumnDefinitionsForTableFiles.Column.FileName].name];
                string displayText = View.GetRowCellDisplayText(rowHandle, fileNameCol);

                int startIdx = displayText.LastIndexOf('.') + 1;
                int len = displayText.Length - startIdx;

                string extension = displayText.Substring(startIdx, len).ToLower();
                //_logger.Trace("Format: " + format);

                ImageAlias imageAlias = ImageManager.GetAliasFromBindedFileExtensions(extension);
                Image image = ImageManager.Get<Image>(imageAlias);
                if (image == null)
                {
                    image = ImageManager.SvgImages.GetImage(imageAlias.ToString(), null, fileIconSize);
                }

                //_logger.Trace("GetImageForRow end");

                return image;
            }

            protected override void Dispose(bool disposing)
            {
                //View.RowClick -= GridView_RowClick;
                View.FocusedRowChanged -= GridView_FocusedRowChanged;
                View.CustomDrawCell -= GridView_CustomDrawCell;
            }

            public DocumentFileInfo GetSelectedFileInfo()
            {
                DocumentFileInfo info = new DocumentFileInfo();
                info.fileId = _selectedFileId;

                if (_selectedFileId == Guid.Empty)
                {
                    info.name = "";
                }
                else
                {
                    info.name = View.GetRowCellValue(
                        View.FocusedRowHandle,
                        ColumnDefinitionsForTableFiles.DefaultDefs[ColumnDefinitionsForTableFiles.Column.FileName].name)
                        .ToString();
                }

                return info;
            }

            private void OnFileSelected(Guid fileId)
            {
                _logger.Trace("OnFileSelected. prev:" + _selectedFileId + " selected:" + fileId);

                if (_selectedFileId != fileId)
                {
                    OnSelectedFileChanged(fileId);
                }
            }

            private void OnSelectedFileChanged(Guid fileId)
            {
                _selectedFileId = fileId;
                if (SelectedFileIdChanged != null) SelectedFileIdChanged.Invoke(fileId);
            }

            private void GridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
            {
                if (e.RowHandle >= 0)
                {
                    _logger.Trace("GridView_RowClick (row>=0)");

                    var view = View;

                    Guid fileId = Guid.Parse(view.GetRowCellValue(e.RowHandle, ColumnDefinitionsForTableFiles.DefaultDefs[ColumnDefinitionsForTableFiles.Column.FileId].name).ToString());
                    OnFileSelected(fileId);
                }
            }

            private void GridView_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
            {
                if (e.FocusedRowHandle >= 0)
                {
                    _logger.Trace("GridView_FocusedRowChanged (row>=0)");

                    try
                    {
                        var mainView = View;

                        Guid fileId = Guid.Parse(mainView.GetRowCellValue(e.FocusedRowHandle, ColumnDefinitionsForTableFiles.DefaultDefs[ColumnDefinitionsForTableFiles.Column.FileId].name).ToString());
                        OnFileSelected(fileId);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex.ToString());
                    }
                }
                else
                {
                    OnFileSelected(Guid.Empty);
                }
            }

            public void AttachLogger(SimpleLogger logger)
            {
                if (logger != null) _logger = logger;
            }

            public void CustomizeGridView()
            {
                _logger.Trace("CustomizeFilesTable() begin");

                var view = View;
                if (view.Columns == null || view.Columns.Count == 0)
                {
                    _logger.Warn("Columns collection is not set");
                    return;
                }

                try
                {
                    view.ClearSorting();
                    view.ClearGrouping();

                    view.SortInfo.Clear();
                    view.OptionsCustomization.AllowSort = true;

                    view.RowHeight = rowHeight;

                    var colDefs = ColumnDefinitionsForTableFiles.DefaultDefs;
                    foreach (var colDef in colDefs)
                    {
                        DevExpress.XtraGrid.Columns.GridColumn col = view.Columns[colDef.Value.name];
                        if (col == null)
                        {
                            _logger.Warn("Column with name '" + colDef.Value.name + "' not found");
                            continue;
                        }

                        col.OptionsColumn.AllowEdit = false;
                        col.OptionsColumn.AllowFocus = false;
                        col.OptionsColumn.ReadOnly = false;

                        col.Caption = colDef.Value.displayName;

                        if (colDef.Key == ColumnDefinitionsForTableFiles.Column.FileName)
                        {
                            col.OptionsColumn.AllowShowHide = false;
                            col.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.False;

                            col.VisibleIndex = 0;
                            col.SortIndex = 1;
                            col.SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
                        }

                        if (colDef.Key == ColumnDefinitionsForTableFiles.Column.CheckinDate)
                        {
                            col.VisibleIndex = 1;
                            col.SortIndex = 2;
                            col.SortOrder = DevExpress.Data.ColumnSortOrder.Descending;
                        }

                        if (colDef.Key == ColumnDefinitionsForTableFiles.Column.FileType)
                        {
                            col.VisibleIndex = 2;
                            col.SortIndex = 0;
                            col.SortOrder = DevExpress.Data.ColumnSortOrder.Descending;
                        }

                        if (colDef.Key == ColumnDefinitionsForTableFiles.Column.FileId)
                        {
                            col.Visible = false;
                        }

                        if (colDef.Key == ColumnDefinitionsForTableFiles.Column.CurrentVersionId)
                        {
                            col.Visible = false;
                        }

                        if (colDef.Key == ColumnDefinitionsForTableFiles.Column.Version)
                        {
                            col.Visible = false;
                        }

                        if (colDef.Key == ColumnDefinitionsForTableFiles.Column.VersionNumber)
                        {
                            col.Visible = false;
                        }
                    }

                    view.BestFitColumns();

                    view.OptionsSelection.EnableAppearanceFocusedCell = true;
                    view.OptionsSelection.MultiSelect = false;
                    view.OptionsSelection.EnableAppearanceHideSelection = false;

                    // -- Настройка группировки
                    view.OptionsView.ShowGroupPanel = false;
                    view.OptionsCustomization.AllowGroup = false;
                    view.OptionsBehavior.AutoExpandAllGroups = false;

                    //DevExpress.XtraGrid.Columns.GridColumn fileTypeCol = view.Columns[colDefs[ColumnDefinitionsForTableFiles.Column.FileType].name];
                    //fileTypeCol.GroupIndex = 0;

                    view.OptionsBehavior.AutoExpandAllGroups = false;
                    // --

                    //DevExpress.XtraGrid.Columns.GridColumn fileNameCol = view.Columns[colDefs[ColumnDefinitionsForTableFiles.Column.FileName].name];

                    // -- Настройка сортировки
                    //view.SortInfo.Clear();
                    //view.OptionsCustomization.AllowSort = true;
                    // Сортировка по типу файла
                    //view.SortInfo.Add(new DevExpress.XtraGrid.Columns.GridColumnSortInfo(fileTypeCol, DevExpress.Data.ColumnSortOrder.Descending));
                    // Сортировка по названию файла
                    //view.SortInfo.Add(new DevExpress.XtraGrid.Columns.GridColumnSortInfo(fileNameCol, DevExpress.Data.ColumnSortOrder.Ascending));
                    // --

                    view.ExpandAllGroups();

                    //_logger.Trace("Col fileType name: " + fileTypeCol.FieldName + " SortOrder: " + fileTypeCol.SortOrder + " GroupIndex: " + fileTypeCol.GroupIndex);
                    //_logger.Trace("Col fileType name: " + fileNameCol.FieldName + " SortOrder: " + fileNameCol.SortOrder + " GroupIndex: " + fileNameCol.GroupIndex);

                    // -- Настройки панели поиска
                    view.OptionsFind.Reset();

                    view.OptionsFind.AllowFindPanel = false;
                    //bool rowsThresholdActive = view.RowCount > _tableFilesSearchPanelThreshold;
                    //view.OptionsFind.AlwaysVisible = rowsThresholdActive;
                    //view.OptionsFind.ShowFindButton = rowsThresholdActive;
                    //view.OptionsFind.ShowClearButton = rowsThresholdActive;
                    // --

                    view.OptionsView.ColumnAutoWidth = true;

                }
                catch (Exception ex)
                {
                    _logger.Error(ex.ToString());
                }

                _logger.Trace("CustomizeFilesTable() end");
            }

            public void RestoreView()
            {
                _logger.Trace("RestoreView begin");

                if (View == null)
                {
                    _logger.Error("View reference is not set");
                }
                else
                {
                    View.PopulateColumns();
                    CustomizeGridView();
                }

                _logger.Trace("RestoreView end");
            }
        }

        private class FilesFromUserDocsGridWrapper : XtraGridWrapper
        {
            private SimpleLogger _logger = new SimpleLogger();
            private FilesFromUserDocsGridViewHandler _viewHandler;
            private SqlServerExtensionResolver _sqlServerExtResolver;

            public FilesFromUserDocsGridViewHandler ViewHandler { get { return _viewHandler; } }
            public Guid SelectedFileId
            {
                get
                {
                    return _viewHandler.SelectedFileId;
                }
                set
                {
                    _viewHandler.SelectedFileId = value;
                }
            }

            public FilesFromUserDocsGridWrapper(GridControl gridControl, FilesFromUserDocsGridViewHandler viewHandler, SqlServerExtensionResolver sqlServerExtResolver) : base(gridControl)
            {
                if (gridControl == null)
                {
                    throw new NullReferenceException("GridControl reference is not set");
                }

                if (viewHandler == null)
                {
                    throw new NullReferenceException("ViewHandler reference is not set");
                }

                if (sqlServerExtResolver == null)
                {
                    throw new NullReferenceException("SqlServerExtensionResolver reference is not set");
                }

                _viewHandler = viewHandler;
                _sqlServerExtResolver = sqlServerExtResolver;
            }

            public override void Initilize()
            {
                if (!_initialized)
                {
                    GridControl gridControl = Control;

                    gridControl.DataSource = CreateEmptyTable();

                    gridControl.DataSourceChanged += GridControl_DataSourceChanged;

                    _viewHandler.Initialize();
                }
                else
                {
                    _logger.Warn("Already initialized");
                }
            }

            protected override void Dispose(bool disposing)
            {
                if (!_disposed)
                {
                    if (disposing)
                    {
                        if (_viewHandler != null)
                        {
                            _viewHandler.Dispose();
                            _viewHandler = null;
                        }

                        GridControl gridControl = Control;
                        gridControl.DataSourceChanged -= GridControl_DataSourceChanged;
                    }

                    _disposed = true;
                }
            }

            private static System.Data.DataTable CreateEmptyTable()
            {
                System.Data.DataTable dt = new System.Data.DataTable("FilesFromUserDocument");
                var colDefs = ColumnDefinitionsForTableFiles.DefaultDefs;
                foreach (var colDef in colDefs)
                {
                    dt.Columns.Add(new DataColumn(colDef.Value.name, colDef.Value.type));
                }

                return dt;
            }

            private System.Data.DataTable ConvertXmlToDataTable(string xml)
            {
                _logger.Trace("ConvertXmlToDataTable begin");

                System.Data.DataTable table = CreateEmptyTable();

                if (string.IsNullOrEmpty(xml))
                {
                    _logger.Trace("ConvertXmlToDataTable end because xml is null or empty");
                    return table;
                }

                using (var reader = XmlReader.Create(new System.IO.StringReader(xml)))
                {
                    table.BeginLoadData();

                    while (reader.Read())
                    {
                        // Ищем элемент <R>
                        if (reader.NodeType == XmlNodeType.Element && reader.Name == "R")
                        {
                            DataRow row = table.NewRow();

                            // Заполняем строку, проверяя наличие каждого атрибута
                            if (reader.HasAttributes)
                            {
                                for (int i = 0; i < reader.AttributeCount; i++)
                                {
                                    reader.MoveToAttribute(i);
                                    string attrName = reader.Name;
                                    string attrValue = reader.Value;

                                    SetValueForDataTableRow(row, attrName, attrValue, false);
                                }
                                reader.MoveToElement(); // Возвращаемся к элементу <R>
                            }

                            table.Rows.Add(row);
                        }
                    }

                    table.EndLoadData();
                }

                table.AcceptChanges();

                _logger.Trace("ConvertXmlToDataTable end");

                return table;
            }

            private static void SetValueForDataTableRow(DataRow row, string columnName, string value, bool checkColumnName)
            {
                if (checkColumnName && !row.Table.Columns.Contains(columnName)) return;

                switch (columnName)
                {
                    case "FileType":
                        if (value == "0") row[columnName] = "Основной";
                        if (value == "1") row[columnName] = "Дополнительный";
                        break;
                    default:
                        row[columnName] = value ?? "";
                        break;
                }
            }

            public void LoadTableData(out int rowCount, Guid documentId)
            {
                _logger.Trace("LoadTableData begin. DocumentId:" + documentId.ToString());

                rowCount = 0;

                try
                {
                    GridControl gridControl = Control;

                    if (documentId != Guid.Empty)
                    {
                        //string xml = _sqlServerExtResolver.GetFiles(documentId);
                        //_logger.Debug(xml);

                        System.Data.DataTable dt = ConvertXmlToDataTable(_sqlServerExtResolver.GetFiles(documentId));

                        rowCount = dt.Rows.Count;

                        if (dt.Rows.Count == 0)
                        {
                            gridControl.DataSource = null;
                        }
                        else
                        {
                            _logger.Trace("Table files should be with some data");

                            gridControl.DataSource = dt;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex.ToString());
                }

                _logger.Trace("LoadTableData end");
            }

            public void AttachLogger(SimpleLogger logger)
            {
                if (logger != null) _logger = logger;
            }

            private void GridControl_DataSourceChanged(object sender, EventArgs e)
            {
                if (Control.DataSource == null)
                {
                    //Control.DataSource = CreateEmptyTable();
                }

                _viewHandler.RestoreView();
            }
        }

        private class UserDocsGridViewHandler : XtraGridViewHandler<DocsVision.BackOffice.WinForms.Controls.GridExView>
        {
            private SimpleLogger _logger = new SimpleLogger();
            private CardUrlHelper _cardUrlHelper;
            private GridViewMenuItemRepository _gridMenuItemsRepository;
            private Guid _selectedDocumentId = Guid.Empty;
            private FilesFromUserDocsGridViewHandler _filesGridViewHandler;

            public ImageManager ImageManager;
            public LayoutItemHelper layoutItemHelper;

            public event Action<Guid> SelectedDocumentIdChanged;
            public Guid SelectedDocumentId
            {
                get
                {
                    return _selectedDocumentId;
                }
                set
                {
                    OnDocumentSelected(value);
                }
            }

            private void OnSelectedDocumentChanged(Guid documentId)
            {
                _selectedDocumentId = documentId;
                if (SelectedDocumentIdChanged != null) SelectedDocumentIdChanged(documentId);
            }

            private void OnDocumentSelected(Guid documentId)
            {
                _filesGridViewHandler.SelectedFileId = Guid.Empty;

                if (_selectedDocumentId != documentId)
                {
                    OnSelectedDocumentChanged(documentId);
                }
            }

            public UserDocsGridViewHandler(
                GridControl gridControl,
                CardUrlHelper cardUrlHelper,
                GridViewMenuItemRepository menuItemRepository,
                FilesFromUserDocsGridViewHandler filesFromUserDocsGridViewHandler,
                ImageManager imageManager,
                LayoutItemHelper layoutItemHelper) : base(gridControl)
            {
                if (gridControl.MainView == null || !(gridControl.MainView is DocsVision.BackOffice.WinForms.Controls.GridExView))
                {
                    View = new DocsVision.BackOffice.WinForms.Controls.GridExView();
                    gridControl.MainView = View;
                }
                else
                {
                    View = gridControl.MainView as DocsVision.BackOffice.WinForms.Controls.GridExView;
                }

                if (View == null)
                {
                    throw new NullReferenceException("View reference is not set");
                }

                if (cardUrlHelper == null)
                {
                    throw new NullReferenceException("CardUrlHelper reference is not set");
                }

                if (filesFromUserDocsGridViewHandler == null)
                {
                    throw new NullReferenceException("FilesFromUserDocsGridViewHandler reference is not set");
                }

                if (imageManager == null)
                {
                    throw new NullReferenceException("ImageManager reference is not set");
                }

                if (layoutItemHelper == null)
                {
                    throw new NullReferenceException("LayoutItemHelper reference is not set");
                }

                _cardUrlHelper = cardUrlHelper;
                _gridMenuItemsRepository = menuItemRepository;
                _filesGridViewHandler = filesFromUserDocsGridViewHandler;
                this.ImageManager = imageManager;
                this.layoutItemHelper = layoutItemHelper;
            }

            public void AttachLogger(SimpleLogger logger)
            {
                if (logger != null) _logger = logger;
            }

            public override void Initialize()
            {
                _logger.Trace("Initialize begin");

                var view = View;
                if (view == null)
                {
                    _logger.Error("View reference is not set");
                    throw new NullReferenceException("View reference is not set");
                }

                view.RowCellClick += GridView_RowCellClick;
                view.FocusedRowChanged += GridView_FocusedRowChanged;
                //view.RowClick += TableDocsView_RowClick;
                view.StartGrouping += GridView_StartGrouping;
                view.EndGrouping += GridView_EndGrouping;
                view.CustomDrawFooter += GridView_CustomDrawFooter;
                view.PopupMenuShowing += GridView_PopupMenuShowing;
                view.GroupRowExpanded += GridView_GroupRowExpanded;
                view.GroupRowCollapsed += GridView_GroupRowCollapsed;
                view.ColumnFilterChanged += GridView_ColumnFilterChanged;

                RestoreView();
                //CustomizeGridView();

                //UpdateUpperToolbarCotrols();

                _logger.Trace("Initialize end");
            }

            private void GridView_GroupRowCollapsed(object sender, DevExpress.XtraGrid.Views.Base.RowEventArgs e)
            {
                UpdateUpperToolbarCotrols();
            }

            private void GridView_GroupRowExpanded(object sender, DevExpress.XtraGrid.Views.Base.RowEventArgs e)
            {
                UpdateUpperToolbarCotrols();
            }

            protected override void Dispose(bool disposing)
            {
                _logger.Debug("Dispose begin. disposing=" + disposing + "; disposed=" + _disposed);

                if (!_disposed)
                {
                    if (disposing)
                    {
                        // TODO: dispose managed state (managed objects)
                        if (_gridMenuItemsRepository != null) _gridMenuItemsRepository.Dispose();

                        var view = View;
                        if (view != null)
                        {
                            view.RowCellClick -= GridView_RowCellClick;
                            view.FocusedRowChanged -= GridView_FocusedRowChanged;
                            //view.RowClick -= TableDocsView_RowClick;
                            view.StartGrouping -= GridView_StartGrouping;
                            view.EndGrouping -= GridView_EndGrouping;
                            view.CustomDrawFooter -= GridView_CustomDrawFooter;
                            view.PopupMenuShowing -= GridView_PopupMenuShowing;
                            view.GroupRowExpanded -= GridView_GroupRowExpanded;
                            view.GroupRowCollapsed -= GridView_GroupRowCollapsed;
                            view.ColumnFilterChanged -= GridView_ColumnFilterChanged;
                        }
                        else
                        {
                            _logger.Error("View reference is not set");
                        }
                    }

                    // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                    // TODO: set large fields to null
                    _disposed = true;
                }

                _logger.Debug("Dispose end");
            }

            private void GridView_ColumnFilterChanged(object sender, EventArgs e)
            {
                BaseLayoutItem clearFiltersLayoutItem = layoutItemHelper.GetItem<BaseLayoutItem>(LayoutItemAlias.ClearFiltersButton);

                if (!string.IsNullOrEmpty(View.ActiveFilterString))
                {
                    clearFiltersLayoutItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                }
                else
                {
                    clearFiltersLayoutItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;
                }
            }

            public void RestoreView()
            {
                _logger.Trace("RestoreView begin");

                View.PopulateColumns();
                CustomizeView();
                UpdateUpperToolbarCotrols();

                _logger.Trace("RestoreView end");
            }

            public void CustomizeView()
            {
                _logger.Trace("CustomizeGridView() begin");

                try
                {
                    View.GroupPanelText = "Перенесите сюда колонку для группировки";

                    View.OptionsMenu.EnableColumnMenu = true;

                    View.OptionsView.ShowGroupPanel = false;
                    View.OptionsView.ShowAutoFilterRow = false;
                    View.OptionsView.ShowFilterPanelMode = DevExpress.XtraGrid.Views.Base.ShowFilterPanelMode.Never;

                    View.OptionsSelection.EnableAppearanceFocusedCell = true;
                    View.OptionsSelection.MultiSelect = false;
                    View.OptionsSelection.EnableAppearanceHideSelection = false;

                    View.OptionsView.ShowFooter = true;
                    View.OptionsMenu.EnableFooterMenu = false;

                    //View.OptionsBehavior.AutoExpandAllGroups = true;
                    View.OptionsBehavior.KeepGroupExpandedOnSorting = true;

                    View.SortInfo.Clear();
                    View.ClearGrouping();

                    CustomizeUpperToolbar();

                    var colDefs = ColumnDefinitionsForTableDocs.DefaultDefs;
                    foreach (var colDef in colDefs)
                    {
                        DevExpress.XtraGrid.Columns.GridColumn col = View.Columns[colDef.Value.name];
                        if (col == null)
                        {
                            _logger.Warn("Column with name '" + colDef.Value.name + "' not found");
                            continue;
                        }

                        col.OptionsColumn.AllowEdit = false;
                        col.OptionsColumn.AllowFocus = true;
                        col.OptionsColumn.ReadOnly = false;

                        col.Caption = colDef.Value.displayName;

                        if (colDef.Key == ColumnDefinitionsForTableDocs.Column.KindName)
                        {
                            //col.GroupIndex = groupIndex;
                            //groupIndex++;
                        }

                        if (colDef.Key == ColumnDefinitionsForTableDocs.Column.Name)
                        {
                            col.AppearanceCell.ForeColor = System.Drawing.Color.Blue;
                            col.AppearanceCell.Options.UseForeColor = true;
                            col.AppearanceCell.Font = new System.Drawing.Font(col.AppearanceCell.Font, System.Drawing.FontStyle.Underline);

                            col.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.False;
                            col.OptionsColumn.AllowShowHide = false;
                        }

                        if (colDef.Key == ColumnDefinitionsForTableDocs.Column.CreationDate)
                        {
                            col.BestFit();
                            //View.SortInfo.Insert(0, col, DevExpress.Data.ColumnSortOrder.Descending);
                            col.SortIndex = 0;
                            col.SortOrder = DevExpress.Data.ColumnSortOrder.Descending;
                        }

                        if (colDef.Key == ColumnDefinitionsForTableDocs.Column.RegDate)
                        {
                            //View.SortInfo.Insert(1, col, DevExpress.Data.ColumnSortOrder.Descending);
                            col.BestFit();
                        }

                        if (colDef.Key == ColumnDefinitionsForTableDocs.Column.StateName)
                        {
                            col.BestFit();
                        }

                        if (colDef.Key == ColumnDefinitionsForTableDocs.Column.CardId)
                        {
                            col.Visible = false;
                            col.OptionsColumn.AllowGroup = DevExpress.Utils.DefaultBoolean.False;
                        }
                    }

                    // !
                    View.OptionsView.ColumnAutoWidth = true;
                }
                catch (Exception ex)
                {
                    _logger.Error(ex.ToString());
                }

                _logger.Trace("CustomizeGridView() end");
            }

            private void GridView_PopupMenuShowing(object sender, DevExpress.XtraGrid.Views.Grid.PopupMenuShowingEventArgs e)
            {
                _logger.Trace("GridView_PopupMenuShowing begin");

                var view = sender as GridExView;
                if (view == null)
                {
                    _logger.Warn("view reference is not set or not a GridExView. Sender type=" + sender.GetType());
                    return;
                }

                if (!ReferenceEquals(view, View))
                {
                    _logger.Warn("sender View reference not equals to handler's view");
                }

                var hitInfo = e.HitInfo;
                if (hitInfo == null)
                {
                    _logger.Error("e.HitInfo is null");
                    return;
                }

                _logger.Debug("GridMenuType=" + e.MenuType + " HitInDataRow=" + hitInfo.InDataRow + " HitInColumnPanel=" + hitInfo.InColumnPanel);

                if (e.Menu != null)
                {
                    e.Menu.CloseUp += GridViewDataRowMenu_CloseUp;

                    var menuItems = e.Menu.Items;
                    //int n = menuItems.Count;

                    if (hitInfo.InDataRow && e.MenuType == GridMenuType.Row)
                    {
                        _logger.Trace("InDataRow");

                        DXMenuItem openCardItem = _gridMenuItemsRepository.AddOrGetItem(ContextGridViewMenuItemHelper.CustomMenuItems.OpenCardItemCaption, (s, args) =>
                        {
                            OpenCard(e.HitInfo.RowHandle);
                        });
                        DXMenuItem copyUrlToClipboardItem = _gridMenuItemsRepository.AddOrGetItem(ContextGridViewMenuItemHelper.CustomMenuItems.CopyUrlItemCaption, (s, args) =>
                        {
                            CopyDocumentURLToClipboard(e.HitInfo.RowHandle);
                            //Clipboard.SetText(_cardUrlHelper.GetUrl(view.GetRowCellValue(e.HitInfo.RowHandle, ColumnDefinitionsForTableDocs.DefaultDefs[ColumnDefinitionsForTableDocs.Column.CardId].name).ToString()));
                        });

                        menuItems.Add(openCardItem);
                        menuItems.Add(copyUrlToClipboardItem);
                    }

                    if (hitInfo.InColumnPanel && e.MenuType == GridMenuType.Column)
                    {
                        ContextGridViewMenuItemHelper.HideDocsGridViewColumnMenuStandardItems(e.Menu);
                        ContextGridViewMenuItemHelper.RenameMenuItem(e.Menu, ContextGridViewMenuItemHelper.BuiltInMenuItems.ClearSortingCaption, ContextGridViewMenuItemHelper.CustomMenuItems.ClearSortingCaption);

                        //if (!string.IsNullOrEmpty(view.ActiveFilterString))
                        //{
                        //    TraceAndSaveItemsImagesFromGridViewMenu(e.Menu, _logger);
                        //}
                    }

                    if (hitInfo.InGroupPanel)
                    {
                        //ContextGridViewMenuItemHelper.HideDocsGridViewGroupPanelMenuStandardItems(e.Menu);
                        e.Allow = false;
                    }

                    // !!!
                    //TraceAndSaveItemsImagesFromGridViewMenu(e.Menu, _logger);
                }
                else
                {
                    _logger.Warn("DevExpress.XtraGrid.Views.Grid.PopupMenuShowingEventArgs e.Menu is null");
                    //e.Allow = false;
                }

                _logger.Trace("GridView_PopupMenuShowing end");
            }

            private void GridViewDataRowMenu_CloseUp(object sender, EventArgs e)
            {
                _logger.Trace("GridViewDataRowMenu_CloseUp begin");

                var menu = sender as GridViewMenu;
                if (menu == null)
                {
                    _logger.Warn("menu reference is not set or not a GridViewMenu. Sender type=" + sender.GetType());
                    return;
                }

                menu.CloseUp -= GridViewDataRowMenu_CloseUp;

                _logger.Trace("GridViewDataRowMenu_CloseUp end");
            }

            public string GetFooterText(int filteredRowCount)
            {
                return "Документов (всего): " + filteredRowCount;
            }

            public string GetFooterText(int filteredRowCount, int totalRowCount)
            {
                return "Документов (отфильтрованных): " + filteredRowCount + " | Всего: " + totalRowCount;
            }

            private void GridView_CustomDrawFooter(object sender, DevExpress.XtraGrid.Views.Base.RowObjectCustomDrawEventArgs e)
            {
                var view = sender as GridView;
                if (view == null) return;

                // variant 1 - via e.Graphics (not recommended, can lead to Exception)
                //e.Appearance.FillRectangle(e.Cache, e.Bounds);

                //e.Appearance.TextOptions.HAlignment = HorzAlignment.Center;
                //e.Appearance.Font = new System.Drawing.Font("Tahoma", 9f, FontStyle.Bold);
                //e.Appearance.ForeColor = Color.Black;

                //string footerText = "Документов: " + view.DataRowCount;

                //SizeF textSize = e.Graphics.MeasureString(footerText, e.Appearance.Font);
                //float x = e.Bounds.Left + 3;
                //float y = e.Bounds.Top + (e.Bounds.Height - textSize.Height) / 2;

                //e.Graphics.DrawString
                //(
                //    footerText,
                //    e.Appearance.Font,
                //    e.Appearance.GetForeBrush(e.Cache),
                //    x,
                //    y
                //);

                // variant 2 - via e.Appearance
                view.FooterPanelHeight = 20;

                int markWidth = 16;
                int offset = -5;
                e.DefaultDraw();
                string text;

                if (string.IsNullOrEmpty(view.ActiveFilterString))
                {
                    text = GetFooterText(view.DataRowCount);
                }
                else
                {
                    if (view.GridControl == null)
                    {
                        throw new NullReferenceException("View.GridControl referenec is not set");
                    }

                    var dataSource = view.GridControl.DataSource;
                    if (dataSource == null)
                    {
                        text = GetFooterText(view.DataRowCount);
                    }
                    else if (view.GridControl.DataSource is System.Data.DataTable)
                    {
                        System.Data.DataTable table = view.GridControl.DataSource as System.Data.DataTable;
                        if (table != null && table.Rows != null)
                        {
                            text = GetFooterText(view.DataRowCount, table.Rows.Count);
                        }
                        else
                        {
                            text = GetFooterText(view.DataRowCount);
                        }
                    }
                    else
                    {
                        text = GetFooterText(view.DataRowCount);
                    }
                }

                if (view.GridControl.DataSource is System.Data.DataTable)
                {
                    System.Data.DataTable table = view.GridControl.DataSource as System.Data.DataTable;
                }

                Rectangle markRectangle = new Rectangle(e.Bounds.X + offset, e.Bounds.Y + offset + (markWidth + offset) * 1, markWidth, markWidth);
                Rectangle textRect = new Rectangle(markRectangle.Right + offset, markRectangle.Y, e.Bounds.Width, markRectangle.Height);
                e.Appearance.TextOptions.HAlignment = HorzAlignment.Near;
                e.Appearance.Options.UseTextOptions = true;
                Font font = new Font("Tahoma", 9f, FontStyle.Bold);
                StringFormat sf = new StringFormat()
                {
                    Alignment = StringAlignment.Near
                };
                //e.Appearance.DrawString(e.Cache, text + view.DataRowCount.ToString(), textRect);
                e.Appearance.DrawString(e.Cache, text, textRect, font, sf);

                e.Handled = true;
            }

            private void GridView_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
            {
                DocsVision.BackOffice.WinForms.Controls.GridExView view;
                if (sender is DocsVision.BackOffice.WinForms.Controls.GridExView)
                {
                    view = sender as DocsVision.BackOffice.WinForms.Controls.GridExView;
                    if (e.FocusedRowHandle >= 0)
                    {
                        _logger.Trace("GridView_FocusedRowChanged (row>=0)");

                        try
                        {
                            string cardIdColName = ColumnDefinitionsForTableDocs.DefaultDefs[ColumnDefinitionsForTableDocs.Column.CardId].name;
                            if (view.Columns[cardIdColName] == null)
                            {
                                _logger.Error("Column '" + cardIdColName + "' not found in TableResults view");
                                return;
                            }

                            Guid docId = Guid.Parse(view.GetRowCellValue(e.FocusedRowHandle, cardIdColName).ToString());

                            _logger.Trace("Selected document id:" + docId.ToString());

                            //view.MakeRowVisible(e.FocusedRowHandle);
                            OnDocumentSelected(docId);
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(ex.ToString());
                        }
                    }
                    else
                    {
                        OnDocumentSelected(Guid.Empty);
                    }
                }
                else
                {
                    _logger.Error("sender is not a DocsVision.BackOffice.WinForms.Controls.GridExView");
                }
            }

            private void GridView_RowCellClick(object sender, DevExpress.XtraGrid.Views.Grid.RowCellClickEventArgs e)
            {
                if (e.Button == System.Windows.Forms.MouseButtons.Left &&
                    e.Column.FieldName == ColumnDefinitionsForTableDocs.DefaultDefs[ColumnDefinitionsForTableDocs.Column.Name].name &&
                    e.RowHandle >= 0)
                {
                    OpenCard(e.RowHandle);
                }
            }

            private void GridView_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
            {
                DocsVision.BackOffice.WinForms.Controls.GridExView view;
                if (sender is DocsVision.BackOffice.WinForms.Controls.GridExView)
                {
                    view = sender as DocsVision.BackOffice.WinForms.Controls.GridExView;

                    if (e.RowHandle >= 0)
                    {
                        _logger.Trace("GridView_RowClick (row>=0)");

                        try
                        {
                            DevExpress.XtraGrid.Columns.GridColumn cardIdCol = view.Columns[ColumnDefinitionsForTableDocs.DefaultDefs[ColumnDefinitionsForTableDocs.Column.CardId].name];
                            if (cardIdCol == null)
                            {
                                _logger.Error("Column '" + ColumnDefinitionsForTableDocs.DefaultDefs[ColumnDefinitionsForTableDocs.Column.CardId].name + "' not found in TableWithDocs view");
                                return;
                            }
                            else
                            {
                                Guid docId = Guid.Parse(view.GetRowCellValue(e.RowHandle, cardIdCol.FieldName).ToString());
                                _logger.Trace("Selected document id:" + docId.ToString());
                                OnDocumentSelected(docId);
                                //SelectedCardId = Guid.Parse(view.GetRowCellValue(e.RowHandle, cardIdCol.FieldName).ToString());

                                //_logger.Trace("Selected card id:" + SelectedCardId.ToString());
                            }

                            view.GridControl.Focus();
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(ex.ToString());
                        }
                    }
                    else
                    {
                        //_selectedFileId = Guid.Empty;
                    }
                }
                else
                {
                    _logger.Error("sender is not a DocsVision.BackOffice.WinForms.Controls.GridExView");
                }
            }

            private void GridView_StartGrouping(object sender, EventArgs e)
            {
                _logger.Trace("GridView_StartGrouping begin");

                if (sender is GridView)
                {
                    var view = sender as GridView;
                    if (view != null)
                    {

                    }
                }

                _logger.Trace("GridView_StartGrouping end");
            }

            private void GridView_EndGrouping(object sender, EventArgs e)
            {
                _logger.Trace("GridView_EndGrouping begin");

                UpdateUpperToolbarCotrols();

                _logger.Trace("GridView_EndGrouping end");
            }

            private void OpenCard(int rowHandle)
            {
                _logger.Trace("OpenCard rowHandle=" + rowHandle);

                if (rowHandle < 0)
                {
                    _logger.Warn("RowHandle less than 0");
                    return;
                }

                var view = View;
                if (view != null)
                {
                    string url = _cardUrlHelper.GetUrl(view.GetRowCellValue(rowHandle, ColumnDefinitionsForTableDocs.DefaultDefs[ColumnDefinitionsForTableDocs.Column.CardId].name).ToString());
                    ReportOnDocumentsScript.OpenCard(url);
                }
                else
                {
                    _logger.Error("View reference is not set");
                }
            }

            private void CopyDocumentURLToClipboard(int rowHandle)
            {
                _logger.Trace("CopyDocumentURLToClipboard rowHandle=" + rowHandle);

                if (rowHandle < 0)
                {
                    _logger.Warn("RowHandle less than 0");
                    return;
                }

                var view = View;
                if (view == null)
                {
                    _logger.Error("view reference is not set");
                    return;
                }

                string url = _cardUrlHelper.GetUrl(view.GetRowCellValue(rowHandle, ColumnDefinitionsForTableDocs.DefaultDefs[ColumnDefinitionsForTableDocs.Column.CardId].name).ToString());
                Clipboard.SetText(url);
            }

            /// <summary>
            /// Обновляет данные элементов на верхней панели управления
            /// </summary>
            public void UpdateUpperToolbarCotrols()
            {
                BaseLayoutItem showGroupPanelLayoutItem = layoutItemHelper.GetItem<BaseLayoutItem>(LayoutItemAlias.GroupPanelVisibilitySwitch);
                DevExpress.XtraEditors.SimpleButton showGroupPanelButton = layoutItemHelper.GetControl<DevExpress.XtraEditors.SimpleButton>(LayoutItemAlias.GroupPanelVisibilitySwitch);

                BaseLayoutItem clearGroupingLayoutItem = layoutItemHelper.GetItem<BaseLayoutItem>(LayoutItemAlias.ClearGroupingButton);
                DevExpress.XtraEditors.SimpleButton clearGroupingButton = layoutItemHelper.GetControl<DevExpress.XtraEditors.SimpleButton>(LayoutItemAlias.ClearGroupingButton);

                BaseLayoutItem collapseExpandGroupsLayoutItem = layoutItemHelper.GetItem<BaseLayoutItem>(LayoutItemAlias.CollapseExpandGroupsButton);
                DevExpress.XtraEditors.SimpleButton collapseExpandGroupsButton = layoutItemHelper.GetControl<DevExpress.XtraEditors.SimpleButton>(LayoutItemAlias.CollapseExpandGroupsButton);

                BaseLayoutItem clearFiltersLayoutItem = layoutItemHelper.GetItem<BaseLayoutItem>(LayoutItemAlias.ClearFiltersButton);
                //DevExpress.XtraEditors.SimpleButton clearFiltersButton = _layoutItemHelper.GetControl<DevExpress.XtraEditors.SimpleButton>(LayoutItemAlias.ClearFiltersButton);

                bool isAllGroupRowsExpanded = Utils.IsAllGroupRowsExpanded(View);
                //bool isAllGroupRowsCollapsed = Utils.IsAllGroupRowsCollapsed(View);

                if (!string.IsNullOrEmpty(View.ActiveFilterString))
                {
                    clearFiltersLayoutItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                }
                else
                {
                    clearFiltersLayoutItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;
                }

                if (View.DataRowCount == 0)
                {
                    showGroupPanelLayoutItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;
                }
                else
                {
                    showGroupPanelLayoutItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                }

                if (View.GroupCount == 0)
                {
                    clearGroupingLayoutItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;

                    collapseExpandGroupsLayoutItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;
                    //collapseExpandGroupsButton.ResetText();
                }
                else
                {
                    clearGroupingLayoutItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                    collapseExpandGroupsLayoutItem.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                }

                if (isAllGroupRowsExpanded)
                {
                    collapseExpandGroupsButton.Text = "Свернуть группы";
                    collapseExpandGroupsButton.ImageOptions.SvgImage = ImageManager.SvgImages[ImageAlias.CollapseAllGroupLevels.ToString()];
                }
                else
                {
                    collapseExpandGroupsButton.Text = "Развернуть группы";
                    collapseExpandGroupsButton.ImageOptions.SvgImage = ImageManager.SvgImages[ImageAlias.ExpandAllGroupLevels.ToString()];
                }

                if (View.OptionsView.ShowGroupPanel)
                {
                    showGroupPanelButton.Text = "Скрыть панель группировки";
                    showGroupPanelButton.ImageOptions.SvgImage = ImageManager.SvgImages[ImageAlias.HideGroupPanel.ToString()];
                }
                else
                {
                    showGroupPanelButton.Text = "Показать панель группировки";
                    showGroupPanelButton.ImageOptions.SvgImage = ImageManager.SvgImages[ImageAlias.ShowGroupPanel.ToString()];
                }
            }

            private void CustomizeUpperToolbar()
            {
                _logger.Trace("CustomizeToolbars begin");

                Size svgSize = new Size(16, 16);

                var controlGroup = layoutItemHelper.DocumentsTableControlPanelGroup;
                controlGroup.TextVisible = false;

                controlGroup.Padding = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);
                controlGroup.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

                ImageAlignToText imageAlignToText = ImageAlignToText.LeftCenter;
                ImageLocation imageLocation = ImageLocation.MiddleLeft;
                HorzAlignment horzAlignment = HorzAlignment.Near;
                //VertAlignment vertAlignment = VertAlignment.Center;

                // Настройка кнопки экпорта в Excel
                SimpleButton exportBtn = layoutItemHelper.ExportToExcelButton;
                exportBtn.ImageOptions.ImageToTextAlignment = imageAlignToText;
                exportBtn.ImageOptions.Location = imageLocation;
                exportBtn.Appearance.TextOptions.HAlignment = horzAlignment;
                Image toXlImg = ImageManager.Get<Image>(ImageAlias.ToExcel);
                if (toXlImg != null)
                {
                    exportBtn.ImageOptions.Image = toXlImg;
                }
                else
                {
                    _logger.Error("ToExcel image reference is not set. ImageAlias: " + ImageAlias.ToExcel);
                }
                //exportBtn.AutoWidthInLayoutControl = false;
                //exportBtn.Width = exportBtn.CalcBestSize().Width;

                // Настройка кнопки выбора видимых столбцов
                DevExpress.XtraEditors.SimpleButton columnsVisibilityButton = layoutItemHelper.GetControl<DevExpress.XtraEditors.SimpleButton>(LayoutItemAlias.ColumnVisibilitySwitch);
                columnsVisibilityButton.ImageOptions.SvgImageSize = columnsVisibilityButton.Size;   // Иконка на весь размер кнопки
                //columnsVisibilityButton.ImageOptions.ImageToTextAlignment = imageAlignToText;
                columnsVisibilityButton.ImageOptions.Location = ImageLocation.MiddleCenter;
                //columnsVisibilityButton.Appearance.TextOptions.HAlignment = horzAlignment;
                //columnsVisibilityButton.MinimumSize = new Size(columnsVisibilityButton.CalcBestSize().Width, columnsVisibilityButton.Height);
                DevExpress.Utils.Svg.SvgImage columnsVisibilitySvg = ImageManager.SvgImages[ImageAlias.ColumnsVisibility.ToString()];
                if (columnsVisibilitySvg != null)
                {
                    columnsVisibilityButton.ImageOptions.SvgImage = columnsVisibilitySvg;
                    columnsVisibilityButton.Padding = new Padding(0, 0, 0, 0);
                }
                else
                {
                    _logger.Error("Columns visibility svg image reference is not set. ImageAlias: " + ImageAlias.ColumnsVisibility);
                }
                //columnsVisibilityButton.AutoWidthInLayoutControl = false;
                //columnsVisibilityButton.Width = columnsVisibilityButton.CalcBestSize().Width;

                // Настройка кнопки с(раз-)ворачивания групп
                DevExpress.XtraEditors.SimpleButton collapseExpandGroupsButton = layoutItemHelper.GetControl<DevExpress.XtraEditors.SimpleButton>(LayoutItemAlias.CollapseExpandGroupsButton);
                collapseExpandGroupsButton.ImageOptions.SvgImageSize = svgSize;
                collapseExpandGroupsButton.ImageOptions.ImageToTextAlignment = imageAlignToText;
                collapseExpandGroupsButton.ImageOptions.Location = imageLocation;
                collapseExpandGroupsButton.Appearance.TextOptions.HAlignment = horzAlignment;
                collapseExpandGroupsButton.Text = "Развернуть группы";
                //collapseExpandGroupsButton.AutoWidthInLayoutControl = false;
                //collapseExpandGroupsButton.Width = collapseExpandGroupsButton.CalcBestSize().Width;

                // Настройка кнопки показа панели группировки
                DevExpress.XtraEditors.SimpleButton groupPanelVisibilityButton = layoutItemHelper.GetControl<DevExpress.XtraEditors.SimpleButton>(LayoutItemAlias.GroupPanelVisibilitySwitch);
                groupPanelVisibilityButton.ImageOptions.SvgImageSize = svgSize;
                groupPanelVisibilityButton.ImageOptions.ImageToTextAlignment = imageAlignToText;
                groupPanelVisibilityButton.ImageOptions.Location = imageLocation;
                groupPanelVisibilityButton.Appearance.TextOptions.HAlignment = horzAlignment;
                groupPanelVisibilityButton.Text = "Показать панель группировки";
                //groupPanelVisibilityButton.AutoWidthInLayoutControl = false;
                //groupPanelVisibilityButton.Width = groupPanelVisibilityButton.CalcBestSize().Width;

                // Настройка кнопки очистки группировки
                DevExpress.XtraEditors.SimpleButton clearGroupingButton = layoutItemHelper.GetControl<DevExpress.XtraEditors.SimpleButton>(LayoutItemAlias.ClearGroupingButton);
                clearGroupingButton.ImageOptions.SvgImageSize = svgSize;
                clearGroupingButton.ImageOptions.ImageToTextAlignment = imageAlignToText;
                clearGroupingButton.ImageOptions.Location = imageLocation;
                clearGroupingButton.Appearance.TextOptions.HAlignment = horzAlignment;
                clearGroupingButton.Text = "Отменить группировку";
                clearGroupingButton.ImageOptions.SvgImage = ImageManager.SvgImages[ImageAlias.ClearGrouping.ToString()];
                //clearGroupingButton.AutoWidthInLayoutControl = false;
                //clearGroupingButton.Width = clearGroupingButton.CalcBestSize().Width;

                // Настрока кнопки очистки фильтров
                DevExpress.XtraEditors.SimpleButton clearFiltersButton = layoutItemHelper.GetControl<DevExpress.XtraEditors.SimpleButton>(LayoutItemAlias.ClearFiltersButton);
                clearFiltersButton.ImageOptions.SvgImageSize = svgSize;
                clearFiltersButton.ImageOptions.ImageToTextAlignment = imageAlignToText;
                clearFiltersButton.ImageOptions.Location = imageLocation;
                clearFiltersButton.Appearance.TextOptions.HAlignment = horzAlignment;
                clearFiltersButton.Text = "Очистить фильтр";
                clearFiltersButton.ImageOptions.SvgImage = ImageManager.SvgImages[ImageAlias.ClearFilters.ToString()];
                //clearFiltersButton.AutoWidthInLayoutControl = false;
                //clearFiltersButton.Width = clearFiltersButton.CalcBestSize().Width;

                _logger.Trace("CustomizeToolbars end");
            }
        }

        private class BackgroundWorkerWrapper
        {
            private BackgroundWorker _worker;
            private List<DoWorkEventHandler> _doWorkHandlers;
            private List<RunWorkerCompletedEventHandler> _workCompletedHandlers;
            private List<ProgressChangedEventHandler> _progressChangedHandlers;
            private List<EventHandler> _disposedHandlers;

            private bool _disposed = false;

            public BackgroundWorker Worker { get { return _worker; } }

            public BackgroundWorkerWrapper(bool workerSupportsCancellation, bool workerReportProgress)
            {
                _worker = new BackgroundWorker();
                _worker.WorkerSupportsCancellation = workerSupportsCancellation;
                _worker.WorkerReportsProgress = workerReportProgress;
            }

            public void AddDoWorkHandler(DoWorkEventHandler handler)
            {
                if (_doWorkHandlers == null)
                {
                    _doWorkHandlers = new List<DoWorkEventHandler>();
                }

                _doWorkHandlers.Add(handler);
                _worker.DoWork += handler;
            }

            public void AddWorkerRunCompletedHandler(RunWorkerCompletedEventHandler handler)
            {
                if (_workCompletedHandlers == null)
                {
                    _workCompletedHandlers = new List<RunWorkerCompletedEventHandler>();
                }

                _workCompletedHandlers.Add(handler);
                _worker.RunWorkerCompleted += handler;
            }

            public void AddProgressChangedHandler(ProgressChangedEventHandler handler)
            {
                if (!_worker.WorkerReportsProgress)
                {
                    throw new Exception("BackgroundWorker doesn't support progress report");
                }

                if (_progressChangedHandlers == null)
                {
                    _progressChangedHandlers = new List<ProgressChangedEventHandler>();
                }

                _progressChangedHandlers.Add(handler);
                _worker.ProgressChanged += handler;
            }

            public void CancelAndWaitForWorkerFinish()
            {
                if (!_worker.IsBusy || !_worker.WorkerSupportsCancellation)
                {
                    return;
                }

                if (!_worker.CancellationPending)
                {
                    _worker.CancelAsync();
                }

                while (_worker.IsBusy && !_worker.CancellationPending)
                {
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(100);
                }
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                CancelAndWaitForWorkerFinish();

                int n, i = 0;

                if (_doWorkHandlers != null)
                {
                    n = _doWorkHandlers.Count;
                    while (i < n)
                    {
                        _worker.DoWork -= _doWorkHandlers[i];
                        i++;
                    }
                    _doWorkHandlers.Clear();
                    _doWorkHandlers = null;
                }

                i = 0;
                if (_workCompletedHandlers != null)
                {
                    n = _workCompletedHandlers.Count;
                    while (i < n)
                    {
                        _worker.RunWorkerCompleted -= _workCompletedHandlers[i];
                        i++;
                    }
                    _workCompletedHandlers.Clear();
                    _workCompletedHandlers = null;
                }

                i = 0;
                if (_progressChangedHandlers != null)
                {
                    n = _progressChangedHandlers.Count;
                    while (i < n)
                    {
                        _worker.ProgressChanged -= _progressChangedHandlers[i];
                        i++;
                    }
                    _progressChangedHandlers.Clear();
                    _progressChangedHandlers = null;
                }

                i = 0;
                if (_disposedHandlers != null)
                {
                    n = _disposedHandlers.Count;
                    while (i < n)
                    {
                        _worker.Disposed -= _disposedHandlers[i];
                        i++;
                    }
                    _disposedHandlers.Clear();
                    _disposedHandlers = null;
                }

                _worker = null;
                _disposed = true;
            }
        }

        protected abstract class XtraGridWrapper : IDisposable
        {
            private GridControl _gridControl;
            protected bool _disposed = false;
            protected bool _initialized = false;

            public GridControl Control
            {
                get
                {
                    if (_gridControl == null)
                    {
                        throw new NullReferenceException("GridControl reference is not set");
                    }

                    return _gridControl;
                }
            }

            public DevExpress.XtraGrid.Views.Base.BaseView MainViewBase { get { return Control.MainView; } }
            public GridExView MainView { get { return Control.MainView as GridExView; } }

            public XtraGridWrapper(GridControl gridControl)
            {
                _gridControl = gridControl;
            }

            public abstract void Initilize();

            protected abstract void Dispose(bool disposing);

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        private struct UnloadToExcelResultInfo
        {
            public int rowCount;
            public Result result;
            public string pathToFile;
            public enum Result { Success, NoData, Fail, Cancelled }

            public UnloadToExcelResultInfo(int rowCount, Result result, string pathToFile)
            {
                this.rowCount = rowCount;
                this.result = result;
                this.pathToFile = pathToFile;
            }
        }

        private class UserDocsGridWrapper : XtraGridWrapper
        {
            private UserDocsGridViewHandler _viewHandler;
            private SimpleLogger _logger = new SimpleLogger();
            private System.Windows.Forms.Label _loadingOverlay;
            private SqlServerExtensionResolver _sqlServerExtResolver;
            private CardUrlHelper _cardUrlHelper;
            private FilesFromUserDocsGridWrapper _filesGridWrapper;

            public string dateFormat = "yyyy-MM-ddTHH:mm:ss";
            public LayoutItemHelper layoutItemHelper;

            public string LastUnloadedXlsDocument { get; protected set; }
            public int LastExportRowCount { get; protected set; }
            public Guid SelectedDocumentId { get { return _viewHandler.SelectedDocumentId; } }
            public UserDocsGridViewHandler ViewHandler { get { return _viewHandler; } }

            public BackgroundWorkerWrapper LoadTableDataWorkerWrapper { get; protected set; }
            public BackgroundWorkerWrapper UnloadToExcelWorkerWrapper { get; protected set; }

            public event Action<System.Data.DataTable> TableLoaded;
            public event Action<UnloadToExcelResultInfo> UnloadedToExcel;

            public void UnloadToExcelAsync(DbSelectDocsOptions options)
            {
                UnloadToExcelWorkerWrapper.Worker.RunWorkerAsync(options);
            }

            public void LoadTableDataAsync(DbSelectDocsOptions options)
            {
                ShowLoadingOverlay(true);
                LoadTableDataWorkerWrapper.Worker.RunWorkerAsync(options);
            }

            public void AttachLogger(SimpleLogger logger)
            {
                if (logger != null)
                {
                    _logger = logger;
                }
            }

            private void OnTableLoaded(System.Data.DataTable table)
            {
                if (TableLoaded != null) TableLoaded.Invoke(table);
            }

            private void OnUnloadToExcel(UnloadToExcelResultInfo resultInfo)
            {
                if (UnloadedToExcel != null) UnloadedToExcel.Invoke(resultInfo);
            }

            public UserDocsGridWrapper(
                GridControl gridControl,
                UserDocsGridViewHandler viewHandler,
                SqlServerExtensionResolver sqlServerExtensionResolver,
                CardUrlHelper cardUrlHelper,
                FilesFromUserDocsGridWrapper filesFromUserDocsGridWrapper,
                LayoutItemHelper layoutItemHelper) : base(gridControl)
            {
                if (gridControl == null)
                {
                    throw new ArgumentNullException("GridControl reference is not set");
                }

                if (viewHandler == null)
                {
                    throw new ArgumentNullException("XtraGridViewHandler reference is not set");
                }

                if (sqlServerExtensionResolver == null)
                {
                    throw new ArgumentNullException("SqlServerExtensionResolver reference is not set");
                }

                if (cardUrlHelper == null)
                {
                    throw new ArgumentNullException("CardUrlHelper reference is not set");
                }

                if (filesFromUserDocsGridWrapper == null)
                {
                    throw new ArgumentNullException("FilesFromUserDocsGridWrapper reference is not set");
                }

                if (layoutItemHelper == null)
                {
                    throw new ArgumentNullException("LayoutItemHelper reference is not set");
                }

                _viewHandler = viewHandler;
                _sqlServerExtResolver = sqlServerExtensionResolver;
                _cardUrlHelper = cardUrlHelper;
                _filesGridWrapper = filesFromUserDocsGridWrapper;
                this.layoutItemHelper = layoutItemHelper;
            }

            public override void Initilize()
            {
                if (!_initialized)
                {
                    _logger.Trace("Initialize begin");

                    LastExportRowCount = 0;

                    DevExpress.XtraGrid.GridControl gridControl = Control;
                    gridControl.DataSource = CreateEmptyTable();

                    gridControl.DataSourceChanged += GridControl_DataSourceChanged;

                    _viewHandler.Initialize();

                    if (LoadTableDataWorkerWrapper == null)
                    {
                        LoadTableDataWorkerWrapper = new BackgroundWorkerWrapper(true, false);
                        LoadTableDataWorkerWrapper.AddDoWorkHandler(LoadTableDataWorker_DoWork);
                        LoadTableDataWorkerWrapper.AddWorkerRunCompletedHandler(LoadTableDataWorker_RunCompleted);
                    }

                    if (UnloadToExcelWorkerWrapper == null)
                    {
                        UnloadToExcelWorkerWrapper = new BackgroundWorkerWrapper(true, false);
                        UnloadToExcelWorkerWrapper.AddDoWorkHandler(UnloadToExcelWorker_DoWork);
                        UnloadToExcelWorkerWrapper.AddWorkerRunCompletedHandler(UnloadToExcelWorker_RunCompleted);
                    }

                    _logger.Trace("Initialize end");

                    _initialized = true;
                }
                else
                {
                    _logger.Warn("Already initialized");
                }
            }

            protected override void Dispose(bool disposing)
            {
                if (!_disposed)
                {
                    if (disposing)
                    {
                        // TODO: dispose managed state (managed objects)

                        LoadTableDataWorkerWrapper.Dispose();
                        UnloadToExcelWorkerWrapper.Dispose();

                        GridControl control = Control;
                        control.DataSourceChanged -= GridControl_DataSourceChanged;

                        if (_loadingOverlay != null)
                        {
                            _loadingOverlay.Visible = false;
                            _loadingOverlay.Dispose();
                        }

                        if (_viewHandler != null) _viewHandler.Dispose();
                        else _logger.Error("ViewHandler reference is not set");
                    }

                    // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                    // TODO: set large fields to null
                    _viewHandler = null;
                    _disposed = true;
                }
            }

            // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
            //~UserDocsGridWrapper()
            //{
            //    // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            //    Dispose(disposing: false);
            //}

            private void UnloadToExcelWorker_RunCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
                LogMemoryUsage("After ReportToExcelWorker_RunCompleted", _logger);

                //_logger.Trace("RunWorkerCompletedEventArgs Cancelled=" + e.Cancelled + " Error is null=" + (e.Error == null));

                UnloadToExcelResultInfo resultInfo = new UnloadToExcelResultInfo();
                resultInfo.rowCount = LastExportRowCount;
                resultInfo.pathToFile = LastUnloadedXlsDocument;

                if (e.Cancelled)
                {
                    resultInfo.result = UnloadToExcelResultInfo.Result.Cancelled;
                }
                else if (e.Error != null || LastExportRowCount < 0)
                {
                    resultInfo.result = UnloadToExcelResultInfo.Result.Fail;
                }
                else if (LastExportRowCount == 0)
                {
                    resultInfo.result = UnloadToExcelResultInfo.Result.NoData;
                }
                else
                {
                    resultInfo.result = UnloadToExcelResultInfo.Result.Success;
                }

                OnUnloadToExcel(resultInfo);
            }

            private void UnloadToExcelWorker_DoWork(object sender, DoWorkEventArgs e)
            {
                DbSelectDocsOptions selectOptions;
                try
                {
                    selectOptions = (DbSelectDocsOptions)e.Argument;
                }
                catch (Exception ex)
                {
                    _logger.Error("Ошибка при приведении аргумента BackgroundWorker к типу " + typeof(DbSelectDocsOptions) + Environment.NewLine + ex.ToString());
                    e.Cancel = true;
                    return;
                }

                //SelectAndLoadToExcel(selectOptions, excelExportBatchSize);
                SelectAndLoadToExcel(selectOptions);
            }

            private void CloseExcelAndReleaseComObjects(Excel.Workbook workbook, Excel.Application excelApp, List<object> objects)
            {
                _logger.Trace("ReleaseExcelObjects begin");

                if (workbook != null) workbook.Close(false);
                if (excelApp != null) excelApp.Quit();

                int n = objects.Count;
                for (int i = 0; i < n; i++)
                {
                    if (objects[i] != null) Marshal.FinalReleaseComObject(objects[i]);
                }

                for (int i = 0; i < n; i++)
                {
                    objects[i] = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                _logger.Trace("ReleaseExcelObjects end");
            }

            /// <summary>
            /// Выгружает данные по документам с фиксированным потреблением памяти.
            /// Создает двумерный массив object с фиксированным объемом.
            /// Запись в файл Excel происходит по частям, число которых зависит от размера батча.
            /// </summary>
            /// <param name="selectOptions"></param>
            /// <param name="batchSize"></param>
            private void SelectAndLoadToExcel(DbSelectDocsOptions selectOptions, int batchSize)
            {
                _logger.Trace("SelectAndLoadToExcel begin. batchSize:" + batchSize);

                LastExportRowCount = 0;
                BackgroundWorker worker = UnloadToExcelWorkerWrapper.Worker;

                long totalMemory = GC.GetTotalMemory(false);
                _logger.Debug("Managed memory on start unloading to excel (B, MB): " + totalMemory / 1024 / 1024);

                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                selectOptions.limit = int.MaxValue;
                string selectResult = _sqlServerExtResolver.GetUserDocuments(selectOptions);
                if (string.IsNullOrEmpty(selectResult))
                {
                    _logger.Trace("Select response string is null or empty");
                    return;
                }

                stopwatch.Stop();
                _logger.Trace("Time elapsed for: 1. SQL command execution; 2. Compressing xml result on server; 3. Decompressing string resut:" + Environment.NewLine +
                    "Total seconds: " + stopwatch.Elapsed.TotalSeconds);

                long totalMemoryDelta = GC.GetTotalMemory(false) - totalMemory;
                _logger.Debug("Managed memory after getting select result (delta, MB): " + totalMemoryDelta / 1024 / 1024);

                if (worker.WorkerSupportsCancellation && worker.CancellationPending)
                {
                    return;
                }

                stopwatch.Restart();

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                Excel.Range creationDateColumn = null, regDateColumn = null;

                List<object> excelObjs = new List<object>();
                excelObjs.Add(creationDateColumn);
                excelObjs.Add(regDateColumn);
                excelObjs.Add(worksheet);
                excelObjs.Add(workbook);
                excelObjs.Add(excelApp);

                excelApp.Visible = false;
                excelApp.ScreenUpdating = false;

                var colDefs = ColumnDefinitionsForTableDocs.DefaultDefs;
                string[] headers = new string[colDefs.Count];
                int colCount = 0;
                foreach (var colDef in colDefs)
                {
                    if (colDef.Key == ColumnDefinitionsForTableDocs.Column.CardId)
                    {
                        headers[colCount] = "Ссылка на карточку";
                    }
                    else
                    {
                        headers[colCount] = colDef.Value.displayName;
                    }

                    colCount++;
                }

                if (worker.WorkerSupportsCancellation && worker.CancellationPending)
                {
                    CloseExcelAndReleaseComObjects(workbook, excelApp, excelObjs);
                    return;
                }

                int startColumnIndex = 1;
                int endColumnIndex = colCount;

                Excel.Range startHeadersCell = (Excel.Range)worksheet.Cells[1, startColumnIndex];
                Excel.Range endHeadersCell = (Excel.Range)worksheet.Cells[1, endColumnIndex];

                Excel.Range headersRange = (Excel.Range)worksheet.Range[startHeadersCell, endHeadersCell];

                excelObjs.Insert(0, headersRange);
                excelObjs.Insert(0, endHeadersCell);
                excelObjs.Insert(0, startHeadersCell);

                // Запись заголовков столбцов
                //Excel.Range headersRange = worksheet.Range["A1", "N1"];
                headersRange.Value = headers;
                headersRange.Font.Bold = true;

                int currentRow = 2;

                _logger.Trace("COL COUNT: " + colCount);

                stopwatch.Stop();
                _logger.Trace("Time elapsed for creating Excel doc in total seconds: " + stopwatch.Elapsed.TotalSeconds);

                totalMemory = GC.GetTotalMemory(false);
                _logger.Debug("Managed memory before xml read (MB): " + totalMemory / 1024 / 1024);

                stopwatch.Restart();

                _logger.Trace("XML read begin");

                List<object[]> rows = new List<object[]>();
                using (XmlReader reader = XmlReader.Create(new System.IO.StringReader(selectResult)))
                {
                    Type tDate = typeof(System.DateTime);

                    string readedValue;

                    while (reader.Read())
                    {
                        if (worker.WorkerSupportsCancellation && worker.CancellationPending)
                        {
                            CloseExcelAndReleaseComObjects(workbook, excelApp, excelObjs);
                            return;
                        }

                        if (reader.NodeType == XmlNodeType.Element && reader.Name == "R")
                        {
                            int i = 0;
                            object[] attrValues = new object[colCount];
                            foreach (var colDefKvp in colDefs)
                            {
                                readedValue = reader.GetAttribute(colDefKvp.Value.name) ?? "";

                                if (readedValue == _minDbDateStr)
                                {
                                    readedValue = "";
                                }

                                if (readedValue != "" && colDefKvp.Value.type == tDate)
                                {
                                    readedValue = readedValue.Replace('T', ' ');
                                }

                                if (colDefKvp.Key == ColumnDefinitionsForTableDocs.Column.CardId)
                                {
                                    attrValues[i] = _cardUrlHelper.GetUrl(readedValue);
                                }
                                else
                                {
                                    attrValues[i] = readedValue;
                                }

                                ++i;
                            }

                            rows.Add(attrValues);
                        }
                    }
                }
                LastExportRowCount = rows.Count;
                _logger.Trace("Rows: " + rows.Count);

                _logger.Trace("XML read end");

                stopwatch.Stop();
                _logger.Trace("Time elapsed for initializing List<object[]> in total seconds: " + stopwatch.Elapsed.TotalSeconds);

                totalMemoryDelta = GC.GetTotalMemory(false) - totalMemory;
                _logger.Debug("Managed memory after xml reading (delta, MB): " + totalMemoryDelta / 1024 / 1024);

                totalMemory = GC.GetTotalMemory(false);
                _logger.Debug("Managed memory before initializing data array (MB): " + totalMemory / 1024 / 1024);

                stopwatch.Restart();

                for (int batchStart = 0; batchStart < rows.Count; batchStart += batchSize)
                {
                    if (worker.WorkerSupportsCancellation && worker.CancellationPending)
                    {
                        CloseExcelAndReleaseComObjects(workbook, excelApp, excelObjs);
                        return;
                    }

                    int batchEnd = Math.Min(batchStart + batchSize, rows.Count);
                    int batchRows = batchEnd - batchStart;

                    object[,] batchData = new object[batchRows, colCount];
                    for (int i = 0; i < batchRows; i++)
                    {
                        Array.Copy(rows[batchStart + i], 0, batchData, i * colCount, colCount);
                    }

                    Excel.Range batchRange = (Excel.Range)worksheet.Range[
                        worksheet.Cells[currentRow, 1],
                        worksheet.Cells[currentRow + batchRows - 1, colCount]
                    ];

                    batchRange.Value = batchData;
                    currentRow += batchRows;

                    _logger.Trace("Batch " + batchStart + "/" + batchSize + ": " + batchRows + " rows written");

                    excelObjs.Add(batchRange);

                    batchData = null;
                    GC.Collect();
                }

                totalMemoryDelta = GC.GetTotalMemory(false) - totalMemory;
                _logger.Debug("Managed memory after writing to all Excel ranges (delta, MB): " + totalMemoryDelta / 1024 / 1024);
                totalMemory = GC.GetTotalMemory(false);

                stopwatch.Restart();

                stopwatch.Stop();
                _logger.Trace("Time elapsed for filling all Excel ranges in total seconds: " + stopwatch.Elapsed.TotalSeconds);

                stopwatch.Restart();

                // Форматирование столбцов с датами
                creationDateColumn = (Excel.Range)worksheet.Columns[4];
                creationDateColumn.NumberFormat = "dd.mm.yyyy hh:mm:ss";
                regDateColumn = (Excel.Range)worksheet.Columns[5];
                regDateColumn.NumberFormat = "dd.mm.yyyy hh:mm:ss";

                //excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                excelApp.ScreenUpdating = true;

                string savePath = GetExcelSavePath();

                _logger.Trace("SavePath to Excel=" + savePath);

                workbook.SaveAs(savePath);

                stopwatch.Stop();
                _logger.Trace("Time elapsed for saving Excel doc in total seconds: " + stopwatch.Elapsed.TotalSeconds);

                LastUnloadedXlsDocument = savePath;

                totalMemoryDelta = GC.GetTotalMemory(false) - totalMemory;
                _logger.Debug("Managed memory after writing range to Excel worksheet (delta, MB): " + totalMemoryDelta / 1024 / 1024);
                _logger.Trace("Export to Excel: OK");


                // REPLACE
                try
                {

                }
                catch (Exception ex)
                {
                    _logger.Error("Export to Excel FAIL: " + ex.ToString());
                }
                finally
                {
                    CloseExcelAndReleaseComObjects(workbook, excelApp, excelObjs);
                }

                stopwatch.Stop();

                _logger.Trace("SelectAndLoadToExcel end");
            }

            /// <summary>
            /// Выгружает данные по документам. 
            /// Создает двумерный массив object с объемом, аналогичным объему данных, полученных в результате выполнения запроса к БД.
            /// Запись в файл Excel происходит в диапазон единоразово.
            /// </summary>
            /// <param name="selectOptions"></param>
            private void SelectAndLoadToExcel(DbSelectDocsOptions selectOptions)
            {
                _logger.Trace("SelectAndLoadToExcel begin. Load mode: single batch");

                LastExportRowCount = 0;
                BackgroundWorker worker = UnloadToExcelWorkerWrapper.Worker;

                long totalMemory = GC.GetTotalMemory(false);
                _logger.Debug("Managed memory on start unloading to excel (B, MB): " + totalMemory / 1024 / 1024);

                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                selectOptions.limit = int.MaxValue;
                string selectResult = _sqlServerExtResolver.GetUserDocuments(selectOptions);
                if (string.IsNullOrEmpty(selectResult))
                {
                    _logger.Trace("Select response string is null or empty");
                    return;
                }

                stopwatch.Stop();
                _logger.Trace("Time elapsed for: 1. SQL command execution; 2. Compressing xml result on server; 3. Decompressing string resut:" + Environment.NewLine +
                    "Total seconds: " + stopwatch.Elapsed.TotalSeconds);

                long totalMemoryDelta = GC.GetTotalMemory(false) - totalMemory;
                _logger.Debug("Managed memory after getting select result (delta, MB): " + totalMemoryDelta / 1024 / 1024);

                if (worker.WorkerSupportsCancellation && worker.CancellationPending)
                {
                    return;
                }

                stopwatch.Restart();

                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                Excel.Range creationDateColumn = null, regDateColumn = null;
                Excel.Range dataRange = null;

                List<object> excelObjs = new List<object>();
                excelObjs.Add(dataRange);
                excelObjs.Add(creationDateColumn);
                excelObjs.Add(regDateColumn);
                excelObjs.Add(worksheet);
                excelObjs.Add(workbook);
                excelObjs.Add(excelApp);

                excelApp.Visible = false;
                excelApp.ScreenUpdating = false;

                var colDefs = ColumnDefinitionsForTableDocs.DefaultDefs;
                string[] headers = new string[colDefs.Count];
                int colCount = 0;
                foreach (var colDef in colDefs)
                {
                    if (colDef.Key == ColumnDefinitionsForTableDocs.Column.CardId)
                    {
                        headers[colCount] = "Ссылка на карточку";
                    }
                    else
                    {
                        headers[colCount] = colDef.Value.displayName;
                    }

                    colCount++;
                }

                if (worker.WorkerSupportsCancellation && worker.CancellationPending)
                {
                    CloseExcelAndReleaseComObjects(workbook, excelApp, excelObjs);
                    return;
                }

                int startColumnIndex = 1;
                int endColumnIndex = colCount;

                Excel.Range startHeadersCell = (Excel.Range)worksheet.Cells[1, startColumnIndex];
                Excel.Range endHeadersCell = (Excel.Range)worksheet.Cells[1, endColumnIndex];

                Excel.Range headersRange = (Excel.Range)worksheet.Range[startHeadersCell, endHeadersCell];

                excelObjs.Insert(0, headersRange);
                excelObjs.Insert(0, endHeadersCell);
                excelObjs.Insert(0, startHeadersCell);

                // Запись заголовков столбцов
                //Excel.Range headersRange = worksheet.Range["A1", "N1"];
                headersRange.Value = headers;
                headersRange.Font.Bold = true;

                _logger.Trace("COL COUNT: " + colCount);

                stopwatch.Stop();
                _logger.Trace("Time elapsed for creating Excel doc in total seconds: " + stopwatch.Elapsed.TotalSeconds);

                totalMemory = GC.GetTotalMemory(false);
                _logger.Debug("Managed memory before xml read (MB): " + totalMemory / 1024 / 1024);

                stopwatch.Restart();

                _logger.Trace("XML read begin");

                List<object[]> rows = new List<object[]>();
                using (XmlReader reader = XmlReader.Create(new System.IO.StringReader(selectResult)))
                {
                    Type tDate = typeof(System.DateTime);

                    string readedValue;

                    while (reader.Read())
                    {
                        if (worker.WorkerSupportsCancellation && worker.CancellationPending)
                        {
                            CloseExcelAndReleaseComObjects(workbook, excelApp, excelObjs);
                            return;
                        }

                        if (reader.NodeType == XmlNodeType.Element && reader.Name == "R")
                        {
                            int i = 0;
                            object[] attrValues = new object[colCount];
                            foreach (var colDefKvp in colDefs)
                            {
                                readedValue = reader.GetAttribute(colDefKvp.Value.name) ?? "";

                                if (readedValue == _minDbDateStr)
                                {
                                    readedValue = "";
                                }

                                if (readedValue != "" && colDefKvp.Value.type == tDate)
                                {
                                    readedValue = readedValue.Replace('T', ' ');
                                }

                                if (colDefKvp.Key == ColumnDefinitionsForTableDocs.Column.CardId)
                                {
                                    attrValues[i] = _cardUrlHelper.GetUrl(readedValue);
                                }
                                else
                                {
                                    attrValues[i] = readedValue;
                                }

                                ++i;
                            }

                            rows.Add(attrValues);
                        }
                    }
                }
                LastExportRowCount = rows.Count;
                _logger.Trace("Rows: " + rows.Count);

                _logger.Trace("XML read end");

                stopwatch.Stop();
                _logger.Trace("Time elapsed for initializing object[] list in total seconds: " + stopwatch.Elapsed.TotalSeconds);

                totalMemoryDelta = GC.GetTotalMemory(false) - totalMemory;
                _logger.Debug("Managed memory after xml reading (delta, MB): " + totalMemoryDelta / 1024 / 1024);

                totalMemory = GC.GetTotalMemory(false);
                _logger.Debug("Managed memory before initializing data array (MB): " + totalMemory / 1024 / 1024);

                stopwatch.Reset();

                try
                {
                    if (colCount > 0 && rows.Count > 0)
                    {
                        int rowCount = rows.Count;

                        // Создаём двумерный массив для записи в диапазон ячеек Excel
                        object[,] data = new object[rowCount, colCount];
                        for (int i = 0; i < rowCount; i++)
                        {
                            if (worker.WorkerSupportsCancellation && worker.CancellationPending)
                            {
                                CloseExcelAndReleaseComObjects(workbook, excelApp, excelObjs);
                                return;
                            }

                            for (int j = 0; j < colCount; j++)
                            {
                                data[i, j] = rows[i][j];
                            }
                        }

                        stopwatch.Stop();
                        _logger.Trace("Time elapsed for initializing 2D object array in total seconds: " + stopwatch.Elapsed.TotalSeconds);

                        totalMemoryDelta = GC.GetTotalMemory(false) - totalMemory;
                        _logger.Debug("Managed memory after initializing data array (delta, MB): " + totalMemoryDelta / 1024 / 1024);
                        totalMemory = GC.GetTotalMemory(false);

                        stopwatch.Restart();

                        // Единовременная запись в диапазон
                        dataRange = (Excel.Range)worksheet.Range["A2", worksheet.Cells[rowCount + 1, colCount]];
                        dataRange.Value = data;

                        stopwatch.Stop();
                        _logger.Trace("Time elapsed for filling Excel range in bulk in total seconds: " + stopwatch.Elapsed.TotalSeconds);

                        stopwatch.Restart();

                        // Форматирование столбцов с датами
                        creationDateColumn = (Excel.Range)worksheet.Columns[4];
                        creationDateColumn.NumberFormat = "dd.mm.yyyy hh:mm:ss";
                        regDateColumn = (Excel.Range)worksheet.Columns[5];
                        regDateColumn.NumberFormat = "dd.mm.yyyy hh:mm:ss";

                        //excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                        excelApp.ScreenUpdating = true;

                        string savePath = GetExcelSavePath();

                        _logger.Trace("SavePath to Excel=" + savePath);
                        workbook.SaveAs(savePath);

                        stopwatch.Stop();
                        _logger.Trace("Time elapsed for saving Excel doc in total seconds: " + stopwatch.Elapsed.TotalSeconds);

                        LastUnloadedXlsDocument = savePath;

                        totalMemoryDelta = GC.GetTotalMemory(false) - totalMemory;
                        _logger.Debug("Managed memory after writing range to Excel worksheet (delta, MB): " + totalMemoryDelta / 1024 / 1024);
                        _logger.Trace("Export to Excel: OK");
                    }
                    else
                    {

                    }
                }
                catch (Exception ex)
                {
                    _logger.Error("Export to Excel FAIL: " + ex.ToString());
                }
                finally
                {
                    CloseExcelAndReleaseComObjects(workbook, excelApp, excelObjs);
                }

                stopwatch.Stop();

                _logger.Trace("SelectAndLoadToExcel end");
            }

            private System.Data.DataTable CreateEmptyTable()
            {
                System.Data.DataTable dt = new System.Data.DataTable("UserDocuments");
                var colDefs = ColumnDefinitionsForTableDocs.DefaultDefs;
                foreach (var colDef in colDefs)
                {
                    dt.Columns.Add(new DataColumn(colDef.Value.name, colDef.Value.type));
                }

                return dt;
            }

            private void LoadTableDataWorker_RunCompleted(object sender, RunWorkerCompletedEventArgs e)
            {
                _logger.Trace("LoadTableDataWorker_RunCompleted begin");

                ShowLoadingOverlay(false);

                try
                {
                    if (e.Error != null)
                    {
                        _logger.Error("Ошибка при загрузке данных: " + e.Error.ToString());
                        return;
                    }

                    System.Data.DataTable table = e.Result as System.Data.DataTable;

                    var gridControl = Control;

                    if (table == null || table.Rows.Count == 0)
                    {
                        gridControl.DataSource = null;
                    }
                    else
                    {
                        gridControl.DataSource = table;
                    }

                    LogMemoryUsage("After LoadTableDataWorker_RunCompleted", _logger);

                    OnTableLoaded(table);
                }
                catch (Exception ex)
                {
                    _logger.Error("LoadTableDataWorker_RunCompleted failed with error: " + ex.ToString());
                }
            }

            private void LoadTableDataWorker_DoWork(object sender, DoWorkEventArgs e)
            {
                //EnableSelectCommands(false);

                System.Data.DataTable dt = LoadDocumentsData(e);
                //DataTable dt = LoadDataFromFile("C:\\Users\\KokurinMR\\Documents\\2.txt");
                if (dt != null)
                {
                    _logger.Info("Total rows in DataTable after search=" + dt.Rows.Count);

                    //DocumentsCountItem.ControlValue = dt.Rows.Count;
                }
                else
                {
                    //DocumentsCountItem.ControlValue = 0;
                }

                e.Result = dt;
            }

            private System.Data.DataTable LoadDocumentsData(DoWorkEventArgs e)
            {
                DbSelectDocsOptions selectOptions;
                if (e.Argument != null)
                {
                    try
                    {
                        selectOptions = (DbSelectDocsOptions)e.Argument;
                    }
                    catch (Exception ex)
                    {
                        _logger.Error("Ошибка при приведении аргумента BackgroundWorker к типу " + typeof(DbSelectDocsOptions) + Environment.NewLine + ex.ToString());
                        return null;
                    }

                    string responseStr = _sqlServerExtResolver.GetUserDocuments(selectOptions);
                    //_logger.Trace("Select response (from next line):" + Environment.NewLine + responseStr);

                    if (LoadTableDataWorkerWrapper.Worker.WorkerSupportsCancellation && LoadTableDataWorkerWrapper.Worker.CancellationPending)
                    {
                        e.Cancel = true;
                        return null;
                    }

                    if (string.IsNullOrEmpty(responseStr))
                    {
                        return null;
                    }

                    System.Data.DataTable table = ConvertXmlToDataTable(responseStr);
                    return table;
                }
                else
                {
                    return null;
                }
            }

            /// <summary>
            /// Загужает данные для таблицы из файла, содержащего результат выполнения запроса к БД
            /// </summary>
            /// <param name="filePath"></param>
            /// <returns></returns>
            private System.Data.DataTable LoadTableData(string filePath)
            {
                if (!System.IO.File.Exists(filePath))
                {
                    _logger.Error("Указанного файла не существует. FilePath=" + (filePath == null ? "NULL" : filePath));
                    return null;
                }

                StringBuilder xmlContent = new StringBuilder(System.IO.File.ReadAllText(filePath));
                //_logger.Trace("File content (from next line):" + Environment.NewLine + xmlContent.ToString());

                if (string.IsNullOrEmpty(xmlContent.ToString()))
                {
                    return null;
                }

                System.Data.DataTable table = ConvertXmlToDataTable(xmlContent.Replace("\n", "").Replace("\t", "").Replace("\r", "").ToString());

                return table;
            }

            private System.Data.DataTable ConvertXmlToDataTable(string xml)
            {
                _logger.Trace("ConvertXmlToDocsDataTable begin");

                long totalMemory = GC.GetTotalMemory(false);
                _logger.Debug("Managed memory on start converting xml to DataTable (MB): " + totalMemory / 1024 / 1024);

                System.Data.DataTable table = CreateEmptyTable();

                BackgroundWorker worker = LoadTableDataWorkerWrapper.Worker;

                using (var reader = XmlReader.Create(new System.IO.StringReader(xml)))
                {
                    table.BeginLoadData();

                    while (reader.Read())
                    {
                        if (worker.WorkerSupportsCancellation && worker.CancellationPending)
                        {
                            table.EndLoadData();
                            return null;
                        }

                        // Ищем элемент <R>
                        if (reader.NodeType == XmlNodeType.Element && reader.Name == "R")
                        {
                            DataRow row = table.NewRow();

                            // Заполняем строку, проверяя наличие каждого атрибута
                            if (reader.HasAttributes)
                            {
                                for (int i = 0; i < reader.AttributeCount; i++)
                                {
                                    reader.MoveToAttribute(i);
                                    string attrName = reader.Name;
                                    string attrValue = reader.Value;

                                    SetValueForDataRow(row, attrName, attrValue, false);
                                }
                                reader.MoveToElement(); // Возвращаемся к элементу <R>
                            }

                            table.Rows.Add(row);
                        }
                    }

                    table.EndLoadData();
                }

                table.AcceptChanges();

                _logger.Debug("Managed memory on finish converting xml to DataTable (delta, MB): " + (GC.GetTotalMemory(false) - totalMemory) / 1024 / 1024);
                _logger.Debug("Rows count: " + table.Rows.Count);

                _logger.Trace("ConvertXmlToDocsDataTable end");

                return table;
            }

            // TO DO: replace string litarals and use ColumnDefinitionsForTableDocs
            private void SetValueForDataRow(DataRow row, string columnName, string value, bool checkColumnName)
            {
                if (checkColumnName && !row.Table.Columns.Contains(columnName)) return;

                switch (columnName)
                {
                    case "Name":
                        if (string.IsNullOrEmpty(value))
                        {
                            object typeData = row["Type"];
                            string typeName = (typeData is System.DBNull || typeData == null) ? null : (string)typeData;

                            if (string.IsNullOrEmpty(typeName))
                            {
                                row[columnName] = "Документ";
                            }
                            else
                            {
                                row[columnName] = typeName;
                            }
                        }
                        else
                        {
                            row[columnName] = value;
                        }

                        break;

                    case "Type":
                        if (string.IsNullOrEmpty(value))
                        {
                            row[columnName] = "";
                        }
                        else
                        {
                            row[columnName] = value;

                            object nameData = row["Name"];
                            string name = (nameData is System.DBNull || nameData == null) ? null : (string)nameData;

                            if (string.IsNullOrEmpty(name) || name == "Документ")
                            {
                                row["Name"] = value;
                            }
                        }

                        break;

                    case "CreationDate":
                    case "RegDate":
                        DateTime tempDate;
                        if (!DateTime.TryParseExact(value, dateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out tempDate) || tempDate.Equals(_minDBDateTime))
                        {
                            row[columnName] = DBNull.Value;
                        }
                        else
                        {
                            row[columnName] = tempDate;
                        }

                        break;

                    default:
                        row[columnName] = value ?? "";
                        break;
                }
            }

            private void ShowLoadingOverlay(bool show)
            {
                _logger.Trace("ShowLoadingOverlay begin. Show: " + show);

                if (_loadingOverlay == null)
                {
                    _loadingOverlay = CreateLoadingOverlay();
                }

                _loadingOverlay.Visible = show;

                _logger.Trace("ShowLoadingOverlay end");
            }

            private System.Windows.Forms.Label CreateLoadingOverlay()
            {
                _logger.Trace("AddLoadingOverlay begin");

                // Создаем Label для сообщения
                var loadingLabel = new System.Windows.Forms.Label
                {
                    Text = "Идет загрузка, пожалуйста подождите...",
                    Font = new System.Drawing.Font("Segoe UI", 12, System.Drawing.FontStyle.Bold),
                    ForeColor = System.Drawing.Color.DarkGray,
                    BackColor = System.Drawing.Color.FromArgb(240, 240, 240, 240), // Полупрозрачный фон
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                    Dock = System.Windows.Forms.DockStyle.Fill,
                    Visible = true
                };

                // Находим родительский контейнер GridControl
                var gridControl = Control;
                var parentControl = gridControl.Parent;

                // Добавляем Label поверх GridControl
                parentControl.Controls.Add(loadingLabel);
                loadingLabel.BringToFront();

                // Привязываем размеры Label к GridControl
                loadingLabel.Location = gridControl.Location;
                loadingLabel.Size = gridControl.Size;

                _logger.Trace("AddLoadingOverlay end");
                return loadingLabel;
            }

            private void GridControl_DataSourceChanged(object sender, EventArgs e)
            {
                _logger.Trace("GridControl_DataSourceChanged begin");

                _viewHandler.SelectedDocumentId = Guid.Empty;
                //_filesGridWrapper.SelectedFileId = Guid.Empty;

                DevExpress.XtraGrid.GridControl control = sender as DevExpress.XtraGrid.GridControl;
                var view = control.MainView as DocsVision.BackOffice.WinForms.Controls.GridExView;

                if (view != null)
                {
                    if (control.DataSource != null)
                    {
                        //view.OptionsView.ShowGroupPanel = true;
                        //control.DataSource = CreateEmptyTable();

                        layoutItemHelper.GetItem<BaseLayoutItem>(LayoutItemAlias.GroupPanelVisibilitySwitch).Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                        if (view.DataRowCount > 0)
                        {
                            layoutItemHelper.GetItem<BaseLayoutItem>(LayoutItemAlias.ColumnVisibilitySwitch).Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                        }
                        else
                        {
                            layoutItemHelper.GetItem<BaseLayoutItem>(LayoutItemAlias.ColumnVisibilitySwitch).Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;
                        }

                        if (view.GroupCount != 0)
                        {
                            layoutItemHelper.GetItem<BaseLayoutItem>(LayoutItemAlias.CollapseExpandGroupsButton).Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                        }
                    }
                    else
                    {
                        //view.OptionsView.ShowGroupPanel = false;

                        layoutItemHelper.GetItem<BaseLayoutItem>(LayoutItemAlias.ColumnVisibilitySwitch).Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;
                        layoutItemHelper.GetItem<BaseLayoutItem>(LayoutItemAlias.GroupPanelVisibilitySwitch).Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;
                        layoutItemHelper.GetItem<BaseLayoutItem>(LayoutItemAlias.CollapseExpandGroupsButton).Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.OnlyInCustomization;
                    }

                    //GridDataManager.ProcessDataSourceChange(control);
                    _viewHandler.RestoreView();
                }

                _logger.Trace("GridControl_DataSourceChanged end");
            }
        }

        private void tryLoadImage_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            //XtraInputBoxArgs args = new XtraInputBoxArgs();
            //args.Editor.va
            //object data = XtraInputBox.Show(args);

            string input = XtraInputBox.Show
            (
                "Введите несжатый base64String изображения:",  // Подсказка над полем ввода
                "Тест загрузки изображения",                 // Заголовок окна
                ""    // Текст по умолчанию в поле
            );

            if (!string.IsNullOrEmpty(input))
            {
                try
                {
                    Image image = Utils.StringToImage(input, false);
                    if (image != null)
                    {
                        CardControl.ShowMessage("Изображение успешно загружено", "Тест загрузки изображения", DocsVision.Platform.CardHost.MessageType.Information, DocsVision.Platform.CardHost.MessageButtons.Ok);
                    }
                }
                catch (Exception ex)
                {
                    CardControl.ShowMessage("Ошибка загрузки изображения", "Тест загрузки изображения", ex.ToString(), DocsVision.Platform.CardHost.MessageType.Error, DocsVision.Platform.CardHost.MessageButtons.Ok);
                }
            }
        }

        private void copyBase64StringFromImage_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Images only
            openFileDialog.Filter =
                "Изображения (*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.tiff;*.ico;*.svg)|" +
                "*.bmp;*.jpg;*.jpeg;*.gif;*.png;*.tiff;*.ico;*.svg";

            openFileDialog.Title = "Выберите изображения";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);

            openFileDialog.CheckFileExists = true;
            openFileDialog.CheckPathExists = true;
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string file = openFileDialog.FileName;
                Clipboard.SetText(Utils.ImageToBase64String(file));
            }
        }

        private void saveImageFromBase64String_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("saveImageFromBase64String_ItemClick begin");

            string folder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);

            string input = XtraInputBox.Show(
                "Введите несжатый base64String изображения",
                "Сохранение изображения",
                ""
                );

            if (!string.IsNullOrWhiteSpace(input))
            {
                int maxPostfixLen = 8;
                int len = input.Length < maxPostfixLen ? input.Length : maxPostfixLen;
                string ext = ".png";
                string path = Path.Combine(folder, "Image_" + input.Substring(0, len) + ext);

                try
                {
                    Image image = Utils.StringToImage(input);
                    image.Save(path);
                }
                catch (Exception ex)
                {
                    CardControl.ShowMessage("Ошибка сохранения изображения", "Сохранение изображения", ex.ToString(), DocsVision.Platform.CardHost.MessageType.Error, DocsVision.Platform.CardHost.MessageButtons.Ok);
                }
            }

            _logger.Trace("saveImageFromBase64String_ItemClick end");
        }

        private void checkGroupExpandStateInGridView_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            bool hasCollapsedGroupRow = !Utils.IsAllGroupRowsExpanded(_xtraGridRepository[Table.UserDocuments].MainView);
            CardControl.ShowMessage("Развернуты все группы: " + (hasCollapsedGroupRow ? "нет" : "да"), "Результат проверки раскрытых групп в таблице с документами", DocsVision.Platform.CardHost.MessageType.Information, DocsVision.Platform.CardHost.MessageButtons.Ok);
        }

        #endregion
    }
}