using DevExpress.Utils;
using DevExpress.XtraEditors;
using DevExpress.XtraLayout;
using DevExpress.XtraLayout.Utils;
using DocsVision.BackOffice.ObjectModel;
using DocsVision.BackOffice.WinForms;
using DocsVision.BackOffice.WinForms.Controls;
using DocsVision.BackOffice.WinForms.Design.LayoutItems;
using DocsVision.BackOffice.WinForms.Design.PropertyControls;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectModel;
using MLog;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Padding = DevExpress.XtraLayout.Utils.Padding;

namespace BackOffice
{
    public class CardDocumentСозданные_мнойScript : CardDocumentScript
    {

        #region Properties

        private Logger _logger;
        private readonly Guid _testCardId = new Guid("B8BEE5B4-FCD5-F011-AFF0-000C290CEAEE");
        private Guid[] _testFilesIds = new Guid[]
        {
            new Guid("5521511A-5811-E211-9F49-1000000198A0"),
            new Guid("DB0A81C3-FCD5-F011-AFF0-000C290CEAEE"),
            new Guid("E40A81C3-FCD5-F011-AFF0-000C290CEAEE"),
            new Guid("FC43B0FF-FA11-E211-9F49-1000000198A0")
        };

        private string _layoutXmlPath = "C:\\Users\\Администратор\\Desktop\\GAZ_Layout.xml";
        private string _pathToTestLayoutXml = "C:\\Users\\Администратор\\Desktop\\Test_Layout.xml";

        private string _filePreviewItemText = "Просмотр файла";

        private DevExpress.XtraLayout.LayoutControlGroup FilesLinksGroup
        {
            get
            {
                return GetFirtstXLayoutItemWithText<DevExpress.XtraLayout.LayoutControlGroup>("Файлы");
            }
        }

        private DevExpress.XtraLayout.LayoutControlGroup FilePreviewGroup
        {
            get
            {
                return GetFirtstXLayoutItemWithText<DevExpress.XtraLayout.LayoutControlGroup>("Просмотр файла");
            }
        }

        private ICustomizableControl CustomizableControl
        {
            get
            {
                return this.CardControl as ICustomizableControl;
            }
        }

        private object FilePreviewAsObj
        {
            get
            {
                return CustomizableControl.FindLayoutItem("FilePreview");
            }
        }


        private ILayoutPropertyItem FilePreviewAsLayoutPropertyItem
        {
            get
            {
                return CustomizableControl.FindPropertyItem<ILayoutPropertyItem>("FilePreview");
            }
        }

        private IPreviewFileControl FilePreview
        {
            get
            {
                return CustomizableControl.FindPropertyItem<IPreviewFileControl>("FilePreview");
            }
        }

        #endregion

        public CardDocumentСозданные_мнойScript() : base()
        {
            if (_logger == null)
                _logger = new Logger("C:\\Logs\\CardScriptLog.log", "ReportOnMyCards");
            else
                _logger.Warn("Logger has been already initialized");

            _logger.Trace("Script instantiated");
        }

        #region Methods

        private void TestDocFile()
        {
            CardData cardData = Session.CardManager.GetCardData(_testCardId);
            _logger.Trace("CardData received");

            BaseCard card = CardControl.ObjectContext.GetObject<BaseCard>(_testCardId);
            _logger.Trace("BaseCard received");

            Document document = CardControl.ObjectContext.GetObject<Document>(_testCardId);
            _logger.Trace("Document received");
            Guid mainFileId = document.MainInfo.FileId;
            _logger.Trace("mainFileId=" + mainFileId.ToString());
        }

        private void Test1()
        {
            _logger.Trace("Test1 begin");

            

            _logger.Trace("Test1 end");
        }

        private void TestLayoutGroup()
        {
            _logger.Trace("TestLayoutGroup begin");

            try
            {
                //var group = Customizable.FindLayoutItem("FilesGroup"); NO
            }
            catch (Exception ex)
            {
                _logger.Error("TestLayoutGroup(): " + ex.ToString());
            }

            _logger.Trace("TestLayoutGroup end");
        }

        private string _fileLinkTag = "FileLink";
        private struct FileLinkTag
        {
            public string tagName;
            public Guid fileId;
        }

        private int AddDummyFilesLinks_MultiColumn(DevExpress.XtraLayout.LayoutControlGroup group, int linksCount)
        {
            _logger.Trace("AddDummyFilesLinks_MultiColumn begin");

            if (group == null || linksCount <= 0)
                return 0;

            string fileName = "...";
            int idsLen = _testFilesIds.Length;

            group.Padding = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);
            group.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

            var scroll = new DevExpress.XtraEditors.XtraScrollableControl();
            scroll.Dock = DockStyle.Fill;
            scroll.AutoScroll = true;

            var innerLayout = new DevExpress.XtraLayout.LayoutControl();
            innerLayout.Dock = DockStyle.Fill;          // теперь заполняем по ширине контейнера
            innerLayout.AutoScroll = false;

            var root = innerLayout.Root;
            root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            root.GroupBordersVisible = false;
            root.Padding = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);
            root.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

            // ВКЛЮЧАЕМ ТАБЛИЧНУЮ РАЗМЕТКУ
            root.LayoutMode = DevExpress.XtraLayout.Utils.LayoutMode.Table;

            // допустим, хотим до 3 ссылок в строке
            var cols = root.OptionsTableLayoutGroup.ColumnDefinitions;
            cols.Clear();
            for (int c = 0; c < 3; c++)
            {
                var col = new DevExpress.XtraLayout.ColumnDefinition();
                col.SizeType = System.Windows.Forms.SizeType.Percent;
                col.Width = 100f / 3f;          // три равные колонки
                cols.Add(col);
            }

            var rows = root.OptionsTableLayoutGroup.RowDefinitions;
            rows.Clear();
            // строки можно добавлять по мере необходимости (см. ниже)

            scroll.Controls.Add(innerLayout);

            CustomizableControl.LayoutControl.BeginUpdate();
            group.BeginUpdate();

            var scrollItem = group.AddItem();
            scrollItem.Control = scroll;
            scrollItem.TextVisible = false;
            scrollItem.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            scrollItem.MinSize = new Size(10, 40);
            scrollItem.MaxSize = new Size(0, 0);
            scrollItem.Padding = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);
            scrollItem.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

            int colIndex = 0;
            int rowIndex = 0;

            // заранее создаём нужное количество строк
            int rowsCount = (linksCount + 2) / 3;   // по 3 ссылки в строке
            for (int r = 0; r < rowsCount; r++)
            {
                var row = new DevExpress.XtraLayout.RowDefinition();
                row.SizeType = System.Windows.Forms.SizeType.AutoSize;
                rows.Add(row);
            }

            for (int i = 0; i < linksCount; i++)
            {
                var linkLabel = new DevExpress.XtraEditors.LabelControl();
                linkLabel.Text = i + "_" + fileName;
                linkLabel.Appearance.ForeColor = Color.Blue;
                linkLabel.Appearance.Font = new Font(linkLabel.Appearance.Font, FontStyle.Underline);
                linkLabel.Cursor = Cursors.Hand;
                linkLabel.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
                linkLabel.Appearance.TextOptions.Trimming = DevExpress.Utils.Trimming.EllipsisCharacter;
                linkLabel.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;

                linkLabel.Tag = new FileLinkTag
                {
                    tagName = _fileLinkTag,
                    fileId = _testFilesIds[i % idsLen]
                };
                linkLabel.Click += FileLink_Click;

                var item = root.AddItem();
                item.Control = linkLabel;
                item.TextVisible = false;
                item.Padding = new DevExpress.XtraLayout.Utils.Padding(4, 0, 0, 0);
                item.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

                // привязка к ячейке таблицы
                var tbl = item.OptionsTableLayoutItem;
                tbl.RowIndex = rowIndex;
                tbl.ColumnIndex = colIndex;

                colIndex++;
                if (colIndex >= 3)
                {
                    colIndex = 0;
                    rowIndex++;
                }
            }

            group.EndUpdate();
            CustomizableControl.LayoutControl.EndUpdate();

            _logger.Trace("AddDummyFilesLinks_MultiColumn end");
            return linksCount;
        }

        private int AddDummyFilesLinks6(DevExpress.XtraLayout.LayoutControlGroup group, int linksCount)
        {
            _logger.Trace("AddDummyFilesLinks6 begin");

            if (group == null || linksCount <= 0)
                return 0;

            string fileName = "sefsfgsjgnsgnsegnsekrgnekrgnekgnekngwejrngeiwgnreiwgnwinwphgnptgnwe.docx";
            int idsLen = _testFilesIds.Length;

            // убираем лишние отступы у группы
            group.Padding = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);
            group.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

            var scroll = new DevExpress.XtraEditors.XtraScrollableControl();
            scroll.Dock = DockStyle.Fill;
            scroll.AutoScroll = true;

            var innerLayout = new DevExpress.XtraLayout.LayoutControl();
            innerLayout.Dock = DockStyle.None;
            innerLayout.AutoScroll = false;

            innerLayout.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            innerLayout.Root.GroupBordersVisible = false;
            innerLayout.Root.Padding = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);
            innerLayout.Root.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

            scroll.Controls.Add(innerLayout);

            CustomizableControl.LayoutControl.BeginUpdate();
            group.BeginUpdate();

            var scrollItem = group.AddItem();
            scrollItem.Control = scroll;
            scrollItem.TextVisible = false;
            scrollItem.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            scrollItem.MinSize = new Size(10, 40);
            scrollItem.MaxSize = new Size(0, 0);
            scrollItem.Padding = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);
            scrollItem.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

            int maxTextWidth = 0;
            DevExpress.XtraLayout.LayoutControlItem lastItem = null;

            for (int i = 0; i < linksCount; i++)
            {
                var linkLabel = new DevExpress.XtraEditors.LabelControl();
                linkLabel.Text = i + "_" + fileName;
                linkLabel.Appearance.ForeColor = Color.Blue;
                linkLabel.Appearance.Font = new Font(linkLabel.Appearance.Font, FontStyle.Underline);
                linkLabel.Cursor = Cursors.Hand;
                linkLabel.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
                linkLabel.Appearance.TextOptions.Trimming = DevExpress.Utils.Trimming.EllipsisCharacter;
                linkLabel.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;

                linkLabel.Tag = new FileLinkTag
                {
                    tagName = _fileLinkTag,
                    fileId = _testFilesIds[i % idsLen]
                };
                linkLabel.Click += FileLink_Click;

                var item = innerLayout.Root.AddItem();
                item.Control = linkLabel;
                item.TextVisible = false;
                item.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
                item.MinSize = new Size(0, 18);
                item.MaxSize = new Size(0, 0);
                item.Padding = new DevExpress.XtraLayout.Utils.Padding(4, 0, 0, 0); // 4px слева
                item.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

                Size best = linkLabel.CalcBestSize();
                maxTextWidth = Math.Max(maxTextWidth, best.Width);

                lastItem = item;
            }

            // Даём LayoutControl полностью переложить элементы
            innerLayout.ResumeLayout(false);
            innerLayout.PerformLayout();           // пересчёт разметки [file:2]

            // Ширина — по самому длинному тексту, небольшой запас
            int contentWidth = maxTextWidth + 10;

            // Высота — по последнему item’у: его Y + высота
            int contentHeight = 0;
            if (lastItem != null)
            {
                var r = lastItem.Bounds;          // реальные координаты в LayoutControl [file:2]
                contentHeight = r.Bottom + 2;     // +2px технический запас
            }

            if (contentHeight < 40)
                contentHeight = 40;

            innerLayout.Size = new Size(contentWidth, contentHeight);

            // Подгоняем ширину всех LabelControl под contentWidth
            foreach (DevExpress.XtraLayout.BaseLayoutItem it in innerLayout.Root.Items)
            {
                var lcItem = it as DevExpress.XtraLayout.LayoutControlItem;

                if (lcItem == null || !(lcItem.Control is DevExpress.XtraEditors.LabelControl))
                {
                    continue;
                }

                var lbl = lcItem.Control as DevExpress.XtraEditors.LabelControl;

                Size best = lbl.CalcBestSize();
                lbl.Size = new Size(contentWidth, best.Height);
            }

            scroll.AutoScrollMinSize = new Size(contentWidth, contentHeight);

            group.EndUpdate();
            CustomizableControl.LayoutControl.EndUpdate();

            _logger.Trace("AddDummyFilesLinks6 end");
            return linksCount;
        }

        private int AddDummyFilesLinks5(DevExpress.XtraLayout.LayoutControlGroup group, int linksCount)
        {
            _logger.Trace("AddDummyFilesLinks5 begin");

            if (group == null || linksCount <= 0)
                return 0;

            string fileName = "sefsfgsjgnsgnsegnsekrgnekrgnekgnekngwejrngeiwgnreiwgnwinwphgnptgnwe.docx";
            int idsLen = _testFilesIds.Length;

            // минимальные отступы у самой группы
            group.Padding = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);
            group.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

            var scroll = new DevExpress.XtraEditors.XtraScrollableControl();
            scroll.Dock = DockStyle.Fill;
            scroll.AutoScroll = true;          // один общий скролл

            var innerLayout = new DevExpress.XtraLayout.LayoutControl();
            innerLayout.Dock = DockStyle.None; // размер задаём вручную
            innerLayout.AutoScroll = false;

            // корневой group внутри layout
            innerLayout.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            innerLayout.Root.GroupBordersVisible = false;
            innerLayout.Root.Padding = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);
            innerLayout.Root.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

            scroll.Controls.Add(innerLayout);

            CustomizableControl.LayoutControl.BeginUpdate();
            group.BeginUpdate();

            var scrollItem = group.AddItem();
            scrollItem.Control = scroll;
            scrollItem.TextVisible = false;
            scrollItem.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            scrollItem.MinSize = new Size(10, 40);
            scrollItem.MaxSize = new Size(0, 0);
            scrollItem.Padding = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);
            scrollItem.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

            int maxTextWidth = 0;
            int totalHeight = 0;

            // запас по высоте на каждую строку
            const int perItemExtraHeight = 6;
            const int bottomExtraHeight = 6;

            for (int i = 0; i < linksCount; i++)
            {
                var linkLabel = new DevExpress.XtraEditors.LabelControl();
                linkLabel.Text = i + "_" + fileName;
                linkLabel.Appearance.ForeColor = Color.Blue;
                linkLabel.Appearance.Font = new Font(linkLabel.Appearance.Font, FontStyle.Underline);
                linkLabel.Cursor = Cursors.Hand;
                linkLabel.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
                linkLabel.Appearance.TextOptions.Trimming = DevExpress.Utils.Trimming.EllipsisCharacter;
                linkLabel.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;

                linkLabel.Tag = new FileLinkTag
                {
                    tagName = _fileLinkTag,
                    fileId = _testFilesIds[i % idsLen]
                };
                linkLabel.Click += FileLink_Click;

                var item = innerLayout.Root.AddItem();
                item.Control = linkLabel;
                item.TextVisible = false;
                item.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
                item.MinSize = new Size(0, 18);
                item.MaxSize = new Size(0, 0);
                // отступ слева 2 px
                item.Padding = new DevExpress.XtraLayout.Utils.Padding(4, 0, 0, 0);
                item.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

                // замер «естественного» размера текста
                Size best = linkLabel.CalcBestSize();
                maxTextWidth = Math.Max(maxTextWidth, best.Width);
                totalHeight += best.Height + perItemExtraHeight;
            }

            // ширина по самому длинному тексту, небольшой запас справа
            int contentWidth = maxTextWidth + 10;
            // высота с запасом, чтобы точно влезли все linksCount строк
            int contentHeight = totalHeight + bottomExtraHeight;

            // задаём физический размер layout’а по контенту
            innerLayout.Size = new Size(contentWidth, contentHeight);

            // чтобы текст не обрезался, выровнять ширину всех LabelControl
            foreach (DevExpress.XtraLayout.BaseLayoutItem it in innerLayout.Root.Items)
            {
                var lcItem = it as DevExpress.XtraLayout.LayoutControlItem;
                if (lcItem != null && lcItem.Control is DevExpress.XtraEditors.LabelControl)
                {
                    DevExpress.XtraEditors.LabelControl lbl = lcItem.Control as DevExpress.XtraEditors.LabelControl;
                    Size best = lbl.CalcBestSize();
                    lbl.Size = new Size(contentWidth, best.Height);
                }
            }

            // скролл видит ровно «тело» ссылок, без пустых хвостов
            scroll.AutoScrollMinSize = new Size(contentWidth, contentHeight);

            group.EndUpdate();
            CustomizableControl.LayoutControl.EndUpdate();

            _logger.Trace("AddDummyFilesLinks5 end");
            return linksCount;
        }

        private int AddDummyFilesLinks3(DevExpress.XtraLayout.LayoutControlGroup group, int linksCount)
        {
            _logger.Trace("AddDummyFilesLinks3 begin");

            if (group == null || linksCount <= 0)
                return 0;

            string fileName = "sefsfgsjgnsgnsegnsekrgnekrgnekgnekngwejrngeiwgnreiwgnwinwphgnptgnwe.docx";
            int idsLen = _testFilesIds.Length;

            var scroll = new DevExpress.XtraEditors.XtraScrollableControl();
            scroll.Dock = DockStyle.Fill;
            scroll.AutoScroll = true;                      // один общий скролл (гор+верт)

            var innerLayout = new DevExpress.XtraLayout.LayoutControl();
            innerLayout.Dock = DockStyle.None;             // ширина по содержимому
            innerLayout.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            innerLayout.Root.GroupBordersVisible = false;
            innerLayout.AutoScroll = false;               // ВАЖНО: без собственного скролла

            scroll.Controls.Add(innerLayout);

            CustomizableControl.LayoutControl.BeginUpdate();
            group.BeginUpdate();

            var scrollItem = group.AddItem();
            scrollItem.Control = scroll;
            scrollItem.TextVisible = false;
            scrollItem.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;

            // даём минимальную высоту, но НЕ ограничиваем сверху,
            // чтобы сплиттер мог растягивать/сжимать группу
            scrollItem.MinSize = new Size(10, 40);
            scrollItem.MaxSize = new Size(0, 0);          // без верхнего ограничения

            int maxTextWidth = 0;
            int totalHeight = 0;

            for (int i = 0; i < linksCount; i++)
            {
                var linkLabel = new DevExpress.XtraEditors.LabelControl();
                linkLabel.Text = i + "_" + fileName;
                linkLabel.Appearance.ForeColor = Color.Blue;
                linkLabel.Appearance.Font = new Font(linkLabel.Appearance.Font, FontStyle.Underline);
                linkLabel.Cursor = Cursors.Hand;
                linkLabel.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
                linkLabel.Appearance.TextOptions.Trimming = DevExpress.Utils.Trimming.EllipsisCharacter;
                linkLabel.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;

                linkLabel.Tag = new FileLinkTag
                {
                    tagName = _fileLinkTag,
                    fileId = _testFilesIds[i % idsLen]
                };
                linkLabel.Click += FileLink_Click;

                var item = innerLayout.Root.AddItem();
                item.Control = linkLabel;
                item.TextVisible = false;
                item.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
                item.MinSize = new Size(0, 18);
                item.MaxSize = new Size(0, 0);
                item.Padding = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);
                item.Spacing = new DevExpress.XtraLayout.Utils.Padding(0, 0, 0, 0);

                var sz = linkLabel.CalcBestSize();
                maxTextWidth = Math.Max(maxTextWidth, sz.Width);
                totalHeight += sz.Height;
            }

            int contentWidth = maxTextWidth + 20;
            int contentHeight = totalHeight + 10;

            // фиксируем размер именно по контенту
            innerLayout.Size = new Size(contentWidth, contentHeight);

            // скролл видит только реальный контент, без «хвоста»
            scroll.AutoScrollMinSize = new Size(contentWidth, contentHeight);

            group.EndUpdate();
            CustomizableControl.LayoutControl.EndUpdate();

            _logger.Trace("AddDummyFilesLinks3 end");
            return linksCount;
        }

        private int AddDummyFilesLinks2(DevExpress.XtraLayout.LayoutControlGroup group, int linksCount)
        {
            _logger.Trace("AddDummyFilesLinks2 begin");

            if (group == null || linksCount <= 0)
                return 0;

            string fileName = "sefsfgsjgnsgnsegnsekrgnekrgnekgnekngwejrngeiwgnreiwgnwinwphgnptgnwe.docx";
            int idsLen = _testFilesIds.Length;

            var scroll = new DevExpress.XtraEditors.XtraScrollableControl();
            scroll.Dock = DockStyle.Fill;
            scroll.AutoScroll = true;                      // один общий скролл (гор+верт)

            var innerLayout = new DevExpress.XtraLayout.LayoutControl();
            innerLayout.Dock = DockStyle.Top;             // ширина по содержимому
            innerLayout.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            innerLayout.Root.GroupBordersVisible = false;
            innerLayout.AutoScroll = false;               // ВАЖНО: без собственного скролла

            scroll.Controls.Add(innerLayout);

            CustomizableControl.LayoutControl.BeginUpdate();
            group.BeginUpdate();

            var scrollItem = group.AddItem();
            scrollItem.Control = scroll;
            scrollItem.TextVisible = false;
            scrollItem.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;

            // даём минимальную высоту, но НЕ ограничиваем сверху,
            // чтобы сплиттер мог растягивать/сжимать группу
            scrollItem.MinSize = new Size(10, 40);
            scrollItem.MaxSize = new Size(0, 0);          // без верхнего ограничения

            int maxTextWidth = 0;
            int totalHeight = 0;

            for (int i = 0; i < linksCount; i++)
            {
                var linkLabel = new DevExpress.XtraEditors.LabelControl();
                linkLabel.Text = i + "_" + fileName;
                linkLabel.Appearance.ForeColor = Color.Blue;
                linkLabel.Appearance.Font = new Font(linkLabel.Appearance.Font, FontStyle.Underline);
                linkLabel.Cursor = Cursors.Hand;
                linkLabel.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.Default;

                linkLabel.Tag = new FileLinkTag
                {
                    tagName = _fileLinkTag,
                    fileId = _testFilesIds[i % idsLen]
                };
                linkLabel.Click += FileLink_Click;

                var item = innerLayout.Root.AddItem();
                item.Control = linkLabel;
                item.TextVisible = false;
                item.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
                item.MinSize = new Size(10, 18);
                item.MaxSize = new Size(0, 0);

                var sz = linkLabel.CalcBestSize();
                maxTextWidth = Math.Max(maxTextWidth, (int)sz.Width);
                totalHeight += (int)sz.Height;
            }

            // настраиваем область прокрутки У scroll (а не у innerLayout)
            scroll.AutoScrollMinSize = new Size(maxTextWidth + 20, totalHeight + 10);

            group.EndUpdate();
            CustomizableControl.LayoutControl.EndUpdate();

            _logger.Trace("AddDummyFilesLinks2 end");
            return linksCount;
        }

        private int AddDummyFilesLinks(DevExpress.XtraLayout.LayoutControlGroup group, int linksCount)
        {
            _logger.Trace("AddDummyFilesLinks begin");

            if (group == null || linksCount <= 0)
                return 0;

            string fileName = "sefsfgsjgnsgnsegnsekrgnekrgnekgnekngwejrngeiwgnreiwgnwinwphgnptgnwe.docx";
            int idsLen = _testFilesIds.Length;

            // контейнер со скроллом, в нём будет второй layout
            var scroll = new DevExpress.XtraEditors.XtraScrollableControl();
            scroll.AutoScroll = true;
            scroll.Dock = DockStyle.Fill;

            var innerLayout = new DevExpress.XtraLayout.LayoutControl();
            innerLayout.Dock = DockStyle.Fill;
            innerLayout.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            innerLayout.Root.GroupBordersVisible = false;

            scroll.Controls.Add(innerLayout);

            CustomizableControl.LayoutControl.BeginUpdate();
            group.BeginUpdate();

            // кладём скроллируемый контейнер в группу одним item’ом
            var scrollItem = group.AddItem();
            scrollItem.Control = scroll;
            scrollItem.TextVisible = false;
            scrollItem.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            scrollItem.MinSize = new Size(10, 60);   // высота области ссылок
            scrollItem.MaxSize = new Size(0, 150);   // при переполнении появится скролл

            for (int i = 0; i < linksCount; i++)
            {
                var linkLabel = new DevExpress.XtraEditors.LabelControl();
                linkLabel.Text = i + "_" + fileName;
                linkLabel.AutoEllipsis = true;
                linkLabel.Appearance.ForeColor = Color.Blue;
                linkLabel.Appearance.Font = new Font(linkLabel.Appearance.Font, FontStyle.Underline);
                linkLabel.Cursor = Cursors.Hand;

                linkLabel.Tag = new FileLinkTag
                {
                    tagName = _fileLinkTag,
                    fileId = _testFilesIds[i % idsLen]
                };
                linkLabel.Click += FileLink_Click;

                var item = innerLayout.Root.AddItem();
                item.Control = linkLabel;
                item.TextVisible = false;
                item.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
                item.MinSize = new Size(10, 18);   // одна строка
                item.MaxSize = new Size(0, 18);    // фиксированная высота строки
            }

            group.EndUpdate();
            CustomizableControl.LayoutControl.EndUpdate();

            _logger.Trace("AddDummyFilesLinks end");
            return linksCount;
        }

        private int AddDummyFilesLinksAsLayoutControlItems(DevExpress.XtraLayout.LayoutControlGroup group, int linksCount)
        {
            _logger.Trace("AddDummyFilesLinksAsLayoutControlItems begin");

            if (group == null)
            {
                return 0;
            }

            CustomizableControl.LayoutControl.BeginUpdate();
            group.BeginUpdate();

            try
            {
                string fileName = "sefsfgsjgnsgnsegnsekrgnekrgnekgnekngwejrngeiwgnreiwgnwinwphgnptgnwe.docx";
                int idsLen = _testFilesIds.Length;

                for (int i = 0; i < linksCount; i++)
                {
                    // Создаем обычный LabelControl
                    LabelControl labelControl = new LabelControl();
                    labelControl.Text = i + "_" + fileName;
                    labelControl.Name = "FileLabel_" + i;

                    // Настраиваем как ссылку
                    labelControl.Appearance.ForeColor = Color.Blue;
                    labelControl.Appearance.Font = new Font(labelControl.Font, FontStyle.Underline);
                    labelControl.Appearance.Options.UseForeColor = true;
                    labelControl.Appearance.Options.UseFont = true;
                    labelControl.Appearance.TextOptions.HAlignment = HorzAlignment.Near;
                    labelControl.Appearance.TextOptions.WordWrap = WordWrap.Wrap;
                    labelControl.Appearance.TextOptions.Trimming = Trimming.EllipsisCharacter;

                    // Курсор в виде руки
                    labelControl.Cursor = Cursors.Hand;

                    // Tag с информацией о файле
                    labelControl.Tag = new FileLinkTag() { tagName = _fileLinkTag, fileId = _testFilesIds[i % idsLen] };

                    // Обработчик клика
                    labelControl.Click += FileLink_Click;

                    // Создаем LayoutControlItem
                    LayoutControlItem layoutItem = new LayoutControlItem();
                    layoutItem.Control = labelControl;
                    layoutItem.Name = "FileLayoutItem_" + i;
                    layoutItem.Text = ""; // Пустой текст
                    layoutItem.TextVisible = false; // Скрываем текст элемента

                    // Критически важные настройки
                    layoutItem.SizeConstraintsType = SizeConstraintsType.Default;
                    layoutItem.ControlAlignment = ContentAlignment.MiddleLeft;

                    // Добавляем в группу
                    group.Add(layoutItem);
                }

                // Настройка группы
                group.OptionsItemText.TextToControlDistance = 0;
                group.OptionsItemText.TextAlignMode = TextAlignModeGroup.AlignLocal;

                // Если много элементов - ограничиваем высоту группы
                if (linksCount > 10)
                {
                    //group.SizeConstraintsType = SizeConstraintsType.Custom;
                    //group.MaxSize = new Size(0, 300); // Максимальная высота
                }
            }
            finally
            {
                group.EndUpdate();
                CustomizableControl.LayoutControl.EndUpdate();

                // Принудительное обновление
                CustomizableControl.LayoutControl.Update();
                CustomizableControl.LayoutControl.PerformLayout();
            }

            _logger.Trace("AddDummyFilesLinksAsLayoutControlItems end. Added " + linksCount + " links");
            return linksCount;
        }

        private int AddDummyFilesLinksUsingEmptySpace(DevExpress.XtraLayout.LayoutControlGroup group, int linksCount)
        {
            _logger.Trace("AddDummyFilesLinksUsingEmptySpace begin");

            if (group == null)
            {
                return 0;
            }

            CustomizableControl.LayoutControl.BeginUpdate();
            group.BeginUpdate();

            try
            {
                // Очищаем группу полностью
                group.Clear();

                string fileName = "testfile.docx";
                int idsLen = _testFilesIds.Length;

                // Создаем контейнерную группу
                LayoutControlGroup containerGroup = new LayoutControlGroup();
                containerGroup.Name = "FilesContainer";
                containerGroup.TextVisible = false;
                containerGroup.GroupBordersVisible = false;

                // Настраиваем табличное расположение (если доступно)
                //if (group.OptionsTableLayoutGroup != null)
                //{
                //    containerGroup.OptionsTableLayoutGroup.Enabled = DefaultBoolean.True;
                //    containerGroup.OptionsTableLayoutGroup.Columns = 1;
                //}

                for (int i = 0; i < linksCount; i++)
                {
                    // Создаем EmptySpaceItem как основу
                    EmptySpaceItem emptySpace = new EmptySpaceItem();
                    emptySpace.Name = "FileSpace_" + i;
                    emptySpace.Size = new Size(0, 20); // Фиксированная высота

                    // Создаем LabelControl для текста
                    LabelControl label = new LabelControl();
                    label.Text = i + "_" + fileName;
                    label.Name = "FileText_" + i;
                    label.Appearance.ForeColor = Color.Blue;
                    label.Appearance.Font = new Font(label.Font, FontStyle.Underline);
                    label.Appearance.TextOptions.HAlignment = HorzAlignment.Near;
                    label.Cursor = Cursors.Hand;
                    label.Tag = new FileLinkTag() { tagName = _fileLinkTag, fileId = _testFilesIds[i % idsLen] };
                    label.Click += FileLink_Click;

                    // Размещаем Label поверх EmptySpaceItem
                    label.Location = new Point(emptySpace.Location.X + 2, emptySpace.Location.Y + 2);
                    _logger.Trace(emptySpace.Width.ToString());
                    label.Size = new Size(emptySpace.Width, 18);

                    // Добавляем в контейнер
                    containerGroup.Add(emptySpace);

                    // Добавляем Label в LayoutControl
                    CustomizableControl.LayoutControl.Controls.Add(label);
                }

                // Добавляем контейнер в основную группу
                group.Add(containerGroup);

                // Настраиваем размеры
                //containerGroup.BestFit();

                // Ограничиваем высоту, если много элементов
                if (linksCount > 15)
                {
                    //containerGroup.SizeConstraintsType = SizeConstraintsType.Custom;
                    //containerGroup.MaxSize = new Size(0, 400);
                }
            }
            finally
            {
                group.EndUpdate();
                CustomizableControl.LayoutControl.EndUpdate();
                CustomizableControl.LayoutControl.PerformLayout();
            }

            _logger.Trace("AddDummyFilesLinksUsingEmptySpace end");
            return linksCount;
        }

        private int AddDummyFilesLinksAsSimpleLabelItem(DevExpress.XtraLayout.LayoutControlGroup group, int linksCount)
        {
            _logger.Trace("AddDummyFilesLinksAsSimpleLabelItem begin");

            if (group == null)
            {
                return 0;
            }

            string fileName = "sefsfgsjgnsgnsegnsekrgnekrgnekgnekngwejrngeiwgnreiwgnwinwphgnptgnwe.docx";
            int idsLen = _testFilesIds.Length;

            CustomizableControl.LayoutControl.BeginUpdate();
            group.BeginUpdate();

            for (int i = 0; i < linksCount; i++)
            {
                var label = new DevExpress.XtraLayout.SimpleLabelItem();
                label.Tag = new FileLinkTag() { tagName = _fileLinkTag, fileId = _testFilesIds[i % idsLen] };

                label.Text = i + "_" + fileName;
                label.Name = label.Text;
                label.ControlName = label.Text;

                label.AppearanceItemCaption.ForeColor = Color.Blue;
                label.AppearanceItemCaption.Options.UseForeColor = true;
                label.AppearanceItemCaption.Font = new Font(label.AppearanceItemCaption.Font, FontStyle.Underline);

                label.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                label.AppearanceItemCaption.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                label.AppearanceItemCaption.TextOptions.WordWrap = DevExpress.Utils.WordWrap.Wrap;
                label.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Default;
                label.TextVisible = false;

                label.Padding = new DevExpress.XtraLayout.Utils.Padding(0);
                label.Spacing = new DevExpress.XtraLayout.Utils.Padding(0);
                label.AppearanceItemCaption.Options.UseTextOptions = true;
                label.AppearanceItemCaption.TextOptions.Trimming = DevExpress.Utils.Trimming.EllipsisCharacter;

                label.Click += FileLink_Click;

                group.Add(label);
            }

            group.OptionsItemText.TextToControlDistance = 0;
            group.OptionsItemText.TextAlignMode = TextAlignModeGroup.AlignLocal;

            group.EndUpdate();

            CustomizableControl.LayoutControl.EndUpdate();
            CustomizableControl.LayoutControl.Update();
            CustomizableControl.LayoutControl.PerformLayout();

            _logger.Trace("AddDummyFilesLinksAsSimpleLabelItem end");

            return linksCount;
        }

        #endregion

        private ITableControl TableDocumentsControl
        {
            get
            {
                return CustomizableControl.FindPropertyItem<ITableControl>("TableResult");
            }
        }

        #region Event Handlers

        // This event invokes twice, somehow
        private void Документ_CardActivated(System.Object sender, DocsVision.Platform.WinForms.CardActivatedEventArgs e)
        {
            _logger.Trace("Документ_CardActivated begin");



            _logger.Trace("Документ_CardActivated end");
        }

        private void Документ_CardClosing(System.Object sender, DocsVision.Platform.WinForms.CardClosingEventArgs e)
        {
            _logger.Trace("Документ_CardClosing begin");



            _logger.Trace("Документ_CardClosing end");
        }

        private void AddDummyFilesLinks_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            AddDummyFilesLinks6(FilesLinksGroup, _linksCount);
            //AddDummyFilesLinksAsSimpleLabelItem(FilesLinksGroup, _linksCount);
            //AddDummyFilesLinksUsingEmptySpace(FilesLinksGroup, _linksCount);
            //AddDummyFilesLinksAsLayoutControlItems(FilePreviewGroup, _linksCount);
            //AddDummyFilesLinksInFlowPanel(FilesLinksGroup, _linksCount);
            //AddDummyFilesLinksAsTable(FilesLinksGroup, _linksCount);
        }

        private void RemoveAllFilesLinks_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            RemoveAllFilesLinks();
        }

        private void FileLink_Click(object sender, EventArgs e)
        {
            if (sender is DevExpress.XtraLayout.SimpleLabelItem)
            {
                DevExpress.XtraLayout.SimpleLabelItem label = sender as DevExpress.XtraLayout.SimpleLabelItem;
                if (label.Tag == null)
                {
                    _logger.Error("Tag is null");
                    return;
                }

                if (label.Tag.GetType() == typeof(FileLinkTag))
                {
                    FileLinkTag tag = (FileLinkTag)label.Tag;
                    if (tag.tagName == _fileLinkTag)
                    {
                        OpenFilePreview(tag.fileId);
                    }
                }
                else
                {
                    _logger.Error("Label.Tag is not a type of FileLinkTag");
                    return;
                }
            }
            else
            {
                _logger.Error("Sender is not a type of DevExpress.XtraLayout.SimpleLabelItem");
            }
        }

        private void OpenFilePreview(Guid fileId)
        {
            try
            {
                var group = GetFirtstXLayoutItemWithText<LayoutControlGroup>(_filePreviewItemText);
                if (group != null)
                {
                    group.Visibility = LayoutVisibility.Always;
                    group.Expanded = true;

                    if (FilePreview != null)
                    {
                        FilePreview.Preview(fileId);
                        _logger.Debug("Preview activated. FileId:" + fileId);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Error opening file preview: " + ex.ToString());
            }
        }

        private void SaveLayout_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            CustomizableControl.LayoutControl.SaveLayoutToXml(_pathToTestLayoutXml);
        }

        private void SearchButton_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("SearchButton_ItemClick begin");



            _logger.Trace("SearchButton_ItemClick end");
        }

        private void ExportToExcelButton_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("ExportToExcelButton_ItemClick begin");



            _logger.Trace("ExportToExcelButton_ItemClick end");
        }

        private void LogResponseStr_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("LogResponseStr_ItemClick begin");



            _logger.Trace("LogResponseStr_ItemClick end");
        }

        private void LogSelectParams_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("LogSelectParams_ItemClick begin");



            _logger.Trace("LogSelectParams_ItemClick end");
        }

        /// <summary>
        /// UI elements debug info
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LogUI_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("LogUI_ItemClick begin");

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
                _logger.Log("UI", "foreach count:" + foreachCount);

                if (items.Count != items.ItemCount)
                {
                    _logger.Log("UI", string.Format("Items collection's counts are not equal. items.Count={0} items.ItemCount={1}", items.Count, items.ItemCount));
                }

                _logger.Log("UI", "Items.GroupCount=" + items.GroupCount);

                int n = items.Count;
                for (int i = 0; i < n; i++)
                {
                    _logger.Log("UI", string.Format("item[{0}] type: {1}", i, items[i].GetType().FullName));

                    if (items[i] is BaseLayoutItem)
                    {
                        BaseLayoutItem baseLayoutItem = (BaseLayoutItem)items[i];
                        _logger.Log("UI", string.Format("item[{0}] name:{1}; parentName:{2}; typeName:{3}, text:{4}, visible:{5}", 
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
                _logger.Error("LogUI_ItemClick(): " + ex.ToString());
            }

            _logger.Trace("LogUI_ItemClick end");
        }

        [Obsolete]
        private DevExpress.XtraLayout.LayoutControlGroup _filesLinksGroup = null;
        [Obsolete]
        private void AddDummyLinksAsSimpleButton()
        {
            _logger.Trace("AddDummyLinksAsSimpleButton begin");

            try
            {
                var filesRootGroup = FilesLinksGroup;
                if (filesRootGroup == null)
                {
                    _logger.Warn("Files root group Файлы doesn't exist");
                    return;
                }

                //RemoveAllLayoutItems("FilesLinksGroup");

                // Создаем группу для кнопок
                if (_filesLinksGroup != null)
                {
                    _logger.Warn("Files links group already exists");
                    return;
                }

                CustomizableControl.LayoutControl.BeginUpdate();

                _filesLinksGroup = new DevExpress.XtraLayout.LayoutControlGroup();
                _filesLinksGroup.Name = "FilesLinksGroup";
                _filesLinksGroup.GroupBordersVisible = true;
                _filesLinksGroup.TextVisible = false;

                //group.AddItem(buttonGroup, DevExpress.XtraLayout.Utils.InsertType.Top);
                filesRootGroup.Add(_filesLinksGroup);

                SimpleButton fileButton1 = new SimpleButton();
                SimpleButton fileButton2 = new SimpleButton();
                SimpleButton fileButton3 = new SimpleButton();
                SimpleButton fileButton4 = new SimpleButton();
                SimpleButton fileButton5 = new SimpleButton();
                SimpleButton fileButton6 = new SimpleButton();
                _filesLinksGroup.AddItem("File1", fileButton2).Name = "File1";
                _filesLinksGroup.AddItem("File2", fileButton1).Name = "File2";
                _filesLinksGroup.AddItem("File3", fileButton3).Name = "File3";
                _filesLinksGroup.AddItem("File4", fileButton4).Name = "File4";
                _filesLinksGroup.AddItem("File5", fileButton5).Name = "File5";
                _filesLinksGroup.AddItem("File6", fileButton6).Name = "File6";

                // НАХОДИМ первый элемент (FilePreviewItem) и перемещаем buttonGroup НАВЕРХ
                //if (_filesLinksGroup.Items.Count > 1)
                //{
                //    BaseLayoutItem filepreview = GetLayoutItem<BaseLayoutItem>("filepreview");
                //    if (filepreview != null)
                //    {
                //        _filesLinksGroup.Move(filepreview, DevExpress.XtraLayout.Utils.InsertType.Top);

                //        _logger.Trace("FilePreview should be moved up");
                //    }
                //}

                // ВСТАВЛЯЕМ ПЕРВЫМ (над FilePreview)
                //var wrapper = filesGroup.AddItem(); // Placeholder сверху
                //wrapper.TextVisible = false;
                //wrapper.Replace(_filesLinksGroup);  // ✅ Заменяем на группу

                CustomizableControl.LayoutControl.EndUpdate();
                filesRootGroup.Update();
            }
            catch (Exception ex)
            {
                _logger.Error("AddDummyLinksAsSimpleButton() " + ex.ToString());
            }
            finally
            {
                _logger.Trace("AddDummyLinksAsSimpleButton end");
            }
        }

        private void TestActionOnGroup_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("TestActionOnGroup_ItemClick begin");

            

            _logger.Trace("TestActionOnGroup_ItemClick end");
        }

        private void RemoveAllFilesLinks()
        {
            _logger.Trace("RemoveAllFilesLinks begin");

            DevExpress.XtraLayout.LayoutControlGroup filesLinksGroup;
            try
            {
                filesLinksGroup = FilesLinksGroup;
            }
            catch (Exception ex)
            {
                _logger.Error("Unnable to find DevExpress.XtraLayout.LayoutControlGroup: " + ex.ToString());
                return;
            }

            if (filesLinksGroup.Items == null)
            {
                _logger.Trace("filesLinksGroup.Items == null");
                return;
            }

            int n = filesLinksGroup.Items.Count;
            _logger.Trace("Items to remove: " + n);

            filesLinksGroup.BeginUpdate();
            //for (int i = n - 1; i >= 0; i--)
            //{
            //    filesLinksGroup.RemoveAt(i);
            //}
            filesLinksGroup.Clear();

            //filesLinksGroup.OptionsTableLayoutGroup.RowDefinitions.Clear();
            //filesLinksGroup.OptionsTableLayoutGroup.ColumnDefinitions.Clear();
            //filesLinksGroup.MinSize = new Size(filesLinksGroup.MinSize.Width, 0);
            //filesLinksGroup.Add(new DevExpress.XtraLayout.EmptySpaceItem());
            //filesLinksGroup.BestFit();
            filesLinksGroup.EndUpdate();

            //if (CustomizableControl.LayoutControl != null)
            //{
            //    CustomizableControl.LayoutControl.PerformLayout();
            //}

            filesLinksGroup.Expanded = false;

            _logger.Trace("RemoveAllFilesLinks end");
        }

        private void RemoveFilesLinksGroup()
        {
            _logger.Trace("RemoveFilesLinksGroup begin");

            if (_filesLinksGroup != null && _filesLinksGroup.Items != null)
            {
                int n = _filesLinksGroup.Items.Count;
                if (n == 0)
                {
                    _logger.Warn("FilesLinksGroup items collection count=0");
                }
                else
                {
                    CustomizableControl.LayoutControl.BeginUpdate();
                    for (int i = n - 1; i >= 0; i--)
                    {
                        _filesLinksGroup.RemoveAt(i);
                    }
                    if (_filesLinksGroup.Parent != null)
                    {
                        _filesLinksGroup.Parent.Remove(_filesLinksGroup);
                    }
                    else
                    {
                        _logger.Warn("FilesLinksGroup Parent is null");
                    }

                    _filesLinksGroup = null;

                    CustomizableControl.LayoutControl.EndUpdate();
                }
            }
            else
            {
                _logger.Warn("FilesLinksGroup or it's items collection is null");
            }

            _logger.Trace("RemoveFilesLinksGroup end");
        }

        [Obsolete]
        private int _cleanLayoutItemCount;
        [Obsolete]
        private void ClearLayoutAggressive()
        {
            RemoveFilesLinksGroup();

            CustomizableControl.LayoutControl.BeginUpdate();
            var items = CustomizableControl.LayoutControl.Items;
            int currentCount = items.Count;
            _logger.Trace("CURRENT COUNT:" + currentCount);
            while (currentCount > _cleanLayoutItemCount)
            {
                var orphan = items[currentCount - 1];
                if (orphan.Parent != null)
                    orphan.Parent.Remove(orphan);
                else
                    CustomizableControl.LayoutControl.Root.Remove(orphan);

                currentCount--;
            }
            CustomizableControl.LayoutControl.EndUpdate();

            _logger.Trace("AFTER CLEAR COUNT:" + currentCount);
        }

        private void RestoreLayout_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("RestoreLayout_ItemClick begin");

            CustomizableControl.LayoutControl.RestoreLayoutFromXml(_pathToTestLayoutXml);

            _logger.Trace("RestoreLayout_ItemClick end");
        }

        private int AddDummyFilesLinksInFlowPanel(DevExpress.XtraLayout.LayoutControlGroup group, int linksCount)
        {
            _logger.Trace("AddDummyFilesLinksInFlowPanel begin");

            if (group == null)
            {
                return 0;
            }

            group.BeginUpdate();

            // Очищаем группу
            group.Clear();

            // Создаем обычную Panel с AutoScroll
            Panel panel = new Panel();
            panel.Name = "FilesPanel";
            panel.AutoScroll = true;
            panel.AutoSize = false;
            panel.Dock = DockStyle.Fill;

            // Рассчитываем высоту панели
            int itemHeight = 22;
            int spacing = 2;
            int panelHeight = Math.Min(linksCount * (itemHeight + spacing), 300); // Максимальная высота 300px

            panel.Height = panelHeight;
            //panel.Width = group.Width - 20; // Учитываем отступы

            string fileName = "testfile.docx";
            int idsLen = _testFilesIds.Length;

            int currentTop = 0;
            for (int i = 0; i < linksCount; i++)
            {
                LinkLabel linkLabel = new LinkLabel();
                linkLabel.Text = i + "_" + fileName;
                linkLabel.Name = "FileLink_" + i;
                linkLabel.AutoSize = false;
                //linkLabel.Width = panel.Width - 10; // Ширина с учетом отступов
                linkLabel.Height = itemHeight;
                linkLabel.Top = currentTop;
                linkLabel.Left = 5;
                linkLabel.Tag = new FileLinkTag() { tagName = _fileLinkTag, fileId = _testFilesIds[i % idsLen] };
                linkLabel.LinkClicked += LinkLabel_LinkClicked;

                panel.Controls.Add(linkLabel);
                currentTop += itemHeight + spacing;
            }

            // Создаем LayoutControlItem для Panel
            LayoutControlItem panelItem = new LayoutControlItem();
            panelItem.Control = panel;
            panelItem.Name = "FilesPanelItem";
            panelItem.Text = "";
            panelItem.TextVisible = false;
            panelItem.SizeConstraintsType = SizeConstraintsType.Custom;
            panelItem.MaxSize = new Size(0, panelHeight);
            panelItem.MinSize = new Size(0, panelHeight);
            panelItem.Height = panelHeight;

            group.Add(panelItem);

            group.EndUpdate();

            // Принудительное обновление
            group.Update();
            if (CustomizableControl.LayoutControl != null)
            {
                CustomizableControl.LayoutControl.Update();
                CustomizableControl.LayoutControl.PerformLayout();
            }

            _logger.Trace("AddDummyFilesLinksInFlowPanel end");
            return linksCount;
        }

        private int AddDummyFilesLinksAsTable(DevExpress.XtraLayout.LayoutControlGroup group, int linksCount)
        {
            _logger.Trace("AddDummyFilesLinksAsTable begin");

            if (group == null)
            {
                return 0;
            }

            group.BeginUpdate();

            // Очищаем группу
            group.Clear();

            // Создаем TableLayoutPanel
            TableLayoutPanel tablePanel = new TableLayoutPanel();
            tablePanel.Name = "FilesTablePanel";
            tablePanel.ColumnCount = 1;
            tablePanel.RowCount = linksCount;
            tablePanel.AutoSize = true;
            tablePanel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            tablePanel.Dock = DockStyle.Top; // Изменено с Fill на Top

            // Настраиваем строки
            for (int i = 0; i < linksCount; i++)
            {
                tablePanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 22));
            }

            string fileName = "testfiletestfiletestfiletestfiletestfiletestfiletestfiletestfiletestfile.docx";
            int idsLen = _testFilesIds.Length;

            for (int i = 0; i < linksCount; i++)
            {
                LinkLabel linkLabel = new LinkLabel();
                linkLabel.Text = i + "_" + fileName;
                linkLabel.Name = "FileLink_" + i;
                linkLabel.AutoSize = true;
                linkLabel.Dock = DockStyle.Fill;
                linkLabel.Tag = new FileLinkTag() { tagName = _fileLinkTag, fileId = _testFilesIds[i % idsLen] };
                linkLabel.LinkClicked += LinkLabel_LinkClicked;

                tablePanel.Controls.Add(linkLabel, 0, i);
            }

            // Рассчитываем высоту таблицы
            int tableHeight = linksCount * 22 + 5; // +5 для отступов

            // Создаем LayoutControlItem для TableLayoutPanel
            LayoutControlItem tableItem = new LayoutControlItem();
            tableItem.Control = tablePanel;
            tableItem.Name = "FilesTableItem";
            tableItem.Text = "";
            tableItem.TextVisible = false;

            // Критически важные настройки размера:
            tableItem.SizeConstraintsType = SizeConstraintsType.Custom;
            tableItem.MaxSize = new Size(group.Width > 0 ? group.Width : 300, tableHeight);
            tableItem.MinSize = new Size(100, tableHeight); // Минимальная ширина 100

            // Устанавливаем размер панели
            tablePanel.Width = group.Width > 0 ? group.Width - 10 : 290;
            tablePanel.Height = tableHeight;

            group.Add(tableItem);

            group.EndUpdate();

            // Принудительно обновляем макет
            if (CustomizableControl.LayoutControl != null)
            {
                CustomizableControl.LayoutControl.Update();
                CustomizableControl.LayoutControl.PerformLayout();
            }

            _logger.Trace("AddDummyFilesLinksAsTable end");
            return linksCount;
        }

        private void LinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            
        }

        private void FillMainTable_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            
        }

        private void ClearMainTable_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            
        }

        #endregion

        #region External

        private string GetTraceInfo(BaseLayoutItem item)
        {
            return string.Format("name:{0}; parentName:{1}; typeName:{2}, text:{3}, textVisible:{4}, visible:{5}, tag:{6}",
                item.Name, item.ParentName, item.TypeName, item.Text, item.TextVisible, item.Visible, item.Tag);
        }

        private TLayout GetFirtstXLayoutItemWithText<TLayout>(string text, bool visible) where TLayout : DevExpress.XtraLayout.BaseLayoutItem
        {
            var items = CustomizableControl.LayoutControl.Items;
            if (items == null)
            {
                throw new Exception("ICustomizableControl.LayoutControl.Items is null");
            }

            int n = items.Count;
            TLayout layoutItem = null;
            for (int i = 0; i < n; i++)
            {
                if (items[i] is TLayout)
                {
                    TLayout castedItem = (TLayout)items[i];
                    if (castedItem.Visible == visible && castedItem.Text == text)
                    {
                        layoutItem = castedItem;
                        break;
                    }
                }
            }

            if (layoutItem == null)
            {
                throw new Exception((typeof(TLayout)) + " wasn't found with visible=true and text=" + text);
            }

            return layoutItem;
        }

        private TLayout GetFirtstXLayoutItemWithText<TLayout>(string text) where TLayout : DevExpress.XtraLayout.BaseLayoutItem
        {
            var items = CustomizableControl.LayoutControl.Items;
            if (items == null)
            {
                throw new Exception("ICustomizableControl.LayoutControl.Items is null");
            }

            int n = items.Count;
            TLayout layoutItem = null;
            for (int i = 0; i < n; i++)
            {
                if (items[i] is TLayout)
                {
                    TLayout castedItem = (TLayout)items[i];
                    if (castedItem.Text == text)
                    {
                        layoutItem = castedItem;
                        break;
                    }
                }
            }

            if (layoutItem == null)
            {
                throw new Exception((typeof(TLayout)) + " wasn't found with text=" + text);
            }

            return layoutItem;
        }

        private TLayout GetFirstXLayoutItem<TLayout>(string name) where TLayout : DevExpress.XtraLayout.BaseLayoutItem
        {
            var items = CustomizableControl.LayoutControl.Items;
            if (items == null)
            {
                throw new Exception("ICustomizableControl.LayoutControl.Items is null");
            }

            int n = items.Count;
            TLayout layoutItem = null;
            for (int i = 0; i < n; i++)
            {
                if (items[i] is TLayout)
                {
                    TLayout castedItem = (TLayout)items[i];
                    if (castedItem.Name == name)
                    {
                        layoutItem = castedItem;
                        break;
                    }
                }
            }

            if (layoutItem == null)
            {
                throw new Exception((typeof(TLayout)) + " wasn't found with name=" + name);
            }

            return layoutItem;
        }

        [Obsolete]
        private void RemoveAllLayoutItems(string name)
        {
            _logger.Trace("RemoveAllLayoutItems begin. Remove with name:" + name);

            DevExpress.XtraLayout.LayoutControl root = CustomizableControl.LayoutControl;

            var items = root.Items;
            if (items == null)
            {
                _logger.Warn("Layout items collection is null");
                return;
            }

            int n = items.Count;
            for (int i = n - 1; i >= 0; i--)
            {
                if (items[i] is BaseLayoutItem)
                {
                    BaseLayoutItem baseLayoutItem = (BaseLayoutItem)items[i];
                    if (baseLayoutItem.Name == name)
                    {
                        if (baseLayoutItem.Parent != null)
                        {
                            _logger.Trace("Removing item from parent: " + GetTraceInfo(baseLayoutItem.Parent));
                            _logger.Trace("Item to be removed: " + GetTraceInfo(baseLayoutItem));
                            baseLayoutItem.Parent.Remove(baseLayoutItem);
                        }
                        else
                        {
                            _logger.Trace("Removing item from ROOT: " + GetTraceInfo(root.Root));
                            _logger.Trace("Item to be removed: " + GetTraceInfo(baseLayoutItem));
                            root.Root.Remove(baseLayoutItem);
                        }
                    }
                }
                else
                {
                    _logger.Warn(string.Format("item[{0}] is not assignable to type {1}", i, typeof(BaseLayoutItem)));
                }
            }

            root.Refresh(); // Обновляем макет

            _logger.Trace("RemoveAllLayoutItems end");
        }

        #endregion

        private int _linksCount = 50;
    }
}