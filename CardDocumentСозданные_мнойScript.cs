using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using MLog;

using DevExpress.XtraEditors;
using DevExpress.XtraLayout;

using DocsVision.BackOffice.ObjectModel;
using DocsVision.BackOffice.WinForms;
using DocsVision.BackOffice.WinForms.Controls;
using DocsVision.BackOffice.WinForms.Design.LayoutItems;
using DocsVision.BackOffice.WinForms.Design.PropertyControls;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectModel;
using System.Collections.Generic;

namespace BackOffice
{
    public class CardDocumentСозданные_мнойScript : CardDocumentScript
    {

        #region Properties

        private Logger _logger;
        private readonly Guid _testCardId = new Guid("B8BEE5B4-FCD5-F011-AFF0-000C290CEAEE");

        private ICustomizableControl Customizable
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
                return Customizable.FindLayoutItem("FilePreview");
            }
        }


        private ILayoutPropertyItem FilePreviewAsLayoutPropertyItem
        {
            get
            {
                return Customizable.FindPropertyItem<ILayoutPropertyItem>("FilePreview");
            }
        }

        private IPreviewFileControl FilePreview
        {
            get
            {
                return Customizable.FindPropertyItem<IPreviewFileControl>("FilePreview");
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

            //_logger.Info("FilePreview type: " + FilePreviewAsObj.GetType().FullName);
            //_logger.Info("FilePreview's LayoutPropertyType: " + FilePreviewAsLayoutPropertyItem.PropertyType);
            _logger.Info("FilePreview is null?" + (FilePreview == null));

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

        // TO DO: additional dll
        private TLayout GetLayoutItemWithText<TLayout>(string text) where TLayout : DevExpress.XtraLayout.BaseLayoutItem
        {
            var items = Customizable.LayoutControl.Items;
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
                    if (castedItem.Visible && castedItem.Text == text)
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

        // TO DO: additional dll
        private TLayout GetLayoutItem<TLayout>(string name) where TLayout : DevExpress.XtraLayout.BaseLayoutItem
        {
            var items = Customizable.LayoutControl.Items;
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

        #endregion

        #region Event Handlers

        // This event invokes twice, somehow
        private void Документ_CardActivated(System.Object sender, DocsVision.Platform.WinForms.CardActivatedEventArgs e)
        {
            _logger.Trace("Документ_CardActivated begin");

            _cleanLayoutItemCount = 19;
            _logger.Trace("COUNT: " + Customizable.LayoutControl.Items.Count);

            //RemoveFilesLinksGroup();
            ClearLayoutAggressive();

            _logger.Trace("Документ_CardActivated end");
        }

        private void Документ_CardClosing(System.Object sender, DocsVision.Platform.WinForms.CardClosingEventArgs e)
        {
            _logger.Trace("Документ_CardClosing begin");

            if (_filesLinksGroup != null)
            {
                //RemoveAllLayoutItems("FilesLinksGroup");
                //RemoveFilesLinksGroup();
            }
            ClearLayoutAggressive();

            _logger.Trace("COUNT: " + Customizable.LayoutControl.Items.Count);

            _logger.Trace("Документ_CardClosing end");
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

        // TO DO: additional dll
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
                var items = Customizable.LayoutControl.Items;
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

        private DevExpress.XtraLayout.LayoutControlGroup _filesLinksGroup = null;

        private void TestActionOnGroup_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("TestActionOnGroup_ItemClick begin");

            try
            {
                var filesRootGroup = GetLayoutItemWithText<DevExpress.XtraLayout.LayoutControlGroup>("Файлы");
                //var filesRootGroup = GetLayoutItem<DevExpress.XtraLayout.LayoutControlGroup>("Файлы");
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

                Customizable.LayoutControl.BeginUpdate();

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

                Customizable.LayoutControl.EndUpdate();
                filesRootGroup.Update();
            }
            catch (Exception ex)
            {
                _logger.Error("TestActionOnGroup_ItemClick() " + ex.ToString());
            }

            _logger.Trace("TestActionOnGroup_ItemClick end");
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
                    Customizable.LayoutControl.BeginUpdate();
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

                    Customizable.LayoutControl.EndUpdate();
                }
            }
            else
            {
                _logger.Warn("FilesLinksGroup or it's items collection is null");
            }

            _logger.Trace("RemoveFilesLinksGroup end");
        }

        private int _cleanLayoutItemCount;
        private void ClearLayoutAggressive()
        {
            RemoveFilesLinksGroup();

            Customizable.LayoutControl.BeginUpdate();
            var items = Customizable.LayoutControl.Items;
            int currentCount = items.Count;
            _logger.Trace("CURRENT COUNT:" + currentCount);
            while (currentCount > _cleanLayoutItemCount)
            {
                var orphan = items[currentCount - 1];
                if (orphan.Parent != null)
                    orphan.Parent.Remove(orphan);
                else
                    Customizable.LayoutControl.Root.Remove(orphan);

                currentCount--;
            }
            Customizable.LayoutControl.EndUpdate();

            _logger.Trace("AFTER CLEAR COUNT:" + currentCount);
        }

        [Obsolete]
        // TO DO: additional dll
        private void RemoveAllLayoutItems(string name)
        {
            _logger.Trace("RemoveAllLayoutItems begin. Remove with name:" + name);

            DevExpress.XtraLayout.LayoutControl root = Customizable.LayoutControl;

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

        private void RestoreLayout_ItemClick(System.Object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            _logger.Trace("RestoreLayout_ItemClick begin");

            RemoveFilesLinksGroup();
            ClearLayoutAggressive();

            _logger.Trace("RestoreLayout_ItemClick end");
        }

        private string GetTraceInfo(BaseLayoutItem item)
        {
            return string.Format("name:{0}; parentName:{1}; typeName:{2}, text:{3}, textVisible:{4}, visible:{5}, tag:{6}",
                item.Name, item.ParentName, item.TypeName, item.Text, item.TextVisible, item.Visible, item.Tag);
        }

        /// <summary>
        /// Вызывается, когда элемент меняет состояние с "показан" на "скрыт" или наоборот. Работает при сворачивании/разв. группы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //private void FilePreview_VisibleChanged(System.Object sender, System.EventArgs e)
        //{
        //    _logger.Trace("FilePreview_VisibleChanged begin");



        //    _logger.Trace("FilePreview_VisibleChanged end");
        //}

        #endregion

    }
}