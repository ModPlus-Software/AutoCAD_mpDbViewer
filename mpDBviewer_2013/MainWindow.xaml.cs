using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Xml.Linq;
using mpBaseInt;
using ModPlusAPI;
using ModPlusAPI.Windows.Helpers;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace mpDbViewer
{
    public partial class MpDbviewerWindow
    {
        private const string LangItem = "mpDBviewer";
        // Текущая коллекция документов (база данных) с которой работаю
        private ICollection<BaseDocument> _currentDocumentCollection;
        // Текущий выбранный документ
        private BaseDocument _currentDocument;
        public MpDbviewerWindow()
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetItem(LangItem, "h1");
        }

        private void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            SizeToContent = SizeToContent.Manual;
            // ОБЯЗАТЕЛЬНО!
            // "Загрузка" документов баз данных
            mpMetall.Metall.LoadAllDocument();
            mpConcrete.Concrete.LoadAllDocument();
            mpWood.Wood.LoadAllDocument();
            mpMaterial.Material.LoadAllDocument();
            mpOther.Other.LoadAllDocument();
            // Марки стали - документы
            CbSteelDocument.ItemsSource = SteelDocuments.GetSteels();
            CbSteelDocument.SelectedIndex = 0;
            // Загружаем значение из файла настроек
            var selectedDb = UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "mpDBviewer", "selectedDB");
            if (!string.IsNullOrEmpty(selectedDb))
            {
                if (int.TryParse(selectedDb, out int index))
                    CbDataBases.SelectedIndex = index;
            }
            var selectedGroup = UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "mpDBviewer", "selectedGroup");
            if (!string.IsNullOrEmpty(selectedGroup))
            {
                if (int.TryParse(selectedGroup, out int index))
                {
                    try
                    {
                        LbGroups.SelectedIndex = index;
                    }
                    catch
                    {
                        //ignored
                    }
                }
            }
        }
        // Очистка всех контроллов
        private void ClearAllControls()
        {
            LbGroups.ItemsSource = null;
            ClearAllExceptGroups();
        }
        // Очистка всех, кроме списка групп
        private void ClearAllExceptGroups()
        {
            DgItems.Columns.Clear();
            DgItems.ItemsSource = null;
            LbDocuments.ItemsSource = null;
            BtShowImage.IsEnabled = false;
            TabItemExport.Visibility = Visibility.Collapsed;
            TbDocumentName.Text = string.Empty;
            // 
            TabControlDetail.SelectedIndex = 0;
            StkSteel.Visibility = Visibility.Collapsed;
            LbDocumentTypes.ItemsSource = null;
            StkNaim.Visibility = Visibility.Collapsed;
            TvDocuments.ItemsSource = null;
            BtShowAsTree.Visibility = Visibility.Collapsed;
            BtShowAsList.Visibility = Visibility.Collapsed;
            _currentDocument = null;
            TbNaimFirst.Text = string.Empty;
            TbNaimSecond.Text = string.Empty;
        }
        // Выбор базы данных
        private void CbDataBases_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ClearAllControls();
            if (sender is ComboBox comboBox)
            {
                switch (comboBox.SelectedIndex)
                {
                    case 0:
                        _currentDocumentCollection = mpMetall.Metall.DocumentCollection;
                        break;
                    case 1:
                        _currentDocumentCollection = mpConcrete.Concrete.DocumentCollection;
                        break;
                    case 2:
                        _currentDocumentCollection = mpWood.Wood.DocumentCollection;
                        break;
                    case 3:
                        _currentDocumentCollection = mpMaterial.Material.DocumentCollection;
                        break;
                    case 4:
                        _currentDocumentCollection = mpOther.Other.DocumentCollection;
                        break;
                    default:
                        _currentDocumentCollection = null;
                        break;
                }
            }
            if (_currentDocumentCollection != null)
                LbGroups.ItemsSource = _currentDocumentCollection.Select(x => x.Group).ToList().Distinct();
        }
        // Выбор группы
        private void LbGroups_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ClearAllExceptGroups();
            var listbox = sender as ListBox;
            if (listbox?.ItemsSource != null)
            {
                var selectedGroup = listbox.SelectedItem.ToString();
                if (!string.IsNullOrEmpty(selectedGroup))
                {
                    // Если для указанной группы значений ShortName больше 1, то отображаем в виде дерева
                    // если всего 1, то в виде списка
                    // Но заполняем и то, и другое, чтобы можно было переключаться
                    var lst = _currentDocumentCollection.Where(x => x.Group.Equals(selectedGroup)).ToList();
                    lst.Sort(new DocumentsSortComparer());
                    LbDocuments.ItemsSource = lst;

                    var hlst = new ObservableCollection<GroupByShortName>();

                    foreach (var srtName in lst.Select(x => x.ShortName).Distinct())
                    {
                        var shortNames = new GroupByShortName { ShortName = srtName };
                        var name = srtName;
                        foreach (var document in lst.Where(document => document.ShortName.Equals(name)))
                        {
                            shortNames.Documents.Add(document);
                        }
                        hlst.Add(shortNames);
                    }

                    TvDocuments.ItemsSource = hlst;

                    if (hlst.Count > 1)
                    {
                        TvDocuments.Visibility = Visibility.Visible;
                        LbDocuments.Visibility = Visibility.Collapsed;
                        BtShowAsList.Visibility = Visibility.Visible;
                        BtShowAsTree.Visibility = Visibility.Collapsed;
                    }
                    else
                    {
                        TvDocuments.Visibility = Visibility.Collapsed;
                        LbDocuments.Visibility = Visibility.Visible;
                        BtShowAsList.Visibility = Visibility.Collapsed;
                        BtShowAsTree.Visibility = Visibility.Visible;
                    }
                }
            }
        }
        private class GroupByShortName
        {
            public GroupByShortName()
            {
                Documents = new ObservableCollection<BaseDocument>();
            }
            public string ShortName { get; set; }
            public ObservableCollection<BaseDocument> Documents { get; private set; }
        }
        // Компаратор для сортировки списка документов
        private class DocumentsSortComparer : Comparer<BaseDocument>
        {
            public override int Compare(BaseDocument x, BaseDocument y)
            {
                if (x != null)
                    if (y != null)
                        return String.Compare(
                            x.DocumentShortName + " (" + x.DocumentType + " " + x.DocumentNumber + ")",
                            y.DocumentShortName + " (" + y.DocumentType + " " + y.DocumentNumber + ")",
                            StringComparison.Ordinal);
                return 0;
            }
        }
        // Выбор документа в ListBox
        private void LbDocuments_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var listbox = sender as ListBox;
            if (listbox?.SelectedItem is BaseDocument docInBase) FillDataGridWithItems(docInBase);
        }
        // Выбор документа в TreeView
        private void TvDocuments_OnSelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var treeView = sender as TreeView;
            var tvi = treeView?.SelectedItem;
            if (tvi is BaseDocument) FillDataGridWithItems(tvi as BaseDocument);
        }

        private void FillDataGridWithItems(BaseDocument baseDocument)
        {
            DgItems.Columns.Clear();
            DgItems.ItemsSource = null;
            // Текущий выбранный документ
            _currentDocument = baseDocument;
            // Заполняем вкладку "Свойства"
            foreach (var itemType in baseDocument.ItemTypes)
            {
                itemType.SelectedItem = itemType.TypeValues[0];
            }
            LbDocumentTypes.ItemsSource = baseDocument.ItemTypes;

            if (baseDocument.ItemTypes.Any())
                LbDocumentTypes.Visibility = Visibility.Visible;
            else LbDocumentTypes.Visibility = Visibility.Collapsed;

            if (!(baseDocument.SymbolCount == 0 | baseDocument.Symbols == null))
            {
                for (var i = 1; i <= baseDocument.SymbolCount; i++)
                {
                    var dgc = new DataGridTextColumn
                    {
                        IsReadOnly = true,
                        Header = baseDocument.Symbols.ElementAt(i - 1),
                        Binding = new Binding("Attribute[Prop" + i.ToString(CultureInfo.InvariantCulture) + "].Value")
                    };
                    DgItems.Columns.Add(dgc);
                }
                DgItems.ItemsSource = baseDocument.Items.Elements("Item");
                // Сразу выберем первый элемент в списке, чтобы отобразилось условное обозначение
                DgItems.SelectedIndex = 0;
            }
            else ShowItemNaim(); // Так как при отсутствии таблицы метода не сработает автоматически
            // Если SymbolCount = 0, значит таблицы нет. Тогда сразу открываем вторую вкладку
            // если вторая вкладка была открыта уже, то не меняем ничего (для удобства пользования)
            if (TabControlDetail.SelectedIndex != 1)
                TabControlDetail.SelectedIndex = baseDocument.SymbolCount == 0 ? 1 : 0;

            // Активация кнопки "изображение"
            BtShowImage.IsEnabled = !string.IsNullOrEmpty(baseDocument.Image);
            // Название
            TbDocumentName.DataContext = baseDocument;
            // Видимость кнопок и прочего
            TabItemExport.Visibility = Visibility.Visible;
            // Если есть сталь
            StkSteel.Visibility = baseDocument.HasSteel ? Visibility.Visible : Visibility.Collapsed;
            NaimSplitter.Visibility = baseDocument.HasSteel ? Visibility.Visible : Visibility.Collapsed;
            TbNaimSecond.Visibility = baseDocument.HasSteel ? Visibility.Visible : Visibility.Collapsed;
            // Включаем отображение "Наименовани"
            StkNaim.Visibility = Visibility.Visible;
        }

        private void BtShowImage_OnClick(object sender, RoutedEventArgs e)
        {
            BaseDocument docInBase = null;
            if (LbDocuments.Visibility == Visibility.Visible)
                docInBase = LbDocuments.SelectedItem as BaseDocument;
            if (TvDocuments.Visibility == Visibility.Visible)
                docInBase = TvDocuments.SelectedItem as BaseDocument;
            if (docInBase != null)
            {
                ShowBaseElementImage win;
                switch (docInBase.DataBaseName)
                {
                    case "DbMetall":
                        win = new ShowBaseElementImage { Img = { Source = new BitmapImage(new Uri(mpMetall.Metall.GetImagePath(docInBase), UriKind.Absolute)) } };
                        win.ShowDialog();
                        break;
                    case "DbConcrete":
                        win = new ShowBaseElementImage { Img = { Source = new BitmapImage(new Uri(mpConcrete.Concrete.GetImagePath(docInBase), UriKind.Absolute)) } };
                        win.ShowDialog();
                        break;
                    case "DbWood":
                        win = new ShowBaseElementImage { Img = { Source = new BitmapImage(new Uri(mpWood.Wood.GetImagePath(docInBase), UriKind.Absolute)) } };
                        win.ShowDialog();
                        break;
                    case "DbMaterial":
                        win = new ShowBaseElementImage { Img = { Source = new BitmapImage(new Uri(mpMaterial.Material.GetImagePath(docInBase), UriKind.Absolute)) } };
                        win.ShowDialog();
                        break;
                    case "DbOther":
                        win = new ShowBaseElementImage { Img = { Source = new BitmapImage(new Uri(mpOther.Other.GetImagePath(docInBase), UriKind.Absolute)) } };
                        win.ShowDialog();
                        break;
                }
            }

        }
        #region Search
        // Ввод текста в строке поиска
        private void TbSearchTxt_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            // Пусть работает от трех символов и выше
            if (sender is TextBox textbox)
            {
                if (textbox.Text.Length >= 2)
                    SearchDocumentsInDb(textbox.Text);
                else LbSearchResults.ItemsSource = null;
            }
        }

        private void SearchDocumentsInDb(string stringForSearch)
        {
            LbSearchResults.ItemsSource = null;

            // Список доступных коллекций
            var documentCollections = new List<ICollection<BaseDocument>>
            {
                mpMetall.Metall.DocumentCollection,
                mpConcrete.Concrete.DocumentCollection,
                mpWood.Wood.DocumentCollection,
                mpMaterial.Material.DocumentCollection,
                mpOther.Other.DocumentCollection
            };
            var resultLst = new List<BaseDocument>();

            foreach (var collection in documentCollections)
            {
                resultLst.AddRange(collection.Where(
                    baseDocument => baseDocument.DocumentName.ToLower().Contains(stringForSearch.ToLower()) |
                    baseDocument.DocumentNumber.ToLower().Contains(stringForSearch.ToLower()) |
                    baseDocument.DocumentShortName.ToLower().Contains(stringForSearch.ToLower()) |
                    baseDocument.DocumentType.ToLower().Contains(stringForSearch.ToLower())));
            }
            if (resultLst.Count > 0)
                LbSearchResults.ItemsSource = resultLst;
        }
        // Выбор элемента в списке результатов
        private void LbSearchResults_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var listBox = sender as ListBox;
            if (listBox?.SelectedItem is BaseDocument baseElement)
            {
                // Выбираем в списке базу
                foreach (ComboBoxItem item in CbDataBases.Items)
                {
                    if (item.Name.Equals(baseElement.DataBaseName))
                    {
                        CbDataBases.SelectedItem = item;
                        break;
                    }
                }
                // Выбираем в списке группу
                listBox = LbGroups;
                if (listBox?.ItemsSource != null)
                    foreach (var item in listBox.Items)
                    {
                        if (item.ToString().Equals(baseElement.Group))
                        {
                            listBox.SelectedItem = item;
                        }
                    }
                // Выбираем в списке документ
                listBox = LbDocuments;
                if (listBox != null && listBox.Visibility == Visibility.Visible)
                {
                    if (listBox.ItemsSource != null)
                        foreach (BaseDocument item in listBox.Items)
                        {
                            if (item.Equals(baseElement))
                            {
                                listBox.SelectedItem = item;
                            }
                        }
                }
                var treeView = TvDocuments;
                if (treeView != null && treeView.Visibility == Visibility.Visible)
                {
                    if (treeView.ItemsSource != null)
                        foreach (var item in treeView.Items)
                        {
                            if (item is GroupByShortName)
                            {
                                foreach (var doc in (item as GroupByShortName).Documents)
                                {
                                    if (doc.Equals(baseElement))
                                    {
                                        treeView.SelectItem(doc);
                                    }
                                }
                            }
                        }
                }
            }
        }
        // Открыть панель поиска
        private void BtOpenSearchPanel_OnClick(object sender, RoutedEventArgs e)
        {
            FlyoutSearch.IsOpen = !FlyoutSearch.IsOpen;
            TbSearchTxt.Focus();
        }
        #endregion
        // Закрытие по Esc
        private void MainWindow_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) Close();
        }
        // Навели мышку на кнопку
        private void DocumentShowButtons_OnMouseEnter(object sender, MouseEventArgs e)
        {
            if (sender is Button btn) btn.Opacity = 1;
        }

        // Убрали мышку с кнопки
        private void DocumentShowButtons_OnMouseLeave(object sender, MouseEventArgs e)
        {
            var btn = sender as Button;
            if (btn != null) btn.Opacity = 0.5;
        }
        // Нажатие кнопки "Отобразить в виде списка" 
        private void BtShowAsList_OnClick(object sender, RoutedEventArgs e)
        {
            TvDocuments.Visibility = Visibility.Collapsed;
            LbDocuments.Visibility = Visibility.Visible;
            BtShowAsList.Visibility = Visibility.Collapsed;
            BtShowAsTree.Visibility = Visibility.Visible;
            // Если был выбран элемент
            if (TvDocuments.SelectedItem is BaseDocument)
            {
                LbDocuments.SelectionChanged -= LbDocuments_OnSelectionChanged;
                LbDocuments.SelectedItem = TvDocuments.SelectedItem;
                LbDocuments.ScrollIntoView(LbDocuments.SelectedItem);
                LbDocuments.SelectionChanged += LbDocuments_OnSelectionChanged;
            }
        }
        // Нажатие кнопки "Отобразить в виде дерева" 
        private void BtShowAsTree_OnClick(object sender, RoutedEventArgs e)
        {
            TvDocuments.Visibility = Visibility.Visible;
            LbDocuments.Visibility = Visibility.Collapsed;
            BtShowAsList.Visibility = Visibility.Visible;
            BtShowAsTree.Visibility = Visibility.Collapsed;
            // Если был выбран элемент
            if (LbDocuments.SelectedItem is BaseDocument)
            {
                TvDocuments.SelectedItemChanged -= TvDocuments_OnSelectedItemChanged;
                TvDocuments.SelectItem(LbDocuments.SelectedItem);
                TvDocuments.SelectedItemChanged += TvDocuments_OnSelectedItemChanged;
            }
        }
        // Выбор документа на сталь
        private void CbSteelDocument_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CbSteelType.ItemsSource = null;
            var comboBox = sender as ComboBox;
            var steelDoc = comboBox?.SelectedItem;
            if (steelDoc is Steel steel)
            {
                CbSteelType.ItemsSource = steel.Values;
                comboBox.ToolTip = steel.DocumentName;
                CbSteelType.SelectedIndex = 0;
                ShowItemNaim();
            }
        }

        private void ShowItemNaim()
        {
            // when return = clean textBox!!!
            if (_currentDocument == null)
            {
                TbNaimFirst.Text = string.Empty;
                TbNaimSecond.Text = string.Empty;
                return;
            }
            // Получаем правило написания для документа
            var rule = _currentDocument.Rule;
            if (string.IsNullOrEmpty(rule))
            {
                TbNaimFirst.Text = string.Empty;
                TbNaimSecond.Text = string.Empty;
                return;
            }
            var brkResult = ModPlusAPI.IO.String.BreakString(rule, '[', ']');
            var sb = new StringBuilder();
            var selectedItem = DgItems.SelectedItem as XElement;
            // Проходим по списку знаков сверяя его с атрибутами (и не только)
            foreach (var _char in brkResult)
            {
                // Добавляем вспомогательгую переменную
                var appended = false;
                // Проходим по атрибутам в документе
                foreach (var docAttr in _currentDocument.XmlDocument.Attributes())
                {
                    if (docAttr.Name.ToString().Equals(_char) && !docAttr.Name.ToString().Contains("ItemType"))
                    {
                        sb.Append(docAttr.Value);
                        appended = true;
                        break;
                    }
                }
                // проходим по ItemTypes
                if (_currentDocument.ItemTypes.Count > 0)
                    foreach (BaseDocument.ItemType itemType in LbDocumentTypes.Items)
                    {
                        if (itemType.TypeName.Equals(_char))
                        {
                            sb.Append(itemType.SelectedItem);
                            appended = true;
                            break;
                        }
                    }

                if (selectedItem != null) // Если выбран табличный элемент
                {
                    foreach (var attribute in selectedItem.Attributes())
                    {
                        if (attribute.Name.ToString().Equals(_char))
                        {
                            sb.Append(attribute.Value);
                            appended = true;
                            break;
                        }
                    }
                }
                // Если предыдущие проверки не дали результат, значит это просто текст
                if (!appended) sb.Append(_char);
            }
            TbNaimFirst.Text = sb.ToString();
            if (CbSteelDocument.SelectedItem is Steel steel)
                TbNaimSecond.Text = CbSteelType.SelectedItem + " " + steel.Document;
        }

        #region Export
        // Експорт документа в Excel
        private void BtExportDocumentToExcel_OnClick(object sender, RoutedEventArgs e)
        {
            if (_currentDocument != null)
            {
                var dialogProgress = new ExportProgressDialog(ModPlusAPI.Language.GetItem(LangItem, "h12"), ExportDocumentToExcel) { Topmost = true };
                dialogProgress.ShowDialog();
            }
        }

        private void ExportDocumentToExcel(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            worker.ReportProgress(0, ModPlusAPI.Language.GetItem(LangItem, "p1"));
            // Создание документа
            // Start Excel and Create new document
            Excel.Range rng;
            Excel._Application oExcel = new Excel.Application();
            oExcel.Visible = true;
            Excel._Workbook oBook = oExcel.Workbooks.Add();
            var oSheet = oExcel.ActiveSheet as Excel._Worksheet;
            try
            {
                // Стили
                //***************************
                Excel.Style oStyle = oBook.Styles.Add("MP_STYLE");
                oStyle.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                oStyle.WrapText = true;
                oStyle.Font.Size = 10;
                oStyle.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                //***************************
                Excel.Style oStyle2 = oBook.Styles.Add("HEAD_STYLE");
                oStyle2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                oStyle2.WrapText = true;
                oStyle2.Font.Bold = true;
                oStyle2.Font.Size = 10;
                oStyle2.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                //***************************
                rng = oSheet.get_Range("A1");
                rng.set_Value(null, ModPlusAPI.Language.GetItem(LangItem, "p2") + " " + _currentDocument.DocumentType + " " + _currentDocument.DocumentNumber + "\n");
                rng.Style = "HEAD_STYLE";
                rng.WrapText = false;
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                rng.Rows.AutoFit();
                //******************************
                rng = oSheet.Range["A2"];
                rng.set_Value(null, ModPlusAPI.Language.GetItem(LangItem, "h4") + ": " + _currentDocument.DocumentName + "\"" + "\n");
                rng.Style = "HEAD_STYLE";
                rng.WrapText = false;
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                //**************************************

                if (_currentDocument.SymbolCount > 0)
                {
                    worker.ReportProgress(0, ModPlusAPI.Language.GetItem(LangItem, "p3"));
                    // Заполняем шапку таблицы
                    for (var j = 0; j < _currentDocument.SymbolCount; j++)
                    {
                        rng = (Excel.Range)oSheet.Cells[3, j + 1];
                        rng.Style = "HEAD_STYLE";
                        rng.set_Value(null, _currentDocument.Symbols.ElementAt(j));
                        rng.BorderAround();
                        rng.Borders.Weight = 3;
                        rng.WrapText = false;
                    }
                    rng = oSheet.get_Range((Excel.Range)oSheet.Cells[3, 1], (Excel.Range)oSheet.Cells[3, _currentDocument.SymbolCount]);
                    rng.Columns.AutoFit();
                    //////////////////////////////////////////////////////////////////
                    var maxRecords = _currentDocument.Items.Elements("Item").Count();
                    // Заполняем ячейки
                    var i = 4;
                    var x = 0;
                    foreach (var item in _currentDocument.Items.Elements("Item"))
                    {
                        if (worker.CancellationPending) // See if cacel button was pressed.
                        {
                            oExcel.Quit();
                            ReleaseComObject(oExcel);
                            break;
                        }
                        worker.ReportProgress(Convert.ToInt32(((decimal)x / maxRecords) * 100), ModPlusAPI.Language.GetItem(LangItem, "p4"));

                        for (int j = 0; j < _currentDocument.SymbolCount; j++)
                        {
                            rng = (Excel.Range)oSheet.Cells[i, j + 1];
                            rng.Style = "MP_STYLE";
                            rng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            rng.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            string str = item.Attribute("Prop" + (j + 1).ToString(CultureInfo.InvariantCulture)).Value;
                            rng.set_Value(null, str);
                            rng.BorderAround();
                            rng.Borders.Weight = 2;
                        }
                        i++;
                        x++;
                    }
                }
                ReleaseComObject(oExcel);
            }
            catch (Exception)
            {
                oExcel.Quit();
                ReleaseComObject(oExcel);
            }
        }
        private static void ReleaseComObject(object excel)
        {
            // Уничтожение объекта Excel.
            Marshal.ReleaseComObject(excel);
            // Вызываем сборщик мусора для немедленной очистки памяти
            GC.GetTotalMemory(true);
        }
        // Получить список всех документов в базе данных
        private void BtExportDocumentsNameToTxtFile_OnClick(object sender, RoutedEventArgs e)
        {
            var dialogProgress = new ExportProgressDialog(ModPlusAPI.Language.GetItem(LangItem, "h18"), GetAllDocuments) { Topmost = true };
            dialogProgress.ShowDialog();
        }
        private static void GetAllDocuments(object sender, DoWorkEventArgs e)
        {
            // Список доступных коллекций
            var documentCollections = new List<ICollection<BaseDocument>>
            {
                mpMetall.Metall.DocumentCollection,
                mpConcrete.Concrete.DocumentCollection,
                mpWood.Wood.DocumentCollection,
                mpMaterial.Material.DocumentCollection,
                mpOther.Other.DocumentCollection
            };
            var str = string.Empty;
            var i = 0;
            foreach (var collection in documentCollections)
            {
                var documents = collection.Select(document => document.DocumentType + " " + document.DocumentNumber + " " + document.DocumentName + Environment.NewLine).ToList();

                documents.Sort();

                foreach (var doc in documents.Distinct())
                {
                    i++;
                    str += i + "\t" + doc;
                }
            }
            ModPlusAPI.IO.String.ShowTextWithNotepad(str);
        }

        // Експорт документа в Word
        private void BtExportDocumentToWord_OnClick(object sender, RoutedEventArgs e)
        {
            if (_currentDocument != null)
            {
                var dialogProgress = new ExportProgressDialog(ModPlusAPI.Language.GetItem(LangItem, "h13"), ExportDocumentToWord) { Topmost = true };
                dialogProgress.ShowDialog();
            }
        }

        private void ExportDocumentToWord(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            worker.ReportProgress(0, ModPlusAPI.Language.GetItem(LangItem, "p1"));
            object SaveChanged = Word.WdSaveOptions.wdDoNotSaveChanges;
            var oWord = new Word.Application();
            try
            {
                // Делаем его видимым
                oWord.Visible = true;
                // Создаем новый документ
                Object template = Type.Missing;
                Object newTemplate = false;
                Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
                Object visible = true;
                oWord.Documents.Add(ref template, ref newTemplate, ref documentType, ref visible);
                // Выбираем его
                Word.Document oDoc = oWord.Documents.get_Item(1);
                // Свойства документа:
                // Поля
                oDoc.PageSetup.LeftMargin = oWord.CentimetersToPoints((float)1.0);
                oDoc.PageSetup.TopMargin = oWord.CentimetersToPoints((float)1.0);
                oDoc.PageSetup.RightMargin = oWord.CentimetersToPoints((float)1.0);
                oDoc.PageSetup.BottomMargin = oWord.CentimetersToPoints((float)1.0);
                // Создаем 2 новых параграфа
                object oMissing = Missing.Value;
                oDoc.Paragraphs.Add(ref oMissing);
                Word.Paragraph oParagraph1 = oDoc.Paragraphs[1];
                oParagraph1.Range.Text = ModPlusAPI.Language.GetItem(LangItem, "p2") + " " + _currentDocument.DocumentType + " " + _currentDocument.DocumentNumber + "\n";
                ///////////////////////////////////////////////
                oDoc.Paragraphs.Add(ref oMissing);
                Word.Paragraph oParagraph2 = oDoc.Paragraphs[2];
                oParagraph2.Range.Text = ModPlusAPI.Language.GetItem(LangItem, "h4") + ": " + _currentDocument.DocumentName + "\"" + "\n";
                // Изображение
                //////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////////////////////////////
                // Создаем третий параграф, если нужен
                // Таблицу вставляем в третий параграф
                if (_currentDocument.SymbolCount > 0)
                {
                    oDoc.Paragraphs.Add(ref oMissing);
                    Word.Paragraph oParagraph4 = oDoc.Paragraphs[4];
                    Word.Range oRange2 = oParagraph4.Range;
                    Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;// Автоматически менять ячейки по тексту
                    Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;// Автоподбор ширины столбцов
                    var maxRecords = _currentDocument.Items.Elements("Item").Count();
                    //Добавляем таблицу и получаем объект wordtable
                    oDoc.Tables.Add(oRange2, (_currentDocument.Items.Elements("Item").Count() + 1),// Число строк
                                                                                                   //2,// Число строк
                                    _currentDocument.SymbolCount,// Число столбцов
                                    ref defaultTableBehavior,
                                    ref autoFitBehavior);
                    Word.Table oTable = oDoc.Tables[1];
                    worker.ReportProgress(0, ModPlusAPI.Language.GetItem(LangItem, "p3"));
                    // Заполняем шапку таблицы
                    for (var j = 0; j < oTable.Columns.Count; j++)
                    {
                        Word.Range wordcellrange = oTable.Cell(1, j + 1).Range;
                        wordcellrange.ParagraphFormat.Alignment =
                                Word.WdParagraphAlignment.wdAlignParagraphCenter;// Выравнивание в ячейке
                        wordcellrange.Text = _currentDocument.Symbols.ElementAt(j);
                    }
                    // Заполняем ячейки
                    var i = 3;
                    var x = 0;
                    foreach (var item in _currentDocument.Items.Elements("Item"))
                    {
                        if (worker.CancellationPending) // See if cacel button was pressed.
                        {
                            oWord.Quit();
                            ReleaseComObject(oWord);
                            break;
                        }
                        worker.ReportProgress(Convert.ToInt32(((decimal)x / maxRecords) * 100), ModPlusAPI.Language.GetItem(LangItem, "p4"));

                        for (var j = 0; j < _currentDocument.SymbolCount; j++)
                        {
                            Word.Range wordcellrange = oTable.Cell(i - 1, j + 1).Range;
                            wordcellrange.ParagraphFormat.Alignment =
                                Word.WdParagraphAlignment.wdAlignParagraphCenter;// Выравнивание в ячейке
                            wordcellrange.Text = item.Attribute("Prop" + (j + 1).ToString(CultureInfo.InvariantCulture)).Value;
                        }
                        i++;
                        x++;
                    }
                }
                ReleaseComObject(oWord);
            }
            catch (Exception)
            {
                oWord.Quit(SaveChanged, Type.Missing, Type.Missing);
                ReleaseComObject(oWord);
            }
        }
        #endregion
        // Выбор элемента в таблице
        private void DgItems_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var dataGrid = sender as DataGrid;
            var selected = dataGrid?.SelectedItem;
            if (selected != null)
                ShowItemNaim();
        }
        // Выбор Itemtype
        private void CbItemType_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ShowItemNaim();
        }
        // Выбор марки стали
        private void CbSteelType_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ShowItemNaim();
        }
        // Window closing
        private void MpDbviewerWindow_OnClosed(object sender, EventArgs e)
        {
            var cb = CbDataBases;
            // Сохраняем в файл настроек выбранную базу
            if (cb?.SelectedIndex != -1)
                UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "mpDBviewer", "selectedDB", cb?.SelectedIndex.ToString(), true);
            // Сохраняем выбранную группу
            var lb = LbGroups;
            if (lb?.SelectedIndex != -1)
                UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "mpDBviewer", "selectedGroup", cb?.SelectedIndex.ToString(), true);
        }

    }
}
