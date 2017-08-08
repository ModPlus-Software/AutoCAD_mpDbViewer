using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
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
using mpSettings;
using ModPlus;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace mpDbViewer
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MpDbviewerWindow
    {
        // Текущая коллекция документов (база данных) с которой работаю
        private ICollection<BaseDocument> _currentDocumentCollection;
        // Текущий выбранный документ
        private BaseDocument _currentDocument;
        public MpDbviewerWindow()
        {
            InitializeComponent();
            MpWindowHelpers.OnWindowStartUp(
                this,
                MpSettings.GetValue("Settings", "MainSet", "Theme"),
                MpSettings.GetValue("Settings", "MainSet", "AccentColor"),
                MpSettings.GetValue("Settings", "MainSet", "BordersType")
                );
        }

        private void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            // ОБЯЗАТЕЛЬНО!
            // "Загрузка" документов баз данных
            mpMetall.Metall.LoadAllDocument();
            mpConcrete.Concrete.LoadAllDocument();
            mpWood.Wood.LoadAllDocument();
            mpMaterial.Material.LoadAllDocument();
            mpOther.Other.LoadAllDocument();
            // Марки стали - документы
            this.CbSteelDocument.ItemsSource = SteelDocuments.GetSteels();
            this.CbSteelDocument.SelectedIndex = 0;
            // Загружаем значение из файла настроек
            var selectedDb = MpSettings.GetValue("Settings", "mpDBviewer", "selectedDB");
            if (!string.IsNullOrEmpty(selectedDb))
            {
                int index;
                if (int.TryParse(selectedDb, out index))
                    this.CbDataBases.SelectedIndex = index;
            }
            var selectedGroup = MpSettings.GetValue("Settings", "mpDBviewer", "selectedGroup");
            if (!string.IsNullOrEmpty(selectedGroup))
            {
                int index;
                if (int.TryParse(selectedGroup, out index))
                {
                    try
                    {
                        this.LbGroups.SelectedIndex = index;
                    }
                    catch
                    {
                        //ignored
                    }
                }
            }
        }
        private void MpDbviewerWindow_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
        // Очистка всех контроллов
        private void ClearAllControls()
        {
            this.LbGroups.ItemsSource = null;
            ClearAllExceptGroups();
        }
        // Очистка всех, кроме списка групп
        private void ClearAllExceptGroups()
        {
            this.DgItems.Columns.Clear();
            this.DgItems.ItemsSource = null;
            this.LbDocuments.ItemsSource = null;
            this.BtShowImage.IsEnabled = false;
            this.TabItemExport.Visibility = Visibility.Collapsed;
            this.TbDocumentName.Text = string.Empty;
            // 
            this.TabControlDetail.SelectedIndex = 0;
            this.StkSteel.Visibility = Visibility.Collapsed;
            this.LbDocumentTypes.ItemsSource = null;
            this.StkNaim.Visibility = Visibility.Collapsed;
            this.TvDocuments.ItemsSource = null;
            this.BtShowAsTree.Visibility = Visibility.Collapsed;
            this.BtShowAsList.Visibility = Visibility.Collapsed;
            _currentDocument = null;
            this.TbNaimFirst.Text = string.Empty;
            this.TbNaimSecond.Text = string.Empty;
        }
        // Выбор базы данных
        private void CbDataBases_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ClearAllControls();
            var comboBox = sender as ComboBox;
            if (comboBox != null)
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
                this.LbGroups.ItemsSource = _currentDocumentCollection.Select(x => x.Group).ToList().Distinct();
        }
        // Выбор группы
        private void LbGroups_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ClearAllExceptGroups();
            var listbox = sender as ListBox;
            if (listbox != null)
                if (listbox.ItemsSource != null)
                {
                    var selectedGroup = listbox.SelectedItem.ToString();
                    if (!string.IsNullOrEmpty(selectedGroup))
                    {
                        // Если для указанной группы значений ShortName больше 1, то отображаем в виде дерева
                        // если всего 1, то в виде списка
                        // Но заполняем и то, и другое, чтобы можно было переключаться
                        var lst = _currentDocumentCollection.Where(x => x.Group.Equals(selectedGroup)).ToList();
                        lst.Sort(new DocumentsSortComparer());
                        this.LbDocuments.ItemsSource = lst;

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

                        this.TvDocuments.ItemsSource = hlst;

                        if (hlst.Count > 1)
                        {
                            this.TvDocuments.Visibility = Visibility.Visible;
                            this.LbDocuments.Visibility = Visibility.Collapsed;
                            this.BtShowAsList.Visibility = Visibility.Visible;
                            this.BtShowAsTree.Visibility = Visibility.Collapsed;
                        }
                        else
                        {
                            this.TvDocuments.Visibility = Visibility.Collapsed;
                            this.LbDocuments.Visibility = Visibility.Visible;
                            this.BtShowAsList.Visibility = Visibility.Collapsed;
                            this.BtShowAsTree.Visibility = Visibility.Visible;
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

            public string ShortName { private get; set; }
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
            if (listbox != null)
            {
                var docInBase = listbox.SelectedItem as BaseDocument;
                if (docInBase != null) FillDataGridWithItems(docInBase);
            }
        }
        // Выбор документа в TreeView
        private void TvDocuments_OnSelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var treeView = sender as TreeView;
            if (treeView != null)
            {
                var tvi = treeView.SelectedItem;
                if (tvi is BaseDocument)
                    if (tvi != null) FillDataGridWithItems(tvi as BaseDocument);
            }
        }

        private void FillDataGridWithItems(BaseDocument baseDocument)
        {
            this.DgItems.Columns.Clear();
            this.DgItems.ItemsSource = null;
            // Текущий выбранный документ
            _currentDocument = baseDocument;
            // Заполняем вкладку "Свойства"
            foreach (var itemType in baseDocument.ItemTypes)
            {
                itemType.SelectedItem = itemType.TypeValues[0];
            }
            this.LbDocumentTypes.ItemsSource = baseDocument.ItemTypes;

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
                    this.DgItems.Columns.Add(dgc);
                }
                this.DgItems.ItemsSource = baseDocument.Items.Elements("Item");
                // Сразу выберем первый элемент в списке, чтобы отобразилось условное обозначение
                this.DgItems.SelectedIndex = 0;
            }
            else ShowItemNaim(); // Так как при отсутствии таблицы метода не сработает автоматически
            // Если SymbolCount = 0, значит таблицы нет. Тогда сразу открываем вторую вкладку
            // если вторая вкладка была открыта уже, то не меняем ничего (для удобства пользования)
            if (this.TabControlDetail.SelectedIndex != 1)
                this.TabControlDetail.SelectedIndex = baseDocument.SymbolCount == 0 ? 1 : 0;

            // Активация кнопки "изображение"
            this.BtShowImage.IsEnabled = !string.IsNullOrEmpty(baseDocument.Image);
            // Название
            this.TbDocumentName.DataContext = baseDocument;
            // Видимость кнопок и прочего
            this.TabItemExport.Visibility = Visibility.Visible;
            // Если есть сталь
            this.StkSteel.Visibility = baseDocument.HasSteel ? Visibility.Visible : Visibility.Collapsed;
            this.NaimSplitter.Visibility = baseDocument.HasSteel ? Visibility.Visible : Visibility.Collapsed;
            this.TbNaimSecond.Visibility = baseDocument.HasSteel ? Visibility.Visible : Visibility.Collapsed;
            // Включаем отображение "Наименовани"
            this.StkNaim.Visibility = Visibility.Visible;
        }

        private void BtShowImage_OnClick(object sender, RoutedEventArgs e)
        {
            BaseDocument docInBase = null;
            if (this.LbDocuments.Visibility == Visibility.Visible)
                docInBase = this.LbDocuments.SelectedItem as BaseDocument;
            if (this.TvDocuments.Visibility == Visibility.Visible)
                docInBase = this.TvDocuments.SelectedItem as BaseDocument;
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
            var textbox = sender as TextBox;
            if (textbox != null)
            {
                if (textbox.Text.Length >= 2)
                    SearchDocumentsInDb(textbox.Text);
                else this.LbSearchResults.ItemsSource = null;
            }
        }

        private void SearchDocumentsInDb(string stringForSearch)
        {
            this.LbSearchResults.ItemsSource = null;

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
                this.LbSearchResults.ItemsSource = resultLst;
        }
        // Выбор элемента в списке результатов
        private void LbSearchResults_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var listBox = sender as ListBox;
            if (listBox != null)
            {
                var baseElement = listBox.SelectedItem as BaseDocument;
                if (baseElement != null)
                {
                    // Выбираем в списке базу
                    foreach (ComboBoxItem item in this.CbDataBases.Items)
                    {
                        if (item.Name.Equals(baseElement.DataBaseName))
                        {
                            this.CbDataBases.SelectedItem = item;
                            break;
                        }
                    }
                    // Выбираем в списке группу
                    listBox = this.LbGroups;
                    if (listBox != null)
                    {
                        if (listBox.ItemsSource != null)
                            foreach (var item in listBox.Items)
                            {
                                if (item.ToString().Equals(baseElement.Group))
                                {
                                    listBox.SelectedItem = item;
                                }
                            }
                    }
                    // Выбираем в списке документ
                    listBox = this.LbDocuments;
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
                    var treeView = this.TvDocuments;
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
        }
        // Открыть панель поиска
        private void BtOpenSearchPanel_OnClick(object sender, RoutedEventArgs e)
        {
            this.FlyoutSearch.IsOpen = !this.FlyoutSearch.IsOpen;
            this.TbSearchTxt.Focus();
        }
        #endregion
        // Закрытие по Esc
        private void MainWindow_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) this.Close();
        }
        // Навели мышку на кнопку
        private void DocumentShowButtons_OnMouseEnter(object sender, MouseEventArgs e)
        {
            var btn = sender as Button;
            if (btn != null) btn.Opacity = 1;
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
            this.TvDocuments.Visibility = Visibility.Collapsed;
            this.LbDocuments.Visibility = Visibility.Visible;
            this.BtShowAsList.Visibility = Visibility.Collapsed;
            this.BtShowAsTree.Visibility = Visibility.Visible;
            // Если был выбран элемент
            if (this.TvDocuments.SelectedItem != null)
                if (this.TvDocuments.SelectedItem is BaseDocument)
                {
                    this.LbDocuments.SelectionChanged -= LbDocuments_OnSelectionChanged;
                    this.LbDocuments.SelectedItem = this.TvDocuments.SelectedItem;
                    this.LbDocuments.ScrollIntoView(this.LbDocuments.SelectedItem);
                    this.LbDocuments.SelectionChanged += LbDocuments_OnSelectionChanged;
                }
        }
        // Нажатие кнопки "Отобразить в виде дерева" 
        private void BtShowAsTree_OnClick(object sender, RoutedEventArgs e)
        {
            this.TvDocuments.Visibility = Visibility.Visible;
            this.LbDocuments.Visibility = Visibility.Collapsed;
            this.BtShowAsList.Visibility = Visibility.Visible;
            this.BtShowAsTree.Visibility = Visibility.Collapsed;
            // Если был выбран элемент
            if (this.LbDocuments.SelectedItem != null)
                if (this.LbDocuments.SelectedItem is BaseDocument)
                {
                    this.TvDocuments.SelectedItemChanged -= TvDocuments_OnSelectedItemChanged;
                    this.TvDocuments.SelectItem(this.LbDocuments.SelectedItem);
                    this.TvDocuments.SelectedItemChanged += TvDocuments_OnSelectedItemChanged;
                }
        }
        // Выбор документа на сталь
        private void CbSteelDocument_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            this.CbSteelType.ItemsSource = null;
            var comboBox = sender as ComboBox;
            if (comboBox != null)
            {
                var steelDoc = comboBox.SelectedItem;
                if (steelDoc != null)
                {
                    var steel = steelDoc as Steel;
                    if (steel != null)
                    {
                        this.CbSteelType.ItemsSource = steel.Values;
                        comboBox.ToolTip = steel.DocumentName;
                        this.CbSteelType.SelectedIndex = 0;
                        ShowItemNaim();
                    }
                }
            }
        }

        private void ShowItemNaim()
        {
            // when return = clean textBox!!!
            if (_currentDocument == null)
            {
                this.TbNaimFirst.Text = string.Empty;
                this.TbNaimSecond.Text = string.Empty;
                return;
            }
            // Получаем правило написания для документа
            var rule = _currentDocument.Rule;
            if (string.IsNullOrEmpty(rule))
            {
                this.TbNaimFirst.Text = string.Empty;
                this.TbNaimSecond.Text = string.Empty;
                return;
            }
            var brkResult = BreakString(rule, '[', ']');
            var sb = new StringBuilder();
            var selectedItem = this.DgItems.SelectedItem as XElement;
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
                    foreach (BaseDocument.ItemType itemType in this.LbDocumentTypes.Items)
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
            this.TbNaimFirst.Text = sb.ToString();
            var steel = this.CbSteelDocument.SelectedItem as Steel;
            if (steel != null)
                this.TbNaimSecond.Text = this.CbSteelType.SelectedItem + " " + steel.Document;
        }
        private static IEnumerable<string> BreakString(string str, char symbol1, char symbol2)
        {
            var result = new List<string>();
            var k = -1;
            var sb = new StringBuilder();
            for (var i = 0; i < str.Length; i++)
            {
                if (str[i].Equals(symbol1))
                {
                    if (sb.Length > 0)
                        result.Insert(k, sb.ToString());
                    sb = new StringBuilder();
                    if (i > 1)
                        if (!str[i - 1].Equals(symbol2))
                            k++;
                }
                else if (str[i].Equals(symbol2))
                {
                    result.Insert(k, sb.ToString());
                    sb = new StringBuilder();
                    k++;
                }
                else
                {
                    if (k == -1)
                        k++;
                    sb.Append(str[i]);
                }
            }
            return result;
        }
        #region Export
        // Експорт документа в Excel
        private void BtExportDocumentToExcel_OnClick(object sender, RoutedEventArgs e)
        {
            if (_currentDocument != null)
            {
                var dialogProgress = new ExportProgressDialog("Экспорт в Excel", ExportDocumentToExcel) { Topmost = true };
                dialogProgress.ShowDialog();
            }
        }

        private void ExportDocumentToExcel(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            worker.ReportProgress(0, "Создание нового документа");
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
                rng.set_Value(null, "Номер документа: " + _currentDocument.DocumentType + " " + _currentDocument.DocumentNumber + "\n");
                rng.Style = "HEAD_STYLE";
                rng.WrapText = false;
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                rng.Rows.AutoFit();
                //******************************
                rng = oSheet.Range["A2"];
                rng.set_Value(null, "Название документа: " + _currentDocument.DocumentName + "\"" + "\n");
                rng.Style = "HEAD_STYLE";
                rng.WrapText = false;
                rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                //**************************************

                if (_currentDocument.SymbolCount > 0)
                {
                    worker.ReportProgress(0, "Заполнение шапки");
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
                        worker.ReportProgress(Convert.ToInt32(((decimal)x / maxRecords) * 100), "Заполнение данных");

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
            catch (System.Exception)
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
            var dialogProgress = new ExportProgressDialog("Список документов в базе", GetAllDocuments) { Topmost = true };
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
            NotepadHelper.ShowMessage(str);


        }
        private static class NotepadHelper
        {
            [DllImport("user32.dll", EntryPoint = "SetWindowText")]
            private static extern int SetWindowText(IntPtr hWnd, string text);

            [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
            private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

            [DllImport("User32.dll", EntryPoint = "SendMessage")]
            private static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);

            public static void ShowMessage(string message = null, string title = null)
            {
                var notepad = Process.Start(new ProcessStartInfo("notepad.exe"));
                notepad.WaitForInputIdle();

                if (!string.IsNullOrEmpty(title))
                    SetWindowText(notepad.MainWindowHandle, title);

                if (notepad != null && !string.IsNullOrEmpty(message))
                {
                    var child = FindWindowEx(notepad.MainWindowHandle, new IntPtr(0), "Edit", null);
                    SendMessage(child, 0x000C, 0, message);
                }
            }
        }
        // Експорт документа в Word
        private void BtExportDocumentToWord_OnClick(object sender, RoutedEventArgs e)
        {
            if (_currentDocument != null)
            {
                var dialogProgress = new ExportProgressDialog("Экспорт в Word", ExportDocumentToWord) { Topmost = true };
                dialogProgress.ShowDialog();
            }
        }

        private void ExportDocumentToWord(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;
            worker.ReportProgress(0, "Создание нового документа");
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
                oParagraph1.Range.Text = "Номер документа: " + _currentDocument.DocumentType + " " + _currentDocument.DocumentNumber + "\n";
                ///////////////////////////////////////////////
                oDoc.Paragraphs.Add(ref oMissing);
                Word.Paragraph oParagraph2 = oDoc.Paragraphs[2];
                oParagraph2.Range.Text = "Название документа: " + _currentDocument.DocumentName + "\"" + "\n";
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
                    worker.ReportProgress(0, "Заполнение шапки");
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
                        worker.ReportProgress(Convert.ToInt32(((decimal)x / maxRecords) * 100), "Заполнение данных");

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
            catch (System.Exception)
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
            if (dataGrid != null)
            {
                var selected = dataGrid.SelectedItem;
                if (selected != null)
                    ShowItemNaim();
            }
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
            var cb = this.CbDataBases;
            // Сохраняем в файл настроек выбранную базу
            if (cb?.SelectedIndex != -1)
                MpSettings.SetValue("Settings", "mpDBviewer", "selectedDB", cb?.SelectedIndex.ToString(), true);
            // Сохраняем выбранную группу
            var lb = this.LbGroups;
            if (lb?.SelectedIndex != -1)
                MpSettings.SetValue("Settings", "mpDBviewer", "selectedGroup", cb?.SelectedIndex.ToString(), true);
        }

    }
}
