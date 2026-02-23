using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace TextEditor
{
    public class ResultItem
    {
        public int Line { get; set; }
        public string Type { get; set; }
        public string Message { get; set; }
    }

    public class ErrorItem
    {
        public int Line { get; set; }
        public int Column { get; set; }
        public string Message { get; set; }
    }

    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private string currentFilePath;
        private bool isModified = false;
        private ObservableCollection<ResultItem> results = new ObservableCollection<ResultItem>();
        private ObservableCollection<ErrorItem> errors = new ObservableCollection<ErrorItem>();
        private double currentFontSize = 12;
        private int tabCounter = 1;
        private string currentLanguage = "ru-RU";
        private bool isClosing = false;
        private DispatcherTimer updateTimer;
        private bool isDarkTheme = false;
        private bool isUpdatingLineNumbers = false;
        private bool _isUpdating = false;
        private DispatcherTimer _analysisTimer;
        private bool _ignoreTextChanges = false;

        public event PropertyChangedEventHandler PropertyChanged;

        public ObservableCollection<ResultItem> Results
        {
            get { return results; }
            set
            {
                results = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Results"));
            }
        }

        public ObservableCollection<ErrorItem> Errors
        {
            get { return errors; }
            set
            {
                errors = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Errors"));
            }
        }

        public MainWindow()
        {
            try
            {
                InitializeComponent();
                DataContext = this;

                updateTimer = new DispatcherTimer();
                updateTimer.Interval = TimeSpan.FromMilliseconds(500);
                updateTimer.Tick += UpdateTimer_Tick;
                updateTimer.Start();

                UpdateLineNumbers();
                UpdateStatusBar();
                UpdateUILanguage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при инициализации: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                UpdateStatusBar();
            }
            catch { }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                string[] args = Environment.GetCommandLineArgs();
                if (args.Length > 1 && File.Exists(args[1]))
                {
                    OpenFile(args[1]);
                }
            }
            catch { }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.F1)
                {
                    Help_Click(sender, e);
                    e.Handled = true;
                }
            }
            catch { }
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            try
            {
                if (isClosing) return;

                if (!CheckSaveChanges())
                {
                    e.Cancel = true;
                    return;
                }

                isClosing = true;

                if (updateTimer != null)
                {
                    updateTimer.Stop();
                    updateTimer = null;
                }

                Application.Current.Shutdown();
                Process.GetCurrentProcess().Kill();
            }
            catch
            {
                Environment.Exit(0);
            }
        }

        // ==================== СИНХРОНИЗАЦИЯ ПРОКРУТКИ ====================
        private void EditorTextBox_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            try
            {
                if (isUpdatingLineNumbers) return;

                isUpdatingLineNumbers = true;

                var scrollViewer = FindVisualChild<ScrollViewer>(EditorTextBox);
                if (scrollViewer != null && LineNumbersScrollViewer != null)
                {
                    LineNumbersScrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset);
                }

                isUpdatingLineNumbers = false;
            }
            catch { }
        }

        private T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child != null && child is T)
                    return (T)child;
                else
                {
                    var descendant = FindVisualChild<T>(child);
                    if (descendant != null)
                        return descendant;
                }
            }
            return null;
        }

        // ==================== ПЕРЕКЛЮЧЕНИЕ ТЕМЫ ====================
        private void LightTheme_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                isDarkTheme = false;
                LightThemeMenuItem.IsChecked = true;
                DarkThemeMenuItem.IsChecked = false;
                ApplyTheme();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при смене темы: {ex.Message}");
            }
        }

        private void DarkTheme_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                isDarkTheme = true;
                LightThemeMenuItem.IsChecked = false;
                DarkThemeMenuItem.IsChecked = true;
                ApplyTheme();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при смене темы: {ex.Message}");
            }
        }

        private void ApplyTheme()
        {
            try
            {
                if (isDarkTheme)
                {
                    // Темная тема
                    Background = new SolidColorBrush(Color.FromRgb(30, 30, 30));

                    MainMenu.Background = new SolidColorBrush(Color.FromRgb(45, 45, 45));
                    MainMenu.Foreground = Brushes.White;

                    MainToolBarTray.Background = new SolidColorBrush(Color.FromRgb(60, 60, 60));

                    EditorTabControl.Background = new SolidColorBrush(Color.FromRgb(30, 30, 30));
                    EditorTabControl.Foreground = Brushes.White;

                    foreach (TabItem tab in EditorTabControl.Items)
                    {
                        tab.Background = new SolidColorBrush(Color.FromRgb(45, 45, 45));
                        tab.Foreground = Brushes.White;

                        var grid = tab.Content as Grid;
                        if (grid != null)
                        {
                            var lineNumbersScroll = grid.Children.OfType<ScrollViewer>().FirstOrDefault();
                            if (lineNumbersScroll != null)
                            {
                                lineNumbersScroll.Background = new SolidColorBrush(Color.FromRgb(60, 60, 60));
                                var lineNumbers = lineNumbersScroll.Content as ItemsControl;
                                if (lineNumbers != null)
                                {
                                    lineNumbers.Background = new SolidColorBrush(Color.FromRgb(60, 60, 60));
                                    foreach (var item in lineNumbers.Items)
                                    {
                                        var container = lineNumbers.ItemContainerGenerator.ContainerFromItem(item);
                                        if (container is ContentPresenter presenter)
                                        {
                                            var textBlock = FindVisualChild<TextBlock>(presenter);
                                            if (textBlock != null)
                                                textBlock.Foreground = Brushes.LightGray;
                                        }
                                    }
                                }
                            }

                            var editor = grid.Children.OfType<RichTextBox>().FirstOrDefault();
                            if (editor != null)
                            {
                                editor.Background = new SolidColorBrush(Color.FromRgb(30, 30, 30));
                                editor.Foreground = Brushes.LightGray;
                                editor.BorderBrush = new SolidColorBrush(Color.FromRgb(80, 80, 80));
                            }
                        }
                    }

                    MainSplitter.Background = new SolidColorBrush(Color.FromRgb(80, 80, 80));

                    OutputTabControl.Background = new SolidColorBrush(Color.FromRgb(30, 30, 30));
                    OutputTabControl.Foreground = Brushes.White;

                    foreach (TabItem tab in OutputTabControl.Items)
                    {
                        tab.Background = new SolidColorBrush(Color.FromRgb(45, 45, 45));
                        tab.Foreground = Brushes.White;
                    }

                    ResultsListView.Background = new SolidColorBrush(Color.FromRgb(30, 30, 30));
                    ResultsListView.Foreground = Brushes.LightGray;

                    OutputTextBox.Background = new SolidColorBrush(Color.FromRgb(30, 30, 30));
                    OutputTextBox.Foreground = Brushes.LightGray;
                    OutputTextBox.BorderBrush = new SolidColorBrush(Color.FromRgb(80, 80, 80));

                    ErrorsListView.Background = new SolidColorBrush(Color.FromRgb(30, 30, 30));
                    ErrorsListView.Foreground = Brushes.LightGray;

                    MainStatusBar.Background = new SolidColorBrush(Color.FromRgb(45, 45, 45));
                    MainStatusBar.Foreground = Brushes.White;

                    UpdateMenuItemsColor(Brushes.White);
                }
                else
                {
                    // Светлая тема
                    Background = SystemColors.WindowBrush;

                    MainMenu.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240));
                    MainMenu.Foreground = Brushes.Black;

                    MainToolBarTray.Background = new SolidColorBrush(Color.FromRgb(245, 245, 245));

                    EditorTabControl.Background = SystemColors.WindowBrush;
                    EditorTabControl.Foreground = Brushes.Black;

                    foreach (TabItem tab in EditorTabControl.Items)
                    {
                        tab.Background = SystemColors.WindowBrush;
                        tab.Foreground = Brushes.Black;

                        var grid = tab.Content as Grid;
                        if (grid != null)
                        {
                            var lineNumbersScroll = grid.Children.OfType<ScrollViewer>().FirstOrDefault();
                            if (lineNumbersScroll != null)
                            {
                                lineNumbersScroll.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240));
                                var lineNumbers = lineNumbersScroll.Content as ItemsControl;
                                if (lineNumbers != null)
                                {
                                    lineNumbers.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240));
                                    foreach (var item in lineNumbers.Items)
                                    {
                                        var container = lineNumbers.ItemContainerGenerator.ContainerFromItem(item);
                                        if (container is ContentPresenter presenter)
                                        {
                                            var textBlock = FindVisualChild<TextBlock>(presenter);
                                            if (textBlock != null)
                                                textBlock.Foreground = new SolidColorBrush(Color.FromRgb(102, 102, 102));
                                        }
                                    }
                                }
                            }

                            var editor = grid.Children.OfType<RichTextBox>().FirstOrDefault();
                            if (editor != null)
                            {
                                editor.Background = Brushes.White;
                                editor.Foreground = Brushes.Black;
                                editor.BorderBrush = new SolidColorBrush(Color.FromRgb(204, 204, 204));
                            }
                        }
                    }

                    MainSplitter.Background = new SolidColorBrush(Color.FromRgb(204, 204, 204));

                    OutputTabControl.Background = SystemColors.WindowBrush;
                    OutputTabControl.Foreground = Brushes.Black;

                    foreach (TabItem tab in OutputTabControl.Items)
                    {
                        tab.Background = SystemColors.WindowBrush;
                        tab.Foreground = Brushes.Black;
                    }

                    ResultsListView.Background = Brushes.White;
                    ResultsListView.Foreground = Brushes.Black;

                    OutputTextBox.Background = Brushes.White;
                    OutputTextBox.Foreground = Brushes.Black;
                    OutputTextBox.BorderBrush = new SolidColorBrush(Color.FromRgb(204, 204, 204));

                    ErrorsListView.Background = Brushes.White;
                    ErrorsListView.Foreground = Brushes.Black;

                    MainStatusBar.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240));
                    MainStatusBar.Foreground = Brushes.Black;

                    UpdateMenuItemsColor(Brushes.Black);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error applying theme: {ex.Message}");
            }
        }

        private void UpdateMenuItemsColor(SolidColorBrush color)
        {
            try
            {
                UpdateMenuItemColor(FileMenu, color);
                UpdateMenuItemColor(EditMenu, color);
                UpdateMenuItemColor(ViewMenu, color);
                UpdateMenuItemColor(HelpMenu, color);
            }
            catch { }
        }

        private void UpdateMenuItemColor(MenuItem menuItem, SolidColorBrush color)
        {
            if (menuItem == null) return;

            menuItem.Foreground = color;

            foreach (var item in menuItem.Items)
            {
                if (item is MenuItem subItem)
                {
                    subItem.Foreground = color;
                    UpdateMenuItemColor(subItem, color);
                }
            }
        }

        // ==================== ПЕРЕКЛЮЧЕНИЕ ЯЗЫКА ====================
        private void SetRussianLanguage_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                currentLanguage = "ru-RU";
                RussianLanguageMenuItem.IsChecked = true;
                EnglishLanguageMenuItem.IsChecked = false;
                UpdateUILanguage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при смене языка: {ex.Message}");
            }
        }

        private void SetEnglishLanguage_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                currentLanguage = "en-US";
                RussianLanguageMenuItem.IsChecked = false;
                EnglishLanguageMenuItem.IsChecked = true;
                UpdateUILanguage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error changing language: {ex.Message}");
            }
        }

        private void UpdateUILanguage()
        {
            try
            {
                if (currentLanguage == "ru-RU")
                {
                    Title = "Текстовый редактор с языковым процессором";

                    FileMenu.Header = "_Файл";
                    NewMenuItem.Header = "_Создать";
                    OpenMenuItem.Header = "_Открыть...";
                    SaveMenuItem.Header = "_Сохранить";
                    SaveAsMenuItem.Header = "Сохранить _как...";
                    ExitMenuItem.Header = "_Выход";

                    EditMenu.Header = "_Правка";
                    UndoMenuItem.Header = "_Отменить";
                    RedoMenuItem.Header = "_Повторить";
                    CutMenuItem.Header = "_Вырезать";
                    CopyMenuItem.Header = "К_опировать";
                    PasteMenuItem.Header = "_Вставить";
                    DeleteMenuItem.Header = "_Удалить";
                    SelectAllMenuItem.Header = "Выделить _все";

                    ViewMenu.Header = "_Вид";
                    IncreaseFontMenuItem.Header = "_Увеличить текст";
                    DecreaseFontMenuItem.Header = "_Уменьшить текст";
                    ResetFontMenuItem.Header = "Сбросить _масштаб";
                    ThemeMenu.Header = "_Тема";
                    LightThemeMenuItem.Header = "Светлая";
                    DarkThemeMenuItem.Header = "Темная";
                    LanguageMenu.Header = "_Язык интерфейса";
                    RussianLanguageMenuItem.Header = "Русский";
                    EnglishLanguageMenuItem.Header = "English";

                    HelpMenu.Header = "_Справка";
                    HelpMenuItem.Header = "_Справка";
                    AboutMenuItem.Header = "_О программе";

                    NewButton.ToolTip = "Создать (Ctrl+N)";
                    OpenButton.ToolTip = "Открыть (Ctrl+O)";
                    SaveButton.ToolTip = "Сохранить (Ctrl+S)";
                    CutButton.ToolTip = "Вырезать (Ctrl+X)";
                    CopyButton.ToolTip = "Копировать (Ctrl+C)";
                    PasteButton.ToolTip = "Вставить (Ctrl+V)";
                    UndoButton.ToolTip = "Отменить (Ctrl+Z)";
                    RedoButton.ToolTip = "Повторить (Ctrl+Y)";
                    DeleteButton.ToolTip = "Удалить (Del)";
                    HelpButton.ToolTip = "Справка (F1)";
                    SizeLabel.Content = "Размер:";
                    ZoomInButton.ToolTip = "Увеличить";
                    ZoomOutButton.ToolTip = "Уменьшить";

                    ResultsTab.Header = "Результаты анализа";
                    OutputTab.Header = "Вывод";
                    ErrorsTab.Header = "Ошибки";

                    var resultsView = ResultsListView.View as GridView;
                    if (resultsView != null)
                    {
                        resultsView.Columns[0].Header = "Строка";
                        resultsView.Columns[1].Header = "Тип";
                        resultsView.Columns[2].Header = "Сообщение";
                    }

                    var errorsView = ErrorsListView.View as GridView;
                    if (errorsView != null)
                    {
                        errorsView.Columns[0].Header = "Строка";
                        errorsView.Columns[1].Header = "Столбец";
                        errorsView.Columns[2].Header = "Ошибка";
                    }

                    StatusText.Text = "Готов";
                }
                else
                {
                    Title = "Text Editor with Language Processor";

                    FileMenu.Header = "_File";
                    NewMenuItem.Header = "_New";
                    OpenMenuItem.Header = "_Open...";
                    SaveMenuItem.Header = "_Save";
                    SaveAsMenuItem.Header = "Save _As...";
                    ExitMenuItem.Header = "E_xit";

                    EditMenu.Header = "_Edit";
                    UndoMenuItem.Header = "_Undo";
                    RedoMenuItem.Header = "_Redo";
                    CutMenuItem.Header = "Cu_t";
                    CopyMenuItem.Header = "_Copy";
                    PasteMenuItem.Header = "_Paste";
                    DeleteMenuItem.Header = "_Delete";
                    SelectAllMenuItem.Header = "Select _All";

                    ViewMenu.Header = "_View";
                    IncreaseFontMenuItem.Header = "_Increase Font Size";
                    DecreaseFontMenuItem.Header = "_Decrease Font Size";
                    ResetFontMenuItem.Header = "_Reset Zoom";
                    ThemeMenu.Header = "_Theme";
                    LightThemeMenuItem.Header = "Light";
                    DarkThemeMenuItem.Header = "Dark";
                    LanguageMenu.Header = "_Language";
                    RussianLanguageMenuItem.Header = "Russian";
                    EnglishLanguageMenuItem.Header = "English";

                    HelpMenu.Header = "_Help";
                    HelpMenuItem.Header = "_Help";
                    AboutMenuItem.Header = "_About";

                    NewButton.ToolTip = "New (Ctrl+N)";
                    OpenButton.ToolTip = "Open (Ctrl+O)";
                    SaveButton.ToolTip = "Save (Ctrl+S)";
                    CutButton.ToolTip = "Cut (Ctrl+X)";
                    CopyButton.ToolTip = "Copy (Ctrl+C)";
                    PasteButton.ToolTip = "Paste (Ctrl+V)";
                    UndoButton.ToolTip = "Undo (Ctrl+Z)";
                    RedoButton.ToolTip = "Redo (Ctrl+Y)";
                    DeleteButton.ToolTip = "Delete (Del)";
                    HelpButton.ToolTip = "Help (F1)";
                    SizeLabel.Content = "Size:";
                    ZoomInButton.ToolTip = "Zoom In";
                    ZoomOutButton.ToolTip = "Zoom Out";

                    ResultsTab.Header = "Analysis Results";
                    OutputTab.Header = "Output";
                    ErrorsTab.Header = "Errors";

                    var resultsView = ResultsListView.View as GridView;
                    if (resultsView != null)
                    {
                        resultsView.Columns[0].Header = "Line";
                        resultsView.Columns[1].Header = "Type";
                        resultsView.Columns[2].Header = "Message";
                    }

                    var errorsView = ErrorsListView.View as GridView;
                    if (errorsView != null)
                    {
                        errorsView.Columns[0].Header = "Line";
                        errorsView.Columns[1].Header = "Column";
                        errorsView.Columns[2].Header = "Error";
                    }

                    StatusText.Text = "Ready";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error updating UI language: {ex.Message}");
            }
        }

        // ==================== МЕНЮ ФАЙЛ ====================
        private void NewFile_Click(object sender, RoutedEventArgs e) => NewFile();
        private void OpenFile_Click(object sender, RoutedEventArgs e) => OpenFile();
        private void SaveFile_Click(object sender, RoutedEventArgs e) => SaveFile();
        private void SaveAsFile_Click(object sender, RoutedEventArgs e) => SaveAsFile();

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            try { Close(); } catch { Environment.Exit(0); }
        }

        // ==================== МЕНЮ ПРАВКА ====================
        private void Undo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var editor = GetCurrentEditor();
                if (editor != null && editor.CanUndo)
                    editor.Undo();
            }
            catch { }
        }

        private void Redo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var editor = GetCurrentEditor();
                if (editor != null && editor.CanRedo)
                    editor.Redo();
            }
            catch { }
        }

        private void Cut_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var editor = GetCurrentEditor();
                if (editor != null)
                    editor.Cut();
            }
            catch { }
        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var editor = GetCurrentEditor();
                if (editor != null)
                    editor.Copy();
            }
            catch { }
        }

        private void Paste_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var editor = GetCurrentEditor();
                if (editor != null)
                    editor.Paste();
            }
            catch { }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var editor = GetCurrentEditor();
                if (editor != null)
                {
                    editor.Selection.Text = "";
                }
            }
            catch { }
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var editor = GetCurrentEditor();
                if (editor != null)
                    editor.SelectAll();
            }
            catch { }
        }

        // ==================== МЕНЮ ВИД ====================
        private void IncreaseFontSize_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                currentFontSize = Math.Min(currentFontSize + 2, 48);
                UpdateFontSize();
            }
            catch { }
        }

        private void DecreaseFontSize_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                currentFontSize = Math.Max(currentFontSize - 2, 8);
                UpdateFontSize();
            }
            catch { }
        }

        private void ResetFontSize_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                currentFontSize = 12;
                UpdateFontSize();
            }
            catch { }
        }

        private void FontSizeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (FontSizeCombo.SelectedItem is ComboBoxItem item &&
                    double.TryParse(item.Content.ToString(), out double size))
                {
                    currentFontSize = size;
                    UpdateFontSize();
                }
            }
            catch { }
        }

        private void UpdateFontSize()
        {
            try
            {
                var editor = GetCurrentEditor();
                if (editor != null)
                {
                    editor.FontSize = currentFontSize;
                    UpdateLineNumbersFontSize();
                    ZoomText.Text = $"{(int)((currentFontSize / 12) * 100)}%";

                    foreach (ComboBoxItem item in FontSizeCombo.Items)
                    {
                        if (item.Content.ToString() == currentFontSize.ToString())
                        {
                            FontSizeCombo.SelectedItem = item;
                            break;
                        }
                    }
                }
            }
            catch { }
        }

        // ==================== МЕНЮ СПРАВКА ====================
        private void Help_Click(object sender, RoutedEventArgs e)
        {
            string helpText = currentLanguage == "ru-RU" ?
                "СПРАВКА ПО ПРОГРАММЕ\n\n" +
                "ОСНОВНЫЕ ФУНКЦИИ:\n" +
                "• Создать (Ctrl+N) - создание нового документа\n" +
                "• Открыть (Ctrl+O) - открытие существующего файла\n" +
                "• Сохранить (Ctrl+S) - сохранение текущего документа\n" +
                "• Сохранить как - сохранение под новым именем\n\n" +
                "ПРАВКА:\n" +
                "• Отменить (Ctrl+Z) - отмена последнего действия\n" +
                "• Повторить (Ctrl+Y) - повтор отмененного действия\n" +
                "• Вырезать (Ctrl+X) - вырезание выделенного текста\n" +
                "• Копировать (Ctrl+C) - копирование выделенного текста\n" +
                "• Вставить (Ctrl+V) - вставка текста\n" +
                "• Удалить (Del) - удаление выделенного текста\n" +
                "• Выделить все (Ctrl+A) - выделение всего текста\n\n" +
                "ВИД:\n" +
                "• Увеличить текст (Ctrl++) - увеличение шрифта\n" +
                "• Уменьшить текст (Ctrl+-) - уменьшение шрифта\n" +
                "• Сбросить масштаб (Ctrl+0) - стандартный размер\n" +
                "• Тема - переключение между светлой и темной темой\n\n" +
                "ДОПОЛНИТЕЛЬНО:\n" +
                "• Подсветка синтаксиса для C#\n" +
                "• Нумерация строк\n" +
                "• Анализ кода с выводом ошибок\n" +
                "• Перетаскивание файлов в окно\n" +
                "• Множество вкладок\n" +
                "• Изменение размера шрифта\n" +
                "• Выбор языка интерфейса\n" +
                "• Строка состояния с информацией"
                :
                "HELP\n\n" +
                "MAIN FUNCTIONS:\n" +
                "• New (Ctrl+N) - create new document\n" +
                "• Open (Ctrl+O) - open existing file\n" +
                "• Save (Ctrl+S) - save current document\n" +
                "• Save As - save with new name\n\n" +
                "EDIT:\n" +
                "• Undo (Ctrl+Z) - undo last action\n" +
                "• Redo (Ctrl+Y) - redo undone action\n" +
                "• Cut (Ctrl+X) - cut selected text\n" +
                "• Copy (Ctrl+C) - copy selected text\n" +
                "• Paste (Ctrl+V) - paste text\n" +
                "• Delete (Del) - delete selected text\n" +
                "• Select All (Ctrl+A) - select all text\n\n" +
                "VIEW:\n" +
                "• Increase Font (Ctrl++) - increase font size\n" +
                "• Decrease Font (Ctrl+-) - decrease font size\n" +
                "• Reset Zoom (Ctrl+0) - standard size\n" +
                "• Theme - switch between light and dark theme\n\n" +
                "ADDITIONAL:\n" +
                "• C# syntax highlighting\n" +
                "• Line numbering\n" +
                "• Code analysis with error output\n" +
                "• Drag & drop files\n" +
                "• Multiple tabs\n" +
                "• Font size adjustment\n" +
                "• Language selection\n" +
                "• Status bar with information";

            MessageBox.Show(helpText, currentLanguage == "ru-RU" ? "Справка" : "Help",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void About_Click(object sender, RoutedEventArgs e)
        {
            string aboutText = "ТЕКСТОВЫЙ РЕДАКТОР С ЯЗЫКОВЫМ ПРОЦЕССОРОМ\n\n" +
                              "Автор: Шипунов Д. Н.\n\n" +
                              "Лабораторная работа по дисциплине\n" +
                              "\"Теория формальных языков и компиляторов\"";

            MessageBox.Show(aboutText, "О программе", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // ==================== ОСНОВНЫЕ ФУНКЦИИ ====================
        private void NewFile()
        {
            try
            {
                if (!CheckSaveChanges()) return;

                tabCounter++;
                string tabName = currentLanguage == "ru-RU" ?
                    $"Новый документ {tabCounter}" :
                    $"New Document {tabCounter}";

                var newTab = new TabItem { Header = tabName, Tag = null };
                var grid = new Grid();
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                var lineNumbersScroll = new ScrollViewer
                {
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
                    VerticalAlignment = VerticalAlignment.Stretch
                };
                lineNumbersScroll.SetValue(Grid.ColumnProperty, 0);

                var lineNumbers = new ItemsControl
                {
                    FontFamily = new FontFamily("Consolas"),
                    FontSize = currentFontSize,
                    Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
                    Padding = new Thickness(5, 2, 5, 2)
                };
                lineNumbersScroll.Content = lineNumbers;

                var editor = new RichTextBox
                {
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    FontFamily = new FontFamily("Consolas"),
                    FontSize = currentFontSize,
                    SpellCheck = { IsEnabled = true },
                    BorderThickness = new Thickness(1),
                    BorderBrush = new SolidColorBrush(Color.FromRgb(204, 204, 204))
                };
                editor.TextChanged += EditorTextBox_TextChanged;
                editor.PreviewKeyDown += EditorTextBox_PreviewKeyDown;
                editor.SetValue(Grid.ColumnProperty, 1);

                grid.Children.Add(lineNumbersScroll);
                grid.Children.Add(editor);
                newTab.Content = grid;

                EditorTabControl.Items.Add(newTab);
                EditorTabControl.SelectedItem = newTab;

                UpdateCurrentFileInfo(null, false);
                UpdateLineNumbers();

                ApplyTheme();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании файла: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenFile(string filePath = null)
        {
            try
            {
                if (!CheckSaveChanges()) return;

                if (filePath == null)
                {
                    var dialog = new Microsoft.Win32.OpenFileDialog
                    {
                        Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*",
                        Title = currentLanguage == "ru-RU" ? "Открыть файл" : "Open file"
                    };

                    if (dialog.ShowDialog() != true)
                        return;

                    filePath = dialog.FileName;
                }

                string content = File.ReadAllText(filePath, Encoding.UTF8);
                string tabName = System.IO.Path.GetFileName(filePath);

                var newTab = new TabItem { Header = tabName, Tag = filePath };
                var grid = new Grid();
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                var lineNumbersScroll = new ScrollViewer
                {
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
                    VerticalAlignment = VerticalAlignment.Stretch
                };
                lineNumbersScroll.SetValue(Grid.ColumnProperty, 0);

                var lineNumbers = new ItemsControl
                {
                    FontFamily = new FontFamily("Consolas"),
                    FontSize = currentFontSize,
                    Background = new SolidColorBrush(Color.FromRgb(240, 240, 240)),
                    Padding = new Thickness(5, 2, 5, 2)
                };
                lineNumbersScroll.Content = lineNumbers;

                var editor = new RichTextBox
                {
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                    FontFamily = new FontFamily("Consolas"),
                    FontSize = currentFontSize,
                    SpellCheck = { IsEnabled = true },
                    BorderThickness = new Thickness(1),
                    BorderBrush = new SolidColorBrush(Color.FromRgb(204, 204, 204))
                };
                editor.TextChanged += EditorTextBox_TextChanged;
                editor.PreviewKeyDown += EditorTextBox_PreviewKeyDown;
                editor.SetValue(Grid.ColumnProperty, 1);

                _ignoreTextChanges = true;
                var range = new TextRange(editor.Document.ContentStart, editor.Document.ContentEnd);
                range.Text = content;
                _ignoreTextChanges = false;

                grid.Children.Add(lineNumbersScroll);
                grid.Children.Add(editor);
                newTab.Content = grid;

                EditorTabControl.Items.Add(newTab);
                EditorTabControl.SelectedItem = newTab;

                UpdateCurrentFileInfo(filePath, false);
                UpdateLineNumbers();
                AnalyzeCode(content);

                ApplyTheme();

                StatusText.Text = (currentLanguage == "ru-RU" ? "Файл открыт: " : "File opened: ") + filePath;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии файла: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveFile()
        {
            try
            {
                var currentTab = EditorTabControl.SelectedItem as TabItem;
                if (currentTab == null) return;

                string filePath = currentTab.Tag as string;

                if (string.IsNullOrEmpty(filePath))
                {
                    SaveAsFile();
                    return;
                }

                var grid = currentTab.Content as Grid;
                var editor = grid?.Children.OfType<RichTextBox>().FirstOrDefault();

                if (editor != null)
                {
                    string content = new TextRange(editor.Document.ContentStart, editor.Document.ContentEnd).Text;
                    File.WriteAllText(filePath, content, Encoding.UTF8);

                    currentTab.Header = System.IO.Path.GetFileName(filePath);
                    UpdateCurrentFileInfo(filePath, false);
                    StatusText.Text = (currentLanguage == "ru-RU" ? "Файл сохранен: " : "File saved: ") + filePath;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveAsFile()
        {
            try
            {
                var dialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*",
                    Title = currentLanguage == "ru-RU" ? "Сохранить файл как" : "Save file as"
                };

                if (dialog.ShowDialog() == true)
                {
                    var currentTab = EditorTabControl.SelectedItem as TabItem;
                    if (currentTab != null)
                    {
                        currentTab.Tag = dialog.FileName;
                        SaveFile();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool CheckSaveChanges()
        {
            try
            {
                if (!isModified) return true;

                var result = MessageBox.Show(
                    currentLanguage == "ru-RU" ? "Сохранить изменения в файле?" : "Save changes to file?",
                    currentLanguage == "ru-RU" ? "Сохранение" : "Save",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Question);

                switch (result)
                {
                    case MessageBoxResult.Yes:
                        SaveFile();
                        return true;
                    case MessageBoxResult.No:
                        return true;
                    case MessageBoxResult.Cancel:
                        return false;
                    default:
                        return false;
                }
            }
            catch
            {
                return true;
            }
        }

        private void UpdateCurrentFileInfo(string filePath, bool modified)
        {
            try
            {
                currentFilePath = filePath;
                isModified = modified;

                var currentTab = EditorTabControl.SelectedItem as TabItem;
                if (currentTab != null && !string.IsNullOrEmpty(filePath))
                {
                    string fileName = System.IO.Path.GetFileName(filePath);
                    currentTab.Header = modified ? fileName + "*" : fileName;
                }

                UpdateTitle();
            }
            catch { }
        }

        private void UpdateTitle()
        {
            try
            {
                var currentTab = EditorTabControl.SelectedItem as TabItem;
                string fileName = currentLanguage == "ru-RU" ? "Новый документ" : "New document";

                if (currentTab != null)
                {
                    fileName = currentTab.Header.ToString().Replace("*", "");
                }

                Title = $"{(currentLanguage == "ru-RU" ? "Текстовый редактор" : "Text Editor")} - {fileName}{(isModified ? "*" : "")}";
            }
            catch { }
        }

        private void AnalyzeCode(string text)
        {
            try
            {
                Results.Clear();
                Errors.Clear();
                OutputTextBox.Clear();

                if (string.IsNullOrWhiteSpace(text))
                {
                    OutputTextBox.Text = currentLanguage == "ru-RU" ? "Нет текста для анализа." : "No text to analyze.";
                    return;
                }

                var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];

                    if (string.IsNullOrWhiteSpace(line))
                    {
                        Results.Add(new ResultItem
                        {
                            Line = i + 1,
                            Type = currentLanguage == "ru-RU" ? "Информация" : "Information",
                            Message = currentLanguage == "ru-RU" ? "Пустая строка" : "Empty line"
                        });
                        continue;
                    }

                    if (line.Length > 100)
                    {
                        Errors.Add(new ErrorItem
                        {
                            Line = i + 1,
                            Column = 101,
                            Message = currentLanguage == "ru-RU" ?
                                "Строка слишком длинная (более 100 символов)" :
                                "Line too long (more than 100 characters)"
                        });
                    }

                    if (line.Contains("\t"))
                    {
                        Results.Add(new ResultItem
                        {
                            Line = i + 1,
                            Type = currentLanguage == "ru-RU" ? "Предупреждение" : "Warning",
                            Message = currentLanguage == "ru-RU" ? "Обнаружен символ табуляции" : "Tab character detected"
                        });
                    }

                    if (Regex.IsMatch(line, @"[^\x20-\x7E\u0400-\u04FF]"))
                    {
                        Errors.Add(new ErrorItem
                        {
                            Line = i + 1,
                            Column = 0,
                            Message = currentLanguage == "ru-RU" ?
                                "Обнаружены недопустимые символы" :
                                "Invalid characters detected"
                        });
                    }
                }

                SyntaxHighlight();

                OutputTextBox.Text = (currentLanguage == "ru-RU" ?
                    $"Анализ завершен.\nВсего строк: {lines.Length}\nРезультатов: {Results.Count}\nОшибок: {Errors.Count}" :
                    $"Analysis complete.\nTotal lines: {lines.Length}\nResults: {Results.Count}\nErrors: {Errors.Count}");

                StatusText.Text = (currentLanguage == "ru-RU" ?
                    $"Анализ завершен. Найдено ошибок: {Errors.Count}" :
                    $"Analysis complete. Errors found: {Errors.Count}");
            }
            catch (Exception ex)
            {
                OutputTextBox.Text = $"Ошибка при анализе: {ex.Message}";
            }
        }

        private void SyntaxHighlight()
        {
            try
            {
                var editor = GetCurrentEditor();
                if (editor == null) return;

                string[] keywords = {
                    "class", "public", "private", "protected", "static", "void",
                    "int", "string", "bool", "double", "float", "if", "else",
                    "for", "foreach", "while", "do", "switch", "case", "break",
                    "continue", "return", "new", "using", "namespace", "try",
                    "catch", "finally", "throw", "this", "base", "true", "false",
                    "null", "var", "async", "await", "task"
                };

                TextPointer pointer = editor.Document.ContentStart;
                while (pointer != null)
                {
                    if (pointer.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
                    {
                        string textRun = pointer.GetTextInRun(LogicalDirection.Forward);
                        if (!string.IsNullOrEmpty(textRun))
                        {
                            var words = Regex.Matches(textRun, @"\b\w+\b");
                            foreach (Match word in words)
                            {
                                TextPointer start = pointer.GetPositionAtOffset(word.Index);
                                TextPointer end = pointer.GetPositionAtOffset(word.Index + word.Length);

                                if (start != null && end != null)
                                {
                                    var wordRange = new TextRange(start, end);
                                    if (keywords.Contains(word.Value))
                                    {
                                        wordRange.ApplyPropertyValue(TextElement.ForegroundProperty, isDarkTheme ? Brushes.Cyan : Brushes.Blue);
                                        wordRange.ApplyPropertyValue(TextElement.FontWeightProperty, FontWeights.Bold);
                                    }
                                    else if (Regex.IsMatch(word.Value, @"^\d+$"))
                                    {
                                        wordRange.ApplyPropertyValue(TextElement.ForegroundProperty, isDarkTheme ? Brushes.LightGreen : Brushes.Green);
                                    }
                                }
                            }
                        }
                    }
                    pointer = pointer.GetNextContextPosition(LogicalDirection.Forward);
                }
            }
            catch { }
        }

        // ==================== НУМЕРАЦИЯ СТРОК ====================
        private void UpdateLineNumbers()
        {
            try
            {
                if (_isUpdating) return;
                _isUpdating = true;

                var currentTab = EditorTabControl.SelectedItem as TabItem;
                if (currentTab == null)
                {
                    _isUpdating = false;
                    return;
                }

                var grid = currentTab.Content as Grid;
                var lineNumbersScroll = grid?.Children.OfType<ScrollViewer>().FirstOrDefault();
                var lineNumbers = lineNumbersScroll?.Content as ItemsControl;
                var editor = grid?.Children.OfType<RichTextBox>().FirstOrDefault();

                if (lineNumbers == null || editor == null)
                {
                    _isUpdating = false;
                    return;
                }

                string text = new TextRange(editor.Document.ContentStart, editor.Document.ContentEnd).Text;

                int lineCount = 1;
                foreach (char c in text)
                {
                    if (c == '\n') lineCount++;
                }

                var numbers = new List<string>();
                for (int i = 1; i <= lineCount; i++)
                {
                    numbers.Add(i.ToString());
                }

                lineNumbers.ItemsSource = numbers;
                _isUpdating = false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка обновления нумерации: {ex.Message}");
                _isUpdating = false;
            }
        }

        private void UpdateLineNumbersFontSize()
        {
            try
            {
                var currentTab = EditorTabControl.SelectedItem as TabItem;
                if (currentTab == null) return;

                var grid = currentTab.Content as Grid;
                var lineNumbersScroll = grid?.Children.OfType<ScrollViewer>().FirstOrDefault();
                var lineNumbers = lineNumbersScroll?.Content as ItemsControl;

                if (lineNumbers != null)
                {
                    lineNumbers.FontSize = currentFontSize;
                }
            }
            catch { }
        }

        private void UpdateStatusBar()
        {
            try
            {
                var editor = GetCurrentEditor();
                if (editor == null) return;

                TextPointer caretPos = editor.CaretPosition;

                int line = 1;
                int column = 1;

                TextPointer pointer = editor.Document.ContentStart;
                while (pointer != null && pointer.CompareTo(caretPos) < 0)
                {
                    if (pointer.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
                    {
                        string text = pointer.GetTextInRun(LogicalDirection.Forward);
                        column += text.Length;
                    }

                    if (pointer.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.ElementStart &&
                        pointer.GetAdjacentElement(LogicalDirection.Forward) is Paragraph)
                    {
                        line++;
                        column = 1;
                    }

                    pointer = pointer.GetNextContextPosition(LogicalDirection.Forward);
                }

                LineColumnText.Text = $"{(currentLanguage == "ru-RU" ? "Стр" : "Ln")}: {line} {(currentLanguage == "ru-RU" ? "Стлб" : "Col")}: {column}";
            }
            catch { }
        }

        private void EditorTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (isClosing || _ignoreTextChanges || _isUpdating) return;

                isModified = true;
                UpdateTitle();
                UpdateLineNumbers();

                var editor = sender as RichTextBox;
                if (editor != null && EditorTabControl.SelectedItem != null)
                {
                    if (_analysisTimer != null)
                    {
                        _analysisTimer.Stop();
                        _analysisTimer = null;
                    }

                    _analysisTimer = new DispatcherTimer
                    {
                        Interval = TimeSpan.FromMilliseconds(500)
                    };
                    _analysisTimer.Tick += (s, args) =>
                    {
                        try
                        {
                            if (!isClosing)
                            {
                                string text = new TextRange(editor.Document.ContentStart, editor.Document.ContentEnd).Text;
                                AnalyzeCode(text);
                            }
                            _analysisTimer?.Stop();
                            _analysisTimer = null;
                        }
                        catch { }
                    };
                    _analysisTimer.Start();
                }
            }
            catch { }
        }

        private void EditorTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl ||
                    e.Key == Key.LeftShift || e.Key == Key.RightShift ||
                    e.Key == Key.LeftAlt || e.Key == Key.RightAlt ||
                    e.Key == Key.CapsLock || e.Key == Key.NumLock ||
                    e.Key == Key.Scroll || e.Key == Key.PrintScreen ||
                    e.Key == Key.Pause)
                {
                    return;
                }

                if (Keyboard.Modifiers == ModifierKeys.Control)
                {
                    switch (e.Key)
                    {
                        case Key.N:
                            e.Handled = true;
                            NewFile();
                            break;
                        case Key.O:
                            e.Handled = true;
                            OpenFile();
                            break;
                        case Key.S:
                            e.Handled = true;
                            SaveFile();
                            break;
                        case Key.Z:
                            e.Handled = true;
                            Undo_Click(sender, e);
                            break;
                        case Key.Y:
                            e.Handled = true;
                            Redo_Click(sender, e);
                            break;
                        case Key.X:
                            e.Handled = true;
                            Cut_Click(sender, e);
                            break;
                        case Key.C:
                            e.Handled = true;
                            Copy_Click(sender, e);
                            break;
                        case Key.V:
                            e.Handled = true;
                            Paste_Click(sender, e);
                            break;
                        case Key.A:
                            e.Handled = true;
                            SelectAll_Click(sender, e);
                            break;
                        case Key.OemPlus:
                            e.Handled = true;
                            IncreaseFontSize_Click(sender, e);
                            break;
                        case Key.OemMinus:
                            e.Handled = true;
                            DecreaseFontSize_Click(sender, e);
                            break;
                        case Key.D0:
                            e.Handled = true;
                            ResetFontSize_Click(sender, e);
                            break;
                    }
                }
                else if (e.Key == Key.F1)
                {
                    e.Handled = true;
                    Help_Click(sender, e);
                }
                else if (e.Key == Key.Delete)
                {
                    Delete_Click(sender, e);
                }
            }
            catch { }
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    if (files.Length > 0 && File.Exists(files[0]))
                    {
                        OpenFile(files[0]);
                    }
                }
            }
            catch { }
        }

        private void ResultsListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (ResultsListView.SelectedItem is ResultItem selected)
                {
                    var editor = GetCurrentEditor();
                    if (editor != null)
                    {
                        TextPointer position = editor.Document.ContentStart;
                        for (int i = 0; i < selected.Line - 1; i++)
                        {
                            position = position.GetLineStartPosition(1);
                            if (position == null) break;
                        }

                        if (position != null)
                        {
                            editor.CaretPosition = position;
                            editor.Focus();

                            TextRange lineRange = new TextRange(
                                position.GetLineStartPosition(0),
                                position.GetLineStartPosition(1));
                            lineRange.ApplyPropertyValue(TextElement.BackgroundProperty, Brushes.Yellow);

                            var timer = new DispatcherTimer
                            {
                                Interval = TimeSpan.FromSeconds(1)
                            };
                            timer.Tick += (s, args) =>
                            {
                                try
                                {
                                    lineRange.ApplyPropertyValue(TextElement.BackgroundProperty, null);
                                    timer.Stop();
                                }
                                catch { }
                            };
                            timer.Start();
                        }
                    }
                }
            }
            catch { }
        }

        private RichTextBox GetCurrentEditor()
        {
            try
            {
                var currentTab = EditorTabControl.SelectedItem as TabItem;
                if (currentTab == null) return null;

                var grid = currentTab.Content as Grid;
                return grid?.Children.OfType<RichTextBox>().FirstOrDefault();
            }
            catch
            {
                return null;
            }
        }
    }

    public static class RichTextBoxExtensions
    {
        public static void AppendText(this RichTextBox box, string text)
        {
            try
            {
                var range = new TextRange(box.Document.ContentEnd, box.Document.ContentEnd);
                range.Text = text;
            }
            catch { }
        }
    }
}