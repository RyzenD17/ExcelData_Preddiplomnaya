using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using ExcelDataReader;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;

using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;
using Page = System.Windows.Controls.Page;
using Word = Microsoft.Office.Interop.Word;
using Task = System.Threading.Tasks.Task;
using MailMessage = System.Net.Mail.MailMessage;
using System.Net.Mime;
using System.Windows.Resources;
using ExcelData.Properties;

namespace ExcelData
{
    /// <summary>
    /// Логика взаимодействия для MainPage.xaml
    /// </summary>
    public partial class MainPage : Page
    {
        
        IExcelDataReader edr;
        public DataTableCollection tableCollection = null;
        public string filePath = "";
        public string fileName = "";
        DataView dataView;

        //флаг переключения режима - таблица или студенты
        public bool isTableShow = false;

        //список студентов (неожиданно)
        public List<Tuple<string, int>> students = new List<Tuple<string, int>>();

        public MainPage()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Открываем диалог с выором файла
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenExcelbtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "EXCEL Files (*.xlsx)|*.xlsx|EXCEL Files 2003 (*.xls)|*.xls|All files (*.*)|*.*";
                if (openFileDialog.ShowDialog() != true) return;
                //берем путь файла
                filePath = openFileDialog.FileName;
                //и его имя
                fileName = openFileDialog.SafeFileName;
                //сохраняем путь в настройки
                Properties.Settings.Default.FilePath = filePath;
                Properties.Settings.Default.Save(); 
                //читаем файл
                readFile(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// Функция рассчета пропусков и всег ос этим связанного
        /// </summary>
        /// <param name="month">месяц за который считать</param>
        public void CalculateSkips(string month = "")
        {
            try
            {
                //все пропуски студента
                int pass = 0;
                //все обоснованные пропуски
                int passOk = 0;
                //итог всех пропусков
                int summPass = 0;
                //итог всех пропусков по неуважительной причине
                int summPassNotOk = 0;
                string passUvString = "";
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = excel.Workbooks.Open(filePath);
                Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[month == "" ? CBChooseList.SelectedItem.ToString() : month];

                //тукущий месяц который выбран для подсчета если не указан в аргументах берем из комбо бокса
                var currentMonth = month == "" ? CBChooseList.SelectedItem.ToString() : month;
                //очищаем список прошлых пропусков чтобы заполнить новыми
                var oldData = YearlySkips.Data.Where(a => a.month == currentMonth).ToList();
                if (oldData.Count > 0)
                {
                    foreach (var item in oldData)
                    {
                        YearlySkips.Data.Remove(item);
                    }
                }
                //заполняем дада вью месяцем если он задан убираем первые 2 строки так как они заголовки
                if (month != "")
                {
                    dataView = new DataView(tableCollection[month].AsDataView().ToTable());
                    dataView.Table.Rows.RemoveAt(0);
                    dataView.Table.Rows.RemoveAt(0);

                }
                //делаем посчит всех ячеек и пропусков по ним из дата вьб
                for (int i = 0; i < dataView.Table.Rows.Count - 1; i++)
                {
                    for (int j = 2; j < dataView.Table.Columns.Count - 2; j++)
                    {
                        //текущая ячейка
                        var b = dataView.Table.Rows[i][j].ToString();

                        if (b == "2" || b == "4" || b == "6" || b == "8")
                        {
                            pass += Convert.ToInt32(b);
                        }
                        if (b == "-2" || b == "-4" || b == "-6" || b == "-8")
                        {
                            passUvString = b.Replace("-", "");
                            pass += Convert.ToInt32(passUvString);
                            passOk += Convert.ToInt32(passUvString);
                        }
                    }
                    //заполняем итоги студента
                    sheet1.Cells[i + 3, dataView.Table.Columns.Count - 1] = pass.ToString();
                    sheet1.Cells[i + 3, dataView.Table.Columns.Count] = (pass - passOk).ToString();
                    summPass += pass;
                    summPassNotOk += pass - passOk;

                    //заполняем итоги общие
                    sheet1.Cells[dataView.Table.Rows.Count + 3 - 1, dataView.Table.Columns.Count - 1] = summPass.ToString();
                    sheet1.Cells[dataView.Table.Rows.Count + 3 - 1, dataView.Table.Columns.Count] = summPassNotOk.ToString();

                    //добавляем инфу о пропуске в глобальные список
                    YearlySkips.Data.Add((dataView.Table.Rows[i][1].ToString(), currentMonth, pass, passOk, pass - passOk));

                    //обнуляем переменные для студента
                    passOk = 0;
                    pass = 0;
                }
                //сохраняем выходим
                workbook.Save();
                workbook.Close(false);
                excel.Quit();
                excel = null;
                workbook = null;
                sheet1 = null;
                //читаем файл заново чтобы обновить данные на новые
                readFile(filePath, true);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
        /// <summary>
        /// Полный пересчет всего года или только выбранных месяцев
        /// </summary>
        /// <param name="months">список месяцев для пересчета</param>
        public void CalculateFullYear(List<string> months = null)
        {
            for (int i = 0; i < (months != null ? months.Count : CBChooseList.Items.Count); i++)
            {
                CalculateSkips(months != null ? months[i] : CBChooseList.Items[i].ToString());

            }
            //обновляем дата вью устанавливая текущее положение чтобы не сбилось
            dataView = new DataView(tableCollection[CBChooseList.SelectedItem.ToString()].AsDataView().ToTable());

            dataView.Table.Rows.RemoveAt(0);
            dataView.Table.Rows.RemoveAt(0);
        }

        /// <summary>
        /// Функция считывания файла с таблицами
        /// </summary>
        /// <param name="filePath">путь к файлу</param>
        /// <param name="isUpdate">флаг обновляем ли мы данные или заново читаем файл</param>
        public void readFile(string filePath, bool isUpdate = false)
        {
            try
            {
                //расщерение файла
                var extension = filePath.Substring(filePath.LastIndexOf('.'));

                FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
                // Читатель для файлов с расширением *.xlsx.
                if (extension == ".xlsx")
                    edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
                // Читатель для файлов с расширением *.xls.
                else if (extension == ".xls")
                    edr = ExcelReaderFactory.CreateBinaryReader(stream);
                //// reader.IsFirstRowAsColumnNames

                DataSet dataSet = edr.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = false
                    }
                });

                tableCollection = dataSet.Tables;

                //ели просто обновляем то заново устанавливаем значение комбо бокса чтобы сработало событие для него лиюо заново заполнием комбо бокс
                if (isUpdate)
                {
                    var currentSelection = CBChooseList.SelectedIndex;
                    CBChooseList.SelectedIndex = -1;
                    CBChooseList.SelectedIndex = currentSelection;
                }
                else
                {
                    CBChooseList.Items.Clear();

                    foreach (DataTable dt in tableCollection)
                    {
                        CBChooseList.Items.Add(dt.TableName);
                    }

                    CBChooseList.SelectedIndex = 0;
                }

                edr.Close();
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        /// <summary>
        /// Заполнение всего дата грида и студентов основываясь на положении комбо бокса
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CBChooseList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
            if (e.AddedItems.Count <= 0) return;
            if (CBChooseList.SelectedIndex < 0) return;
            
            try
            {
                //очищаем все
                DbGrig.ItemsSource = null;
                students.Clear();

                //получаем таблицу на выбранный месяц и убираем сохраняем первые 2 строки так как будет использовать их как заголовки
                dataView = new DataView(tableCollection[CBChooseList.SelectedItem.ToString()].AsDataView().ToTable());
                var r = Array.ConvertAll(dataView.Table.Rows[0].ItemArray, a => a.ToString());
                var r1 = Array.ConvertAll(dataView.Table.Rows[1].ItemArray, a => a.ToString());

                //после удаления удалем их - они не нужны(
                dataView.Table.Rows.RemoveAt(0);
                dataView.Table.Rows.RemoveAt(0);

                //заполняем колонки таблицы
                DbGrig.Columns.Clear();
                for (int i = 0; i < dataView.Table.Columns.Count; i++)
                {
                    DbGrig.Columns.Add(new DataGridTextColumn() { Binding = new Binding(dataView.Table.Columns[i].ColumnName) });
                }
                DbGrig.ItemsSource = dataView;
                //настраиваем колонки таблицы
                for (int i = 0; i < DbGrig.Columns.Count; i++)
                {
                    DbGrig.Columns[i].Width = DataGridLength.Auto;
                    DbGrig.Columns[i].Header = r1[i].ToString();
                }
                //задаем заголовок группы над списком студентов
                txtHeader.Content = r[0].ToString();

                //заполняем список студентов
                for (int i = 0; i < dataView.Table.Rows.Count; i++)
                {
                    if (dataView.Table.Rows[i].ItemArray[0].ToString() == "")
                    {
                        continue;
                    }
                    students.Add(new Tuple<string, int>(dataView.Table.Rows[i].ItemArray[1].ToString(), 0));
                }
                if (!isTableShow)
                {
                    stackStudents.Visibility = Visibility.Visible;
                }
                list.DataContext = students;
                list.ItemsSource = students;
                txtGroup.Text = System.IO.Path.GetFileNameWithoutExtension(fileName);
            }
            catch { return; }
            

        }
        /// <summary>
        /// Переход на форму заполнения пропусков
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAddSkip_Click(object sender, RoutedEventArgs e)
        {
            if (list.SelectedItems.Count <= 0)
            {
                MessageBox.Show("Сначала выберите студента!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            //формируем студента которого будем редактировать и отправляем на новую форму
            var s = new StudentsData();
            var fio = (Tuple<string, int>)list.SelectedItem;
            s.FIO = fio.Item1.ToString();
            new AddSkipWindow(s, dataView).Show();
        }

        /// <summary>
        /// Переключение режимов просмотра студентов и общей таблицы пропусков для всей группы на месяц
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFullScreen_Click(object sender, RoutedEventArgs e)
        {
            if (tableCollection == null)
            {
                MessageBox.Show("Сначала загрузите файл!", "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            isTableShow = !isTableShow;

            stackTable.Visibility = isTableShow ? Visibility.Visible : Visibility.Collapsed;
            stackStudents.Visibility = isTableShow ? Visibility.Collapsed : Visibility.Visible;
            stackSideBar.Visibility = isTableShow ? Visibility.Collapsed : Visibility.Visible;
            stackBottomBar.Visibility = isTableShow ? Visibility.Visible: Visibility.Collapsed;

            UpdateGroupColors();
        }
        /// <summary>
        /// Открытие справки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAbout_Click(object sender, RoutedEventArgs e)
        {
            new AboutWindow().Show();
        }
        private static async Task SendEmailWordAsync(string txtTo, string path, string body = "")
        {
            var toS = txtTo;
            //smtp почта / логин
            var fromS = Properties.Settings.Default.EmailSmtp;
            //пароль от smtp 
            var pass = Properties.Settings.Default.PasswordSmtp;

            //формируем письмо
            MailAddress from = new MailAddress(fromS, "Классный руководитель");
            MailAddress to = new MailAddress(toS);
            MailMessage m = new MailMessage(from, to);


            //тема
            m.Subject = "Отчет Word от приложения Классный руководитель";
            //тело
            m.Body = body;

            m.Attachments.Add(new Attachment(path, MediaTypeNames.Application.Octet));

            //создаем подключение smtp и отправляем письмо
            SmtpClient smtp = new SmtpClient("smtp.mail.ru", 587);
            smtp.UseDefaultCredentials = true;
            smtp.Credentials = new NetworkCredential(fromS, pass);
            smtp.EnableSsl = true;
            //ожидаем отправки
            await smtp.SendMailAsync(m);

            Console.WriteLine("Письмо отправлено");
        }
        /// <summary>
        /// Функция формирования ворд отчета по месяцам основываясь на эксель таблице
        /// </summary>
        /// <param name="path">путь к таблице</param>
        /// <param name="path2">путь к ворд файлу куда сохранять</param>
        /// <param name="bottomRightCoords">координаты самого нижнего правого угла, нужно для выделения всей таблицы</param>
        /// <param name="isWordExport">флаг отвечающий делаем мы отчет по месяцам либо экспорт текущего месяца таблицы дата грида</param>
        public async void MonthsReportWord(string path, string path2, (int r, int c) bottomRightCoords, bool isWordExport = true)
        {
            //создаем эксель
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbooks workbooks = xlApp.Workbooks;
            Excel.Workbook wb = workbooks.Open(path);
            Excel.Worksheet worksheet = wb.Sheets[1];

            //создаем ворд
            Word.Application wdApp = new Word.Application();
            Word.Document document = wdApp.Documents.Add();

            //меняем ориентацию листа и отступы
            document.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            document.PageSetup.LeftMargin = 10;
            document.PageSetup.RightMargin = 10;

            //задаем авто колонки
            worksheet.Columns.AutoFit();

            //задаем ширину колонок в записимости от флага чтобы все поместилось в документе
            if (isWordExport)
            {
                //все колонки
                worksheet.Columns.ColumnWidth = 3;

                //колонка студентов и итогов
                worksheet.Columns[2].ColumnWidth = 15;
                worksheet.Columns[bottomRightCoords.c].ColumnWidth = 5;
                worksheet.Columns[bottomRightCoords.c - 1].ColumnWidth = 5;
            }
            else
            {
                //все колонки
                worksheet.Columns.ColumnWidth = 7;

                //колонка студентов
                worksheet.Columns[1].ColumnWidth = 8;
            } 

            //2 ячейки которые формируют зону выделения которую мы будет копировать в ворд документ
            var topLeft = worksheet.Cells[1, 1];
            var bottomRight = worksheet.Cells[bottomRightCoords.r, bottomRightCoords.c];

            //выделяем зону
            dynamic range = worksheet.Range[topLeft, bottomRight];

            //копируем
            range.Copy();

            //в зависимости от режима вставляем в документ с разными аргументами чтобы влезла вся таблица
            if (isWordExport)
            {
                document.Range().PasteExcelTable(false, true, true);
            }
            else
            {
                document.Range().PasteExcelTable(false, true, false);
                
            }

            //выделяем весь документ
            var start = document.Content.Start;
            var end = document.Content.End;
            var docRange = document.Range(start, end);

            //и меняем размер текста в нем длякаждого режима свой
            if (isWordExport)
            {
                docRange.Font.Size = 8;
            }
            else
            {
                docRange.Font.Size = 6;
            }

            //сохраняем документ ворда и закрываем
            document.SaveAs2(path2);
            document.Close();

            //сохраняем таблицу закрываем и выходим из всего
            wb.Save();
            wb.Close();
            xlApp.Quit();
            wdApp.Quit();

            //удаляем временный файл экселя и запускаем GC
            File.Delete(path);
            GC.Collect();

            if (MessageBox.Show("Хотите отправить файл на почту?", "Экспорт", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                
                new AskEmail().ShowDialog();
                await SendEmailWordAsync(Settings.Default.EmailUser, path2, "Ваш отчет Word экспортирован!");
                
            }


        }

        /// <summary>
        /// Фуекция формирования отчета за год либо за выбранные месяцы
        /// </summary>
        /// <param name="months">месяцы для отчета</param>
        public void YearReport(List<string> months = null)
        {
            Excel.Application xlApp = new Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            //путь на рабочий стол
            var path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            //отпределение переменной имени временного файла либо как временный файл либо как годовой отчет
            string reportFileName = "";
            if (months != null)
            {
                //делаем уникальное имя чтобы при сохранении случайно не перезаписалось что то если есть такое имя
                reportFileName = $"{Guid.NewGuid()}.xlsx";
            }
            else
            {
                
                reportFileName = $"Годовой отчет {System.IO.Path.GetFileNameWithoutExtension(fileName)}.xlsx";
            }
            //считаем строки
            var totalRow = students.Count + 2;

            //просто нул значение для экселя
            object misValue = System.Reflection.Missing.Value;

            //переменная для нижней правой ячейки 
            (int r, int c) bottomRightCoords = (0, 0);

            //сичтаем сколько месяцев
            var monthsCount = (months != null ? months.Count : CBChooseList.Items.Count);

            Excel.Workbooks xlWorkBooks;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Sheets xlWorkSheets;

            //выключаем все возможные предепреждения экселя
            xlApp.DisplayAlerts = false;

            xlWorkBooks = xlApp.Workbooks;
            xlWorkBook = xlWorkBooks.Add(misValue);

            xlWorkSheets = xlWorkBook.Worksheets;
            xlWorkSheet = xlWorkSheets.get_Item(1);

            //делаем пересчет выбранных месяцев
            CalculateFullYear(months);

            //заполнения таблицы экселя студентами и их данными о пропуске из глобального списка о пропусках
            for (int i = 0; i < students.Count; i++)
            {
                int stRow = i + 2;

                //заполняем ячуйку студентом
                xlWorkSheet.Cells[stRow, 1] = students[i].Item1;

                //изем все пропуски для него
                var currStudentData = YearlySkips.Data.Where(a => a.student == students[i].Item1).ToList();

                //для кождого пропуска и месяца делаем заполнение в таблицу
                for (int j = 0; j < currStudentData.Count; j++)
                {
                    for (int k = 0; k < CBChooseList.Items.Count; k++)
                    {
                        var currMonth = CBChooseList.Items[k].ToString();
                        var currData = currStudentData[j];

                        if (currData.month != currMonth) continue;

                        //заполняем ячейки
                        xlWorkSheet.Cells[stRow, 2 + ((j * 3) + 0)] = currData.allSkips;
                        xlWorkSheet.Cells[stRow, 2 + ((j * 3) + 1)] = currData.okSkips;
                        xlWorkSheet.Cells[stRow, 2 + ((j * 3) + 2)] = currData.notOkSkips;

                    }
                }
            }

            //заполнияем строку итого
            xlWorkSheet.Cells[totalRow, 1] = "ИТОГО:";


            for (int k = 0; k < monthsCount; k++)
            {
                var currMonth = months != null ? months[k] : CBChooseList.Items[k].ToString();

                //суммируем все пропуски студентов и находим итоги по месяцам
                var allSkip = YearlySkips.Data.Where(t => t.month == currMonth).Sum(t => t.allSkips);
                var okSkip = YearlySkips.Data.Where(t => t.month == currMonth).Sum(t => t.okSkips);
                var badSkip = YearlySkips.Data.Where(t => t.month == currMonth).Sum(t => t.notOkSkips);

                //заполняем иготи в ячейки
                xlWorkSheet.Cells[totalRow, 2 + ((k * 3) + 0)] = allSkip;
                xlWorkSheet.Cells[totalRow, 2 + ((k * 3) + 1)] = okSkip;
                xlWorkSheet.Cells[totalRow, 2 + ((k * 3) + 2)] = badSkip;

                //заполняем загаловки дял месяцев делаем центрирование и объединяем ячейки
                xlWorkSheet.Cells[1, 2 + (k * 3)] = months != null ? months[k] : CBChooseList.Items[k].ToString();
                xlWorkSheet.Cells[1, 2 + (k * 3)].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Range[xlWorkSheet.Cells[1, 2 + (k * 3)], xlWorkSheet.Cells[1, 2 + (k * 3) + 2]].Merge();

                //если последний цикл цикла ищем самую угловую ячейку - тогововые значения пропусков по н/у причине т.к она самая крайняя 
                if (k == monthsCount - 1)
                {
                    bottomRightCoords = (totalRow, 2 + ((k * 3) + 2));
                }
            }

            xlWorkSheet.Cells[1, 1] = txtGroup.Text;

            //сохраняем выходим и обнулчем все что создали
            xlWorkBook.SaveAs(System.IO.Path.Combine(path, reportFileName));
            xlWorkBook.Close(false);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkSheets);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkBooks);
            Marshal.ReleaseComObject(xlApp);

            xlApp = null;
            xlWorkBooks = null;
            xlWorkBook = null;
            xlWorkSheet = null;
            xlWorkSheets = null;

            //если у нас указаны месяцы значит мы делаем отчет по месяцам а значит надо делать ворд файл
            if (months != null)
            {
                var wordFileName = $"Отчет по месяцам {txtGroup.Text}.docx";
                MonthsReportWord(System.IO.Path.Combine(path, reportFileName), System.IO.Path.Combine(path, wordFileName), bottomRightCoords, false);
            }

            GC.Collect();
        }
        /// <summary>
        /// Экспорт таблицы дата грид в ворд файл
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSaveWord_Click(object sender, RoutedEventArgs e)
        {
            var path = "";
            var dialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            var mainWin = (MainWindow)App.Current.Windows[0];

            //открываем окно выбора папки сохранения
            if (dialog.ShowDialog(mainWin).GetValueOrDefault())
            {
                path = dialog.SelectedPath;
            }
            else
            {
                return;
            }

            var fileWordName = $"Отчет за {dataView.Table.TableName} {txtGroup.Text}.docx";

            //для создания будем использовать уже имеющуюся таблицу экселя - просто удалим все ненужные листы из него и соханим в ворд
            var fullPath = System.IO.Path.Combine(path, $"{Guid.NewGuid()}.xlsx");

            //копируем файл в новое место
            File.Copy(filePath, fullPath, true);

            Excel.Application excel = new Excel.Application();
            Excel.Workbooks workbooks = excel.Workbooks;
            Excel.Workbook wb = workbooks.Open(fullPath);
            Sheets sheets = wb.Sheets;

            //удаляем все листы кроме нужной
            for (int i = sheets.Count; i > 0; i--)
            {
                if (sheets[i].Name != dataView.Table.TableName)
                {
                    excel.DisplayAlerts = false;
                    sheets[i].Delete();
                    excel.DisplayAlerts = true;
                }
            }

            //ищем нижный край
            (int r, int c) bottomRightCoords = (dataView.Table.Rows.Count + 2, dataView.Table.Columns.Count);

            wb.Save();
            wb.Close(true);
            excel.Quit();

            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(sheets);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(excel);

            wb = null;
            workbooks = null;
            excel = null;

            GC.Collect();

            //создаепм ворд документ
            MonthsReportWord(fullPath, System.IO.Path.Combine(path, fileWordName), bottomRightCoords);
            MessageBox.Show("Отчет создан!");

        }

        /// <summary>
        /// создание экспорта в эксель - делаем то же самое что и выще только без ворда
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSaveExcel_Click(object sender, RoutedEventArgs e)
        {
            var path = "";
            var dialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            var mainWin = (MainWindow)App.Current.Windows[0];

            if (dialog.ShowDialog(mainWin).GetValueOrDefault())
            {
                path = dialog.SelectedPath;
            }
            else
            {
                return;
            }
            var fullPath = System.IO.Path.Combine(path, $"Отчет за {dataView.Table.TableName} {txtGroup.Text}.xlsx");

            File.Copy(filePath, fullPath, true);

            Excel.Application excel = new Excel.Application();
            Excel.Workbooks workbooks = excel.Workbooks;
            Excel.Workbook wb = workbooks.Open(fullPath);

            //удаляем все листы кроме нужной
            for (int i = wb.Sheets.Count; i > 0; i--)
            {
                if (wb.Sheets[i].Name != dataView.Table.TableName)
                {
                    excel.DisplayAlerts = false;
                    wb.Sheets[i].Delete();
                    excel.DisplayAlerts = true;
                }
            }

            //сохраняем
            wb.Save();
            wb.Close(false);
            excel.Quit();

            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(excel);

            wb = null;
            workbooks = null;
            excel = null;

            GC.Collect();

            MessageBox.Show("Отчет создан!");
        }

        /// <summary>
        /// Кнопка пересчета пропусков 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCalculate_Click(object sender, RoutedEventArgs e)
        {
            CalculateSkips();

            //создаем годовой отчет после пересчета
            YearReport();
        }

        /// <summary>
        /// открывает форму выбора месяцев для создания отчета по месяцам
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReport_Click(object sender, RoutedEventArgs e)
        {
            new SelectMonthsReport().Show();
        }
        /// <summary>
        /// Получаем ячейку из таблицы
        /// </summary>
        /// <param name="cellInfo">координаты ячейки</param>
        /// <returns></returns>
        public DataGridCell GetDataGridCell(DataGridCellInfo cellInfo)
        {
            var cellContent = cellInfo.Column.GetCellContent(cellInfo.Item);
            if (cellContent != null)
                return (DataGridCell)cellContent.Parent;

            return null;
        }
        /// <summary>
        /// ОБновление цветов ФИО в таблице
        /// </summary>
        public void UpdateGroupColors()
        {
            for (int i = 0; i < students.Count; i++)
            {
                if (students[i].Item2 == 1)
                {
                    var dataGridCellInfo = new DataGridCellInfo(DbGrig.Items[i], DbGrig.Columns[1]);
                    var cell = GetDataGridCell(dataGridCellInfo);
                    if (cell == null)
                    {
                        continue;
                    }
                    cell.Background = new SolidColorBrush(Colors.LightBlue);

                }
                if (students[i].Item2 == 2)
                {
                    var dataGridCellInfo = new DataGridCellInfo(DbGrig.Items[i], DbGrig.Columns[1]);
                    var cell = GetDataGridCell(dataGridCellInfo);
                    if (cell == null)
                    {
                        continue;
                    }
                    cell.Background = new SolidColorBrush(Colors.LightYellow);


                }
            }
        }
        private void btnGroups_Click(object sender, RoutedEventArgs e)
        {
            new GroupsWindow(ref students).Show();
        }
        /// <summary>
        /// Событие для обновление таблицы и ячеек
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DbGrig_LayoutUpdated(object sender, EventArgs e)
        {
            if (DbGrig.ItemsSource == null) return;
            if (DbGrig.Visibility != Visibility.Visible) return;
            if (!isTableShow) return;

            UpdateGroupColors();
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {

            Excel.Application xlApp = new Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            //путь на рабочий стол
            var path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            //отпределение переменной имени временного файла либо как временный файл либо как годовой отчет
            string reportFileName = $"Годовой отчет {System.IO.Path.GetFileNameWithoutExtension(fileName)}.xlsx";
            
            //считаем строки
            var totalRow = students.Count + 2;

            //сичтаем сколько месяцев
            var monthsCount = CBChooseList.Items.Count;

            Excel.Workbooks xlWorkBooks;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Sheets xlWorkSheets;

            //выключаем все возможные предепреждения экселя
            xlApp.DisplayAlerts = false;

            xlWorkBooks = xlApp.Workbooks;
            xlWorkBook = xlWorkBooks.Add(System.Reflection.Missing.Value);

            xlWorkSheets = xlWorkBook.Worksheets;
            xlWorkSheet = xlWorkSheets.get_Item(1);

            //заполнения таблицы экселя студентами и их данными о пропуске из глобального списка о пропусках
            for (int i = 0; i < students.Count; i++)
            {
                //заполняем ячуйку студентом
                xlWorkSheet.Cells[i + 2, 1] = students[i].Item1;

            }

            //заполнияем строку итого
            xlWorkSheet.Cells[totalRow, 1] = "ИТОГО:";

            for (int k = 0; k < monthsCount; k++)
            {
                var currMonth = CBChooseList.Items[k].ToString();

                //заполняем загаловки дял месяцев делаем центрирование и объединяем ячейки
                xlWorkSheet.Cells[1, 2 + (k * 3)] = CBChooseList.Items[k].ToString();
                xlWorkSheet.Cells[1, 2 + (k * 3)].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Range[xlWorkSheet.Cells[1, 2 + (k * 3)], xlWorkSheet.Cells[1, 2 + (k * 3) + 2]].Merge();
            }

            xlWorkSheet.Cells[1, 1] = txtGroup.Text;

            //сохраняем выходим и обнулчем все что создали
            xlWorkBook.SaveAs(System.IO.Path.Combine(path, reportFileName));
            xlWorkBook.Close(false);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkSheets);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkBooks);
            Marshal.ReleaseComObject(xlApp);

            xlApp = null;
            xlWorkBooks = null;
            xlWorkBook = null;
            xlWorkSheet = null;
            xlWorkSheets = null;

            GC.Collect();
        }
    }
}
