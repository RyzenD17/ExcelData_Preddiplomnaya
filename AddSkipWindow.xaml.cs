using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelData
{
    /// <summary>
    /// Логика взаимодействия для AddSkipWindow.xaml
    /// </summary>
    public partial class AddSkipWindow : Window
    {
        //храним данные о текущем студенте
        StudentsData student;
        //данные таблицы
        DataView dataView;
        //номер строки и стобца студента в эксель документе
        int rowStudent = 0, colStudent = 0;
        //номер строки и стобца текущего дня который мы меняем в эксель документе
        int rowCurrentDay = 0, colCurrentDay = 0;
        //список пропущенных дней у студента - строка и столбец в документе
        List<(int row, int col)> missedDaysCoords;
        //наличие пропусков для проверки
        bool hasSkips = false;
        //ссылка на главную страницу
        MainPage mainPage;

        /// <summary>
        /// КОнструктор страницы
        /// </summary>
        /// <param name="data">Передаем студента для его заполнения</param>
        /// <param name="table">Передаем всю таблицу для взятия данных</param>
        public AddSkipWindow(StudentsData data, DataView table)
        {
            InitializeComponent();

            //деляем первоначальные заполнения
            missedDaysCoords = new List<(int row, int col)>();
            datePicker.IsTodayHighlighted = true;
            datePicker.IsEnabled = false;
            student = data;
            dataView = table;
            stackWarning.Visibility = Visibility.Collapsed;

            //ищем координаты текущего студента в таблице
            for (int i = 0; i < dataView.Table.Rows.Count; i++)
            {
                var tmp = dataView.Table.Rows[i][1].ToString();
                if (tmp == student.FIO)
                {
                    colStudent = 1;
                    rowStudent = i;
                }
            }
            
            //делаем обновление дат и прочего
            UpdateDates();

            mainPage = (MainPage)((MainWindow)App.Current.Windows[0]).frame.Content;
        }

        private void btnEditDate_Click(object sender, RoutedEventArgs e)
        {
            datePicker.IsEnabled = true;
            btnEditDate.Visibility = Visibility.Hidden;
        }
        /// <summary>
        /// Меняем значение переменных с координатами на координаты выбранного дня
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void datePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (datePicker.SelectedDate != null)
            {
                colCurrentDay = missedDaysCoords.Where(a => a.col - 1 == datePicker.SelectedDate.Value.Day).FirstOrDefault().col;
            }
            rowCurrentDay = rowStudent;
        }
        /// <summary>
        /// ПРи закрытии заного сичтываем файл обновляем
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            mainPage.readFile(Properties.Settings.Default.FilePath, true);
        }
        public void UpdateDates()
        {
            //делаем очистку дат и переменных
            missedDaysCoords.Clear();
            datePicker.BlackoutDates.Clear();
            hasSkips = false;
            //ищем текущий месяц по таблице
            DateTime dateOfTable = DateTime.ParseExact($"01.{dataView.Table.TableName}.{DateTime.Today.Year}", "dd.MMMM.yyyy", CultureInfo.CurrentCulture);
            bool tableIsCurrentMonth = dateOfTable.Month == DateTime.Today.Month;

            //ищем дни пропусков по таблице
            for (int i = 0; i < dataView.Table.Columns.Count; i++)
            {
                if (dataView.Table.Rows[rowStudent][i].ToString() == "")
                {
                    hasSkips = true;
                    missedDaysCoords.Add((rowStudent, i));
                }
            }
            if (hasSkips)
            {
                stackWarning.Visibility = Visibility.Visible;
                var month = dataView.Table.TableName;
                var datePart = "." + month + "." + DateTime.Today.Year.ToString();
                string currDayTable = "";
                DateTime date = default;
                student.data = new List<(DateTime, int?)>();
                var excludeDates = new List<DateTime>();
                var excludeIndexes = new List<int>();
                //переводим все пропуски из таблицы в реальные даты
                for (int i = 0; i < missedDaysCoords.Count; i++)
                {
                    currDayTable = (missedDaysCoords[i].col - 1).ToString("00");
                    date = DateTime.ParseExact(currDayTable + datePart, "dd.MMMM.yyyy", CultureInfo.CurrentCulture);

                    //исключаем дни которые позже сегоднешней
                    if (date > DateTime.Today)
                    {
                        excludeDates.Add(date);
                        excludeIndexes.Add(i);
                    }
                        student.data.Add((date, null));
                }
                excludeIndexes = excludeIndexes.OrderByDescending(a => a).ToList();
                student.data.RemoveAll(a => excludeDates.Contains(a.Item1));
                //исключаем дни которые позже сегоднешней
                foreach (var item in excludeIndexes)
                {
                    missedDaysCoords.RemoveAt(item);
                }
                //если дней нет то меняем флаг
                if (missedDaysCoords.Count == 0 || student.data.Count == 0)
                {
                    hasSkips = false;
                }
            }
            if (!hasSkips)
            {
                stackWarning.Visibility = Visibility.Hidden;
            }

            //корректируем даты для календаря которые нужно убрать из выделения
            var dates = hasSkips ? student.data.Select(a => a.Item1).ToList() : new List<DateTime>() { dateOfTable.AddDays(-1) };
            if (!tableIsCurrentMonth && hasSkips)
            {
                dates.Insert(0, dateOfTable.AddDays(-1));
            }
            if (!hasSkips)
            {
                dates.Add(dateOfTable.AddMonths(1));
            }
            if (hasSkips && student.data.Count == 1)
            {
                dates.Add(dateOfTable.AddMonths(1));
            }

            datePicker.SelectedDate = null;

            var firstDate = dates.First();
            var lastDate = dates.Last();
            var dateCounter = firstDate;

            if (firstDate.Day != 1)
            {
                var start = new DateTime(firstDate.Year, firstDate.Month, 1);
                var end = firstDate.AddDays(-1);
                datePicker.BlackoutDates.Add(new CalendarDateRange(start, end));
            }
            if (lastDate.Day != DateTime.Today.Day)
            {
                var day = DateTime.Today;
                datePicker.BlackoutDates.Add(new CalendarDateRange(day));
            }


            //заполняем зачеркнутые дни чтобы их нельзя было выбрать
            //заполняем из dates
            foreach (var d in dates.Skip(1))
            {
                if (d.AddDays(-1).Date != dateCounter.Date)
                {
                    var from = dateCounter.AddDays(1);
                    var to = d.AddDays(-1);
                    if (from != to)
                    {
                        datePicker.BlackoutDates.Add(new CalendarDateRange(from, to));

                    }
                    else
                    {
                        datePicker.BlackoutDates.Add(new CalendarDateRange(from));
                    }

                }

                dateCounter = d;
            }
            

            //если делаем текущий месяц то ставим выделение на сегодня и ставим значения переменныз в текущие координаты дня по таблице
            if (!datePicker.BlackoutDates.Any(a => a.Start <= DateTime.Today && a.End >= DateTime.Today) && tableIsCurrentMonth)
            {
                datePicker.SelectedDate = DateTime.Today;
                datePicker.DisplayDate = DateTime.Today;

                var missD = missedDaysCoords.Where(a => a.col - 1 == datePicker.SelectedDate.Value.Day).FirstOrDefault();
                colCurrentDay = missD.col;
                rowCurrentDay = rowStudent;
            }
            if (!hasSkips)
            {
                datePicker.IsEnabled = false;
            }
            //ставим границв календаря - месяц который редактируем
            datePicker.DisplayDateStart = dateOfTable;
            datePicker.DisplayDateEnd = dateOfTable.AddMonths(1).AddDays(-1);

        }
        /// <summary>
        /// Обновление состояния окна дат и прочего после внесения пропуска в таблицу
        /// </summary>
        /// <param name="lastValue"></param>
        public void UpdateState(string lastValue = "")
        {
            //добавляем значение часов в таблицу если есть
            if (lastValue != "")
            {
                dataView.Table.Rows[rowStudent][colCurrentDay] = lastValue;
            }
            //убираем этот день из списков пропусков
            var studentDataRemove = student.data.Where(a => a.Item1 == datePicker.SelectedDate.Value).FirstOrDefault();
            student.data.Remove(studentDataRemove);
            var missedRemove = missedDaysCoords.Where(a => a.col == colCurrentDay).FirstOrDefault();
            missedDaysCoords.Remove(missedRemove);


            //облнвляем даты и каледарь
            UpdateDates();

            txtHours.Text = "";
            checkOkSkip.IsChecked = false;

        }
        /// <summary>
        /// Добавляем пропуск студенту на выбранный день
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //день не выбран значит уже проставлен пропуск
            if (datePicker.SelectedDate == null)
            {
                MessageBox.Show("На сегодня уже заполнено посещение");
                return;
            }
            try
            {
                //создаем подключение к экселю
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = excel.Workbooks.Open(Properties.Settings.Default.FilePath);
                Excel.Worksheet sheet1 = (Excel.Worksheet)workbook.Sheets[dataView.Table.TableName];


                //создаем строку с результатом пропуска
                var hours = int.Parse(txtHours.Text);
                var check = (bool)checkOkSkip.IsChecked;
                var result = $"{(check && hours != 0 ? "-" : "")}{hours}";

                var r = sheet1.Cells[rowCurrentDay + 3, colCurrentDay + 1];
                Excel.Range myRange = (Excel.Range)r;
                //заполняем ячейку
                if (result != null || result != "")
                {
                    myRange.Value2 = result;
                }
                else
                {
                    myRange.Value2 = "";
                }
                //сохраняем и выходим из всего
                workbook.Save();
                workbook.Close(true);
                excel.Quit();
                excel = null;
                workbook = null;
                sheet1 = null;

                //обновляем состояние окна
                UpdateState(result);

            }
            catch (Exception ex)
            {

                MessageBox.Show("Ошибка! " + ex.Message);
            }
        }

        
    }
}
