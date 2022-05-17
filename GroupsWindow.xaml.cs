using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace ExcelData
{
    /// <summary>
    /// Логика взаимодействия для GroupsWindow.xaml
    /// </summary>
    public partial class GroupsWindow : Window
    {
        List<Tuple<string, int>> students;
        int lastGroupSelected = 0;
        public GroupsWindow(ref List<Tuple<string, int>> st)
        {
            InitializeComponent();

            //созраняем ссылку для изменения групп 
            students = st;

            //заполняем список
            list.DataContext = students;
            list1.DataContext = students;
            list2.DataContext = students;

            UpdateList();


        }
        /// <summary>
        /// Обновляет все списки с группами и студентами на основе их значения группы
        /// </summary>
        void UpdateList()
        {
            list.Items.Clear();
            list1.Items.Clear();
            list2.Items.Clear();

            for (int i = 0; i < students.Count; i++)
            {
                if (students[i].Item2 == 0)
                {
                    list.Items.Add(new Tuple<string, int>(students[i].Item1, students[i].Item2));
                }
                if (students[i].Item2 == 1)
                {
                    list1.Items.Add(new Tuple<string, int>(students[i].Item1, students[i].Item2));
                }
                if (students[i].Item2 == 2)
                {
                    list2.Items.Add(new Tuple<string, int>(students[i].Item1, students[i].Item2));
                }
            }
        }
        /// <summary>
        /// Назначение группы 1
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTo1Group_Click(object sender, RoutedEventArgs e)
        {
            if (list.SelectedItems.Count <= 0)
            {
                return;
            }
            //получаем выбранного студента и меняем его группу и обновляем списки
            var s = (Tuple<string, int>)list.SelectedItem;
            var tmp = s.Item1;
            var index = students.FindIndex(a => a.Item1 == tmp);
            students[index] = new Tuple<string, int>(tmp, 1);

            UpdateList();

        }

        /// <summary>
        /// Назначение группы 2
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTo2Group_Click(object sender, RoutedEventArgs e)
        {
            if (list.SelectedItems.Count <= 0)
            {
                return;
            }
            //получаем выбранного студента и меняем его группу и обновляем списки
            var s = (Tuple<string, int>)list.SelectedItem;
            var tmp = s.Item1;
            var index = students.FindIndex(a => a.Item1 == tmp);
            students[index] = new Tuple<string, int>(tmp, 2);

            UpdateList();
        }
        /// <summary>
        /// Удаление из групп
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTo0Group_Click(object sender, RoutedEventArgs e)
        {
            if (list1.SelectedItems.Count <= 0 && list2.SelectedItems.Count <= 0)
            {
                return;
            }
            Tuple<string, int> s = null;
            if (lastGroupSelected == 1)
            {
                s = (Tuple<string, int>)list1.SelectedItem;
            }
            else
            {
                s = (Tuple<string, int>)list2.SelectedItem;
            }
            //получаем выбранного студента и убираем его группу и обновляем списки
            var tmp = s.Item1;
            var index = students.FindIndex(a => a.Item1 == tmp);
            students[index] = new Tuple<string, int>(tmp, 0);

            UpdateList();
        }

        private void list1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            lastGroupSelected = 1;
        
        }

        private void list2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            lastGroupSelected = 2;
        }

        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
           
            var mainPage = (MainPage)((MainWindow)App.Current.Windows[0]).frame.Content;
            mainPage.list.Items.Refresh();
            this.Close();

        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var mainPage = (MainPage)((MainWindow)App.Current.Windows[0]).frame.Content;
            mainPage.list.Items.Refresh();
        }
    }
}
