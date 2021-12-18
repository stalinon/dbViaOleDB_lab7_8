using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace dataBaseWPF
{
    /// <summary>
    ///     Главное окно приложения
    /// </summary>
    public partial class MainWindow : Window
    {
        private DataTable dataTable;

        private static OleDbConnection connectionDB = new OleDbConnection(
            @"Provider=Microsoft.ACE.OLEDB.12.0;" +
            @"Data Source=""D:\Projects\dataBaseWPF\dataBaseWPF\files\base.accdb"""
            );

        public MainWindow()
        {
            InitializeComponent();
            TableNameListBox.ItemsSource = tabs.Keys;
            ListBox.ItemsSource = tabs2.Keys;
            ListBox2.ItemsSource = tabs3.Keys;
        }

        #region Словари для заполнения списков таблиц и запросов

        /// <summary>
        ///     Названия таблиц
        /// </summary>
        private static readonly Dictionary<string, int> tabs = new Dictionary<string, int>
        {
            {"Студент",1 }, {"Факультет",2 }, {"Группа",3 }, {"Текущая успеваемость", 4 },{"Семестр", 5 },{"Предмет",6 }
        };

        /// <summary>
        ///  Названия запросов без параметров
        /// </summary>
        private static readonly Dictionary<string, int> tabs2 = new Dictionary<string, int>
        {
            {"Студенты, родившиеся после 2000", 1 }, {"Предметы по убыванию кол-ва часов", 2}, {"ИТОГИ Количество студентов", 3},
            {"6 ВЛОЖЕННЫЙ Самые младшие студенты", 4}, {"6 ВЛОЖЕННЫЙ Самые старшие студенты", 5},{"6 ВЛОЖЕННЫЙ С ALL", 6},
            {"SQL c BETWEEN", 7}, {"SQL Итоговый", 8}, {"SQL Многотабличный", 9}, {"SQL с сортировкой по возрастанию", 10},
            {"6 UNION Номера студенческих билетов старост и профоргов", 11}
        };

        /// <summary>
        ///  Названия запросов с параметрами
        /// </summary>
        private static readonly Dictionary<string, int> tabs3 = new Dictionary<string, int>
        {
            {"Успеваемость по оценке", 1}, {"Студенты с фамилией", 2}, {"Факультет Запрос", 3}
        };

        #endregion

        #region Обработчики событий с удалением, обновлением и добавлением

        /// <summary>
        ///     ДОБАВЛЕНИЕ
        /// </summary>
        private void DataTable_TableNewRow(object sender, DataTableNewRowEventArgs e)
        {
            dataTable.ColumnChanged -= DataTable_ColumnChanged;
            dataTable.RowChanged += DataTable_RowChanged;
        }

        /// <summary>
        ///     ДОБАВЛЕНИЕ
        /// </summary>
        private void DataTable_RowChanged(object sender, DataRowChangeEventArgs e)
        {
            var row = e.Row;
            if (row != null)
            {
                //  создание инструкции
                var sql = $"insert into [{TableNameListBox.SelectedItem}] (";
                //  добавление названий колонок через запятую
                foreach (DataColumn column in dataTable.Columns)
                {
                    sql += $"[{column.ColumnName}], ";
                }
                //  удаление последней запятой, начало добавления значений
                sql = sql.Substring(0, sql.LastIndexOf(','));
                sql += ") values (";
                //  добавление значений в колонки
                foreach (object cell in row.ItemArray)
                {
                    sql += ParseSQLValue(cell.ToString()) + ", ";
                }
                //  окончание создания инструкции, выполнение
                sql = sql.Substring(0, sql.LastIndexOf(','));
                sql += ")";
                ExecuteCmd(sql);

                try
                {
                    //  обновление базы данных
                    var dataAdapter = new OleDbDataAdapter(sql, connectionDB);
                    var data = new DataSet();
                    dataAdapter.Fill(data);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "WARNING", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            dataTable.RowChanged -= DataTable_RowChanged;
            dataTable.ColumnChanged += DataTable_ColumnChanged;
        }

        /// <summary>
        ///     УДАЛЕНИЕ
        /// </summary>
        private void DataTable_RowDeleting(object sender, DataRowChangeEventArgs e)
        {
            var sql = $"DELETE * FROM [{TableNameListBox.SelectedItem}] WHERE [{dataTable.Columns[0].ColumnName}] = {(Table.CurrentItem as DataRowView).Row.ItemArray[0]};";
            ExecuteCmd(sql);
        }

        /// <summary>
        ///     ОБНОВЛЕНИЕ
        /// </summary>
        private void DataTable_ColumnChanged(object sender, DataColumnChangeEventArgs e)
        {
            var sql = $"Update [{TableNameListBox.SelectedItem}] set [{Table.CurrentColumn.Header}] = {ParseSQLValue(e.ProposedValue.ToString())} " +
                        $"where [{dataTable.Columns[0].ColumnName}] = {(Table.CurrentItem as DataRowView).Row.ItemArray[0]}";
            ExecuteCmd(sql);
        }

        #endregion

        #region Обработчики событий элементов GUI, вызывающих запросы на вывод таблицы

        /// <summary>
        ///     ВЫЗОВ ЗАПРОСА БЕЗ ПАРАМЕТРОВ
        /// </summary>
        private void TableNameListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var sql = $"SELECT * FROM [{(sender as ListBox).SelectedItem}]";
            UpdateTab(sql);
        }

        /// <summary>
        ///     ВЫЗОВ ЗАПРОСА С ПАРАМЕТРАМИ
        /// </summary>
        private void ListBox2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var window = new Window1();
            window.ShowDialog();

            if (!window.IsActive && window.value != null)
            {
                var cell = window.value;
                var sql = $"EXEC [{(sender as ListBox).SelectedItem}] ";
                
                sql += ParseSQLValue(cell) + ";";
                UpdateTab(sql);
            }
        }

        #endregion

        #region Обработчики событий взаимодействия с окном

        private void Border_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
                DragMove();
        }

        private void ButtonMinimize_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void ButtonWindowState_Click(object sender, RoutedEventArgs e)
        {
            if (WindowState != WindowState.Maximized)
                WindowState = WindowState.Maximized;
            else
                WindowState = WindowState.Normal;
        }

        private void ButtonCloseWindow_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        #endregion

        #region Методы выполнения инструкций SQL

        /// <summary>
        ///     ПАРС СТРОКИ ЗНАЧЕНИЯ SQL,
        ///     ПРЕОБРАЗОВАНИЕ В ПОДОБАЮЩИЙ ВИД
        /// </summary>
        private static string ParseSQLValue(string cell)
        {
            if (int.TryParse(cell, out int res))
                return $"{res}";
            else if (bool.TryParse(cell, out bool resB))
                return resB ? $"TRUE" : $"FALSE";
            else if (DateTime.TryParse(cell, out DateTime resD))
                return $"#{resD.ToString().Split(' ').First().Replace('.', '/')}#";
            else if (cell is string)
                return $"\'{cell}\'";
            return null;
        }

        /// <summary>
        ///     ВЫПОЛНЕНИЕ КОМАНДЫ SQL
        /// </summary>
        private static void ExecuteCmd(string sql)
        {
            try
            {
                connectionDB.Open();
                var cmd = new OleDbCommand
                {
                    Connection = connectionDB,
                    CommandText = sql
                };
                cmd.ExecuteNonQuery();
                connectionDB.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "WARNING", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally
            {
                connectionDB.Close();
            }
        }

        /// <summary>
        ///     ОБНОВЛЕНИЕ ЭЛЕМЕНТА КОНТРОЛЯ 
        ///     GUI ТАБЛИЦЫ
        /// </summary>
        private void UpdateTab(string sql)
        {
            try
            {
                var dataAdapter = new OleDbDataAdapter(sql, connectionDB);
                var data = new DataSet();
                dataAdapter.Fill(data);
                dataTable = data.Tables[0];
                Table.ItemsSource = dataTable.DefaultView;
                dataTable.ColumnChanged += DataTable_ColumnChanged;
                dataTable.RowDeleting += DataTable_RowDeleting;
                dataTable.TableNewRow += DataTable_TableNewRow;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "WARNING", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        } 

        #endregion
    }
}
