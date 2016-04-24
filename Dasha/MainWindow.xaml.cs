using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Collections.Generic;
using System.Windows.Media;

namespace Dasha
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string PASSWORD = "5565";//Пароль сохранен на уровне кода
        string connect = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source = C:\Sklad\Dasha\Затраты.accdb;Mode=Share Deny None;Persist Security Info=True;";
        OleDbConnection ConnectDB; 

        List<string> names = new List<string>();//список наименований счетчиков и материалов
        List<System.DateTime> Dates = new List<System.DateTime>();
        Dictionary<string, List<DateTime>> dateTimes = new Dictionary<string, List<DateTime>>();
        List<Expense> exs = new List<Expense>();
        BackgroundWorker backgroundworker;
        BackgroundWorker backgroundworker1;
        BackgroundWorker backgroundworker2;
        DataTable activeTable = new DataTable();
        DataTable asyncTable = new DataTable();
        Border activeBorder;
        TextBlock activeTextBlock = new TextBlock();

        DateTime MinDate = new DateTime();
        DateTime MaxDate = new DateTime();

        /// <summary>
        /// 
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            this.backgroundworker = (BackgroundWorker)this.FindResource("backgroundWorker");
            this.backgroundworker1 = (BackgroundWorker)this.FindResource("backgroundWorker_1");
            this.backgroundworker1.WorkerSupportsCancellation = true;
            this.backgroundworker2 = (BackgroundWorker)this.FindResource("backgroundWorker_2");

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Start_Button_Click(object sender, RoutedEventArgs e)
        {
            if (this.Password_Validation(this.Password_Text.Password))
            {
                this.Sign.Visibility = System.Windows.Visibility.Hidden;
                this.Process.Visibility = System.Windows.Visibility.Visible;


                try
                {
                    (this.Objects_TreeView.Items[0] as TreeViewItem).Focus();
                }
                catch(System.InvalidOperationException)
                {

                }
            }
            else
            {
                MessageBox.Show("Пароль неправильный");                
            }

            this.Password_Text.Password = "";
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (this.Sign.IsVisible)
            {
                if (e.Key == Key.Escape)
                {
                    this.Close();
                    return;
                }

                if (e.Key == Key.Enter)
                {
                    if (this.Password_Text.IsFocused)
                    {
                        this.Start_Button_Click(this.Start_Button, new MouseButtonEventArgs(InputManager.Current.PrimaryMouseDevice, 1, MouseButton.Left));
                    }
                    return;
                }
            }


            if (this.Process.IsVisible)
            {                
                if (e.Key == Key.Enter)
                {
                    if (this.Object_TextBox.IsFocused && this.ObjectChecking())
                    {
                        this.Value_TextBox.Focus();
                        return;
                    }
                    if (this.Value_TextBox.IsFocused && this.ValueChecking(this.Value_TextBox.Text))
                    {
                        this.Date_DatePicker.Focus();
                        return;
                    }
                    if (this.Date_DatePicker.IsKeyboardFocusWithin)
                    {
                        this.Input.Focus();
                        this.Input_MouseDown(this.Input, null);
                        return;
                    }
                    if (this.Notes_TextBox.IsFocused)
                    {
                        this.Input_MouseDown(this.Input, null);
                        return;
                    }             
                }
                if (e.Key == Key.Right)
                {
                    if (this.Date_DatePicker.IsKeyboardFocusWithin)
                    {
                        this.Notes_TextBox.Focus();
                        
                        this.DateChecking(this.Date_DatePicker.DisplayDate.Date);//!!!!!!!!!!!
                        return;
                    }
                }

                if (e.Key == Key.Up)
                {
                    if (this.Date_DatePicker.IsKeyboardFocusWithin)
                    {
                        DateTime dt = this.Date_DatePicker.DisplayDate.Date;
                        dt = dt.AddDays(1.0);
                        this.Date_DatePicker.SelectedDate = dt;
                        this.Date_DatePicker.DisplayDate = dt;

                        if (this.Dates.Contains(dt))
                        {
                            this.Date_DatePicker.Foreground = new SolidColorBrush(Colors.Red);
                            
                        }
                        else
                        {
                            this.Date_DatePicker.Foreground = new SolidColorBrush(Colors.Black);
                        }
                        return;
                    }
                }

                if (e.Key == Key.Down)
                {
                    if (this.Date_DatePicker.IsKeyboardFocusWithin)
                    {
                        DateTime dt = this.Date_DatePicker.DisplayDate.Date;
                        dt = dt.AddDays(-1.0);
                        this.Date_DatePicker.SelectedDate = dt;
                        this.Date_DatePicker.DisplayDate = dt;

                        if (this.Dates.Contains(dt))
                        {
                            this.Date_DatePicker.Foreground = new SolidColorBrush(Colors.Red);
                        }
                        else
                        {
                            this.Date_DatePicker.Foreground = new SolidColorBrush(Colors.Black);
                        }
                        return;
                    }
                }


                if (e.Key == Key.Escape)
                {
                    if (this.Data.IsKeyboardFocusWithin)
                    {
                        this.Objects_TreeView.Focus();
                        return;
                    }
                    if (this.Find_TextBox.IsFocused)
                    {
                        this.Find_TextBox.Text = "";
                        this.Objects_TreeView.Focus();
                        this.Objects_ListBox.Visibility = System.Windows.Visibility.Hidden;
                        return;
                    }
                    if (this.Objects_ListBox.IsKeyboardFocusWithin)
                    {
                        this.Find_TextBox.Focus();
                        return;
                    }
                    if (this.Object_TextBox.IsFocused)
                    {
                        this.Objects_TreeView.Focus();
                        return;
                    }
                    if (this.Value_TextBox.IsFocused)
                    {
                        this.Objects_TreeView.Focus();
                        return;
                    }
                    if (this.Notes_TextBox.IsFocused)
                    {
                        this.Date_DatePicker.Focus();
                        return;
                    }
                    if (this.Date_DatePicker.IsKeyboardFocusWithin)
                    {
                        this.Value_TextBox.Focus();
                        return;
                    }
                    
                    if (this.Objects_TreeView.IsKeyboardFocusWithin && MessageBox.Show("Выйти?", "", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        this.Process.Visibility = System.Windows.Visibility.Hidden;
                        this.Sign.Visibility = System.Windows.Visibility.Visible;
                        this.Password_Text.Focus();
                        return;
                    }                   
                }

                if (e.Key == Key.Space && this.Objects_TreeView.IsKeyboardFocusWithin)
                {
                    this.TextBox_GotFocus(this.Find_TextBox, new RoutedEventArgs());
                    return;
                } 

                if (e.Key == Key.Down && this.Find_TextBox.IsFocused && this.Objects_ListBox.IsVisible && this.Objects_ListBox.Items.Count > 0)
                {
                    this.Objects_ListBox.Focus();
                    return;
                }
            }
            
        }

        private void ListBoxFilling(List<string> L)
        {
            this.Objects_ListBox.Items.Clear();

            foreach (string s in L)
            {
                this.Objects_ListBox.Items.Add(this.CreateTextBlock(s));
            }
        }
        /// <summary>
        /// Проверка пароля на правильность
        /// </summary>
        /// <param name="pass"></param>
        /// <returns></returns>
        private bool Password_Validation(string pass)
        {
            if (pass.Equals(this.PASSWORD))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Dasha_Loaded(object sender, RoutedEventArgs e)
        {
            this.Sign.Visibility = System.Windows.Visibility.Visible;
            this.Process.Visibility = System.Windows.Visibility.Hidden;
            this.Objects_ListBox.Visibility = System.Windows.Visibility.Hidden;

            this.ConnectDB = new OleDbConnection(this.connect);

            this.activeBorder = Border1;

            this.Date_DatePicker.SelectedDate = System.DateTime.Now;
            this.Password_Text.Focus();
            
            
            this.Data.ItemsSource = this.activeTable.DefaultView;

            this.activeTable.Columns.Add(this.CreateColumn("ID", false));
            this.activeTable.Columns.Add(this.CreateColumn("Наименование", false));
            this.activeTable.Columns.Add(this.CreateColumn("Ед изм", false));
            this.activeTable.Columns.Add(this.CreateColumn("Дата", false));
            this.activeTable.Columns.Add(this.CreateColumn("Показания", false));
            this.activeTable.Columns.Add(this.CreateColumn("Расход",false));
            this.activeTable.Columns.Add(this.CreateColumn("План", false));
            this.activeTable.Columns.Add(this.CreateColumn("Разница", false));
            this.activeTable.Columns.Add(this.CreateColumn("Разница (%)", false));
            this.activeTable.Columns.Add(this.CreateColumn("В деньгах", false));
            this.activeTable.Columns.Add(this.CreateColumn("Описание", false));
            this.activeTable.Columns.Add(this.CreateColumn("Примечание", false));

            this.asyncTable.Columns.Add(this.CreateColumn("ID", false));
            this.asyncTable.Columns.Add(this.CreateColumn("Наименование", false));
            this.asyncTable.Columns.Add(this.CreateColumn("Ед изм", false));
            this.asyncTable.Columns.Add(this.CreateColumn("Дата", false));
            this.asyncTable.Columns.Add(this.CreateColumn("Показания", false));
            this.asyncTable.Columns.Add(this.CreateColumn("Расход", false));
            this.asyncTable.Columns.Add(this.CreateColumn("План", true));
            this.asyncTable.Columns.Add(this.CreateColumn("Разница", false));
            this.asyncTable.Columns.Add(this.CreateColumn("Разница (%)", false));
            this.asyncTable.Columns.Add(this.CreateColumn("В деньгах", false));
            this.asyncTable.Columns.Add(this.CreateColumn("Описание", false));
            this.asyncTable.Columns.Add(this.CreateColumn("Примечание", false));

            this.backgroundworker.RunWorkerAsync(this);
        }

        /// <summary>
        /// чтение базы данных в асинхронном режиме
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            //string query1 = "SELECT * FROM Данные";
            string query2 = "SELECT * FROM Ед_измерения";
            string query3 = "SELECT * FROM Материалы";
            string query4 = "SELECT * FROM Оборудование";
            string query5 = "SELECT * FROM Объекты";
            string query6 = "SELECT * FROM Счетчики";
            string query7 = "SELECT * FROM Тип_счетчика";
            string query8 = "SELECT * FROM Группы";

            //OleDbDataAdapter cmd1 = new OleDbDataAdapter(query1, ConnectDB);
            OleDbDataAdapter cmd2 = new OleDbDataAdapter(query2, ConnectDB);
            OleDbDataAdapter cmd3 = new OleDbDataAdapter(query3, ConnectDB);
            OleDbDataAdapter cmd4 = new OleDbDataAdapter(query4, ConnectDB);
            OleDbDataAdapter cmd5 = new OleDbDataAdapter(query5, ConnectDB);
            OleDbDataAdapter cmd6 = new OleDbDataAdapter(query6, ConnectDB);
            OleDbDataAdapter cmd7 = new OleDbDataAdapter(query7, ConnectDB);
            OleDbDataAdapter cmd8 = new OleDbDataAdapter(query8, ConnectDB);

            ConnectDB.Open();//Открытие базы данных

            DataSet dtSet = new DataSet();

            //cmd1.Fill(dtSet, "Данные");//0
            cmd2.Fill(dtSet, "Ед_измерения");//1
            cmd3.Fill(dtSet, "Материалы");//2
            cmd4.Fill(dtSet, "Оборудование");//3
            cmd5.Fill(dtSet, "Объекты");//4
            cmd6.Fill(dtSet, "Счетчики");//5
            cmd7.Fill(dtSet, "Тип_счетчика");//6
            cmd8.Fill(dtSet, "Группы");//7

            ConnectDB.Close();

            e.Result = dtSet;
        }
        


        void t_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {

                this.ObjectPress(sender as TreeViewItem);
                
            }
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="t"></param>
        void ObjectPress(TreeViewItem t)
        {
            if (this.names.Contains(t.Header.ToString()))
            {
                this.Object_TextBox.Focus();
                this.Object_TextBox.Text = t.Header.ToString();
                this.Value_TextBox.Focus();
            }
            else
            {
                t.IsExpanded = true;
            }
        }

        void t_Selected(object sender, RoutedEventArgs e)
        {

            TreeViewItem t = sender as TreeViewItem;
            


            if (t.IsSelected)
            {
                t.IsSelected = true;
                
                this.asyncTable.Clear();
                this.dateTimes.Clear();
                this.Dates.Clear();

                this.Object_TextBox.Text = this.Value_TextBox.Text = this.Notes_TextBox.Text = "";

                TextBox_LostFocus(this.Object_TextBox, new RoutedEventArgs());
                TextBox_LostFocus(this.Value_TextBox, new RoutedEventArgs());
                TextBox_LostFocus(this.Notes_TextBox, new RoutedEventArgs());

                List<string> l = new List<string>();
                this.Going(t, l);
                
                try
                {
                    this.backgroundworker1.RunWorkerAsync(l);
                }
                catch(System.InvalidOperationException)
                {
                    this.backgroundworker1.CancelAsync();
                }

                
                
            }
        }




        /// <summary>
        /// Пробег по всем дочерним элементам дерева
        /// </summary>
        /// <param name="tvi"></param>
        private void Going(TreeViewItem tvi, List<string> l)
        {
            string s = tvi.Header.ToString();
            if (this.names.Contains(s))
            {
                l.Add(s);
            }
            
            

            for (int i = 0; i < tvi.Items.Count; i++)
            {
                this.Going(tvi.Items[i] as TreeViewItem, l);
            }

        }





        /// <summary>
        /// Чтение данных по конкретному наименованию из базы
        /// </summary>
        /// <param name="s"></param>
        private void Reading(string s)
        {
            string query = string.Format("SELECT * FROM Данные WHERE {0}='{1}'", "Наименование", s);

            OleDbDataAdapter cmd = new OleDbDataAdapter(query, ConnectDB);

            DataSet dt = new DataSet();
            cmd.Fill(dt, "Данные");//0
            

            foreach (DataRow r in dt.Tables[0].Rows)
            {
                

                double str = double.Parse(r["Расход_деньги"].ToString().Replace('.', ','));

                Expense ex = this.FindExpense(r["Наименование"].ToString());

                
                if (!this.dateTimes.ContainsKey(ex.Name))
                {
                    this.Dates = new List<DateTime>();
                    this.dateTimes.Add(ex.Name, this.Dates);
                    this.MinDate = DateTime.MaxValue.Date;
                    this.MaxDate = DateTime.MinValue.Date;
                }
                DateTime date = System.DateTime.Parse(r["Дата"].ToString()).Date;
                this.Dates.Add(date);

                if (date < this.MinDate)
                {
                    MinDate = date;
                }
                if (date > this.MaxDate)
                {
                    MaxDate = date;
                }

                this.AddRow(this.asyncTable, r["ID"].ToString(), ex.Name, ex.Dim, r["Дата"].ToString(), r["Показания"].ToString(), r["РасходФакт"].ToString(), r["Расход_план"].ToString(), r["Разница"].ToString(), "", r["Расход_деньги"].ToString(), ex.Description, r["Примечание"].ToString());
            }
        }


        void t_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            this.ObjectPress(sender as TreeViewItem);      
        }

        private void Masked_TextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            TextBox t = sender as TextBox;
            t.Visibility = System.Windows.Visibility.Hidden;
            ((TextBox)t.Tag).Focus();
        }


        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox t = sender as TextBox;
            if (t.Text.Length == 0)
            {
                ((TextBox) t.Tag).Visibility = System.Windows.Visibility.Visible;
            }
            
        }

        private void Object_TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (this.Object_TextBox.Text.Length > 0)
            {
                this.Value_TextBox.IsHitTestVisible = true;
            }
            else
            {
                this.Value_TextBox.IsHitTestVisible = false;
            }
            
            
        }


        private void Masked_TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            this.Masked_TextBox_PreviewMouseDown(sender, new MouseButtonEventArgs(InputManager.Current.PrimaryMouseDevice, 1, MouseButton.Left));
        }

        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBox t = sender as TextBox;
            this.Masked_TextBox_PreviewMouseDown((TextBox)t.Tag, new MouseButtonEventArgs(InputManager.Current.PrimaryMouseDevice, 1, MouseButton.Left));
            
            if (t.TabIndex == 20)
            {
                //this.Find_TextBox.Text = "";
                this.Find_TextBox_TextChanged(this.Find_TextBox, null);
                this.Objects_ListBox.Visibility = System.Windows.Visibility.Visible;
            }
        }


        private void AddRow(DataTable table, params string[] s)
        {
            double x = 0.0;
            System.DateTime dt;

            System.DateTime.TryParse(s[3], out dt);
            s[3] = System.String.Format("{0:dd.MM.yyyy}", dt);//Дата
            double.TryParse(s[4], out x);
            s[4] = System.String.Format("{0:N}", x);//показания
            double.TryParse(s[5], out x);
            s[5] = System.String.Format("{0:N}", x);//расходФакт
            double.TryParse(s[6], out x);
            s[6] = System.String.Format("{0:N}", x);//расходПлан
            double.TryParse(s[7], out x);
            s[7] = System.String.Format("{0:N}", x);//разница
            double.TryParse(s[9], out x);
            s[9] = System.String.Format("{0:N}", x);//разница
            table.Rows.Add(s[0], s[1], s[2], s[3], s[4], s[5], s[6], s[7], s[8], s[9], s[10], s[11]);
        }


        private bool ObjectChecking()
        {
            if (this.names.Contains(this.Object_TextBox.Text))
            {

                return true;
            }
            else
            {
                MessageBox.Show("Введенный бъект расхода не существует!");
                this.Object_TextBox.Text = "";
                this.Object_TextBox.Focus();
                return false;
            }
        }

        private bool ValueChecking(string val)
        {
            double x = 0;
            if (double.TryParse(val, out x) && x > 0)
            {
                this.Value_TextBox.Text = string.Format("{0:N}", x);
                return true;
            }
            else
            {
                MessageBox.Show("Введенное значение неверно!");
                this.Value_TextBox.Text = "";
                this.Value_TextBox.Focus();
                return false;
            }
        }
        /// <summary>
        /// сделать метод универсальным!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        /// </summary>
        /// <returns></returns>
        private bool DateChecking(DateTime dt)
        {
            if (this.Dates.Contains(dt.Date))
            {
                MessageBox.Show("Показатель за эту дату уже введен!!");
                this.Date_DatePicker.Focus();
                return false ;
            }
            else
            {
                return true;
            }
        }

        private Expense FindExpense(string find)
        {
            Expense ex = null;
            foreach (Expense exp in this.exs)
            {
                if (exp.Name.Equals(find))
                {
                    ex = exp;
                    break;
                }
            }
            return ex;
        }

        private TreeViewItem F(ItemCollection ic, string element)
        {
            for (int i = 0; i < ic.Count; i++)
            {
                if (((TreeViewItem)ic[i]).Header.ToString().Equals(element))
                {
                    return (TreeViewItem)ic[i];
                }
                else
                {
                    if (((TreeViewItem)ic[i]).Items.Count > 0)
                    {
                        TreeViewItem tvi = this.F(((TreeViewItem)ic[i]).Items, element);
                        if (tvi != null)
                            return tvi;
                    }
                }
            }
            return null;
        }


        

        private void insert(string s)
        {
            //ConnectDB.Open();
            OleDbCommand myOleDbCommand = new OleDbCommand(s, ConnectDB);
            myOleDbCommand.ExecuteNonQuery();
            myOleDbCommand.Dispose();
            //ConnectDB.Close();
        }

        private void Find_TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            List<string> list = new List<string>();

            foreach (string s in this.names)
            {
                if (s.ToLower().Contains(this.Find_TextBox.Text.Trim().ToLower()))
                {
                    list.Add(s);
                }
            }

            this.ListBoxFilling(list);
        }
        

        private void Objects_ListBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Up && this.Objects_ListBox.SelectedIndex == 0)
            {
                this.Find_TextBox.Focus();
            }
            if (e.Key == Key.Enter)
            {
                string s = (this.Objects_ListBox.SelectedItem as TextBlock).Text;
                TreeViewItem t = this.F(Objects_TreeView.Items, s);

                this.Find_TextBox.Focus();

                

                this.Find_TextBox.Text = "";
                this.Objects_TreeView.Focus();
                this.Objects_ListBox.Visibility = System.Windows.Visibility.Hidden;

                TreeViewItem t1 = t, t2;
                while ((t2 = t1.Parent as TreeViewItem) != null)
                {
                    t2.IsExpanded = true;
                    t1 = t2;
                }
                

                t.IsSelected = true;
                t.BringIntoView();
                //this.Objects_TreeView
                this.ObjectPress(t);
            }
        }

        

        private void border_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            Border b = sender as Border;
            b.BorderBrush = this.activeBorder.BorderBrush;
            this.activeBorder.BorderBrush = new SolidColorBrush(SystemColors.ScrollBarColor);
            this.Process.Background = b.Background;//поменять цвет фона для окна
            this.activeBorder = b;
        }

        private void BackgroundWorker_DoWork_1(object sender, DoWorkEventArgs e)
        {
            List<string> l = e.Argument as List<string>;

            ConnectDB.Open();
            foreach(string s in l)
            {
                this.Reading(s);
            }

            
        }

        private void BackgroundWorker_RunWorkerCompleted_1(object sender, RunWorkerCompletedEventArgs e)
        {
            ConnectDB.Close();
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message, "Произошла ошибка");
            }
            else
            {
                this.activeTable.Clear();
                //копирование таблицы
                foreach (DataRow r in this.asyncTable.Rows)
                {
                    activeTable.Rows.Add(r["ID"].ToString(), r["Наименование"].ToString(), r["Ед изм"].ToString(), r["Дата"].ToString(), r["Показания"].ToString(), r["Расход"].ToString(), r["План"].ToString(), r["Разница"].ToString(), r["Разница (%)"].ToString(), r["В деньгах"].ToString(), r["Описание"].ToString(), r["Примечание"].ToString());
                }

                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                if (this.Data.Items.Count > 0)
                {
                    this.Data.ScrollIntoView(this.Data.Items[this.Data.Items.Count - 1]);
                }

                if (this.dateTimes.Count == 1 && this.Dates.Contains(this.Date_DatePicker.DisplayDate.Date))
                {
                    this.Date_DatePicker.Foreground = new SolidColorBrush(Colors.Red);
                }
                else
                {
                    this.Date_DatePicker.Foreground = new SolidColorBrush(Colors.Black);
                }
            }
            
        }

        private void T_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            string s = (sender as TextBlock).Text;
            TreeViewItem t = this.F(Objects_TreeView.Items, s);

            this.Find_TextBox.Focus();



            this.Find_TextBox.Text = "";
            this.Objects_TreeView.Focus();
            this.Objects_ListBox.Visibility = System.Windows.Visibility.Hidden;

            TreeViewItem t1 = t, t2;
            while ((t2 = t1.Parent as TreeViewItem) != null)
            {
                t2.IsExpanded = true;
                t1 = t2;
            }


            t.IsSelected = true;
            t.BringIntoView();
            //this.Objects_TreeView
            this.ObjectPress(t);
        }

        private void Data_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                if (MessageBox.Show("Удалить строку?", "", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    this.Data.CanUserDeleteRows = true;
                    List<string> id = new List<string>();

                    foreach (DataRowView row in this.Data.SelectedItems)
                    {
                        id.Add(row.Row.ItemArray[0].ToString());//достать id
                    }
                    this.backgroundworker2.RunWorkerAsync(id);

                    
                }
                else
                {
                    this.Data.CanUserDeleteRows = false;
                }
            }
        }

        private void Data_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            string s = e.Column.Header.ToString();

            string cell = (e.EditingElement as TextBox).Text;//новые, введенные данные
            string id = (e.Row.Item as DataRowView).Row.ItemArray[0].ToString();
            double x = 0.0;
            System.DateTime dt = System.DateTime.Now;



            if (s.Equals("Дата"))
            {         
                //старая дата
                string cel = (e.Row.Item as DataRowView).Row.ItemArray[3].ToString();

                if (DateTime.TryParse(cell, out dt))
                {
                    
                    if (!cell.Equals(cel))//если были произведены изменения в ячейке
                    {
                        //найти соответствующий счетчик или материал в коллекции
                        Expense ex = this.FindExpense((e.Row.Item as DataRowView).Row.ItemArray[1].ToString());

                        List<DateTime> ldt;
                        this.dateTimes.TryGetValue(ex.Name, out ldt);

                        if (ldt.Contains(dt.Date))
                        {
                            MessageBox.Show("Показатель за эту дату уже введен!!");
                            (e.EditingElement as TextBox).Text = cel;
                        }
                        else
                        {
                            foreach (DataRow dtRow in this.activeTable.Rows)
                            {
                                if (dtRow.ItemArray[0].Equals(id))
                                {
                                    dtRow["Дата"] = string.Format("{0:dd.MM.yyyy HH:mm:ss}", dt);

                                    string sql = string.Format("UPDATE Данные SET {0} = '{1}' WHERE {2} = {3}", "Дата", dt.ToString(), "ID", id);

                                    this.ConnectDB.Open();
                                    this.insert(sql);
                                    this.ConnectDB.Close();

                                    ldt.Add(dt.Date);
                                    ldt.Remove(DateTime.Parse(cel).Date);

                                    break;
                                }
                            }
                        }
                        
                    }            
                }
                else
                {
                    (e.EditingElement as TextBox).Text = cel;
                    MessageBox.Show("Введено некорректное значение");
                }
            }



            if (s.Equals("Показания"))
            {
                string cel = (e.Row.Item as DataRowView).Row.ItemArray[4].ToString();

                if (double.TryParse(cell, out x))
                {
                    if (!cell.Equals(cel))
                    {
                        Expense ex = this.FindExpense((e.Row.Item as DataRowView).Row.ItemArray[1].ToString());
                        x = x * ex.k;//??????????????????

                        foreach (DataRow dtRow in this.activeTable.Rows)
                        {
                            if (dtRow.ItemArray[0].Equals(id))
                            {
                                dtRow["Показания"] = string.Format("{0:N}", x);
                                dtRow["Расход"] = string.Format("{0:N}", x);
                                dtRow["В деньгах"] = string.Format("{0:N}", x * ex.Price);

                                string sql1 = string.Format("UPDATE Данные SET {0} = {1} WHERE {2} = {3}", "Показания", x.ToString(), "ID", id);
                                string sql2 = string.Format("UPDATE Данные SET {0} = {1} WHERE {2} = {3}", "РасходФакт", x.ToString(), "ID", id);
                                string sql3 = string.Format("UPDATE Данные SET {0} = {1} WHERE {2} = {3}", "Расход_деньги", x.ToString(), "ID", id, "РасходФакт", dtRow.ItemArray[9]);

                                this.ConnectDB.Open();
                                this.insert(sql1);
                                this.insert(sql2);
                                this.insert(sql3);
                                this.ConnectDB.Close();

                                break;
                            }
                        }
                    }                    
                }
                else
                {
                    (e.EditingElement as TextBox).Text = cel;
                    MessageBox.Show("Введено некорректное значение");
                }
            }



            if (s.Equals("Примечание"))
            {
                string sql = string.Format("UPDATE Данные SET {0} = '{1}' WHERE {2} = {3}", "Примечание", cell, "ID", id);

                this.ConnectDB.Open();
                this.insert(sql);
                this.ConnectDB.Close();
            }
        }
        /// <summary>
        /// во время поиска поменять цвет шрифта в TextBlock при выделении
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Objects_ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            this.activeTextBlock.Foreground = new SolidColorBrush(Colors.Black);
            this.activeTextBlock = this.Objects_ListBox.SelectedItem as TextBlock;
            if (this.activeTextBlock == null)
            {
                this.activeTextBlock = new TextBlock();
            }
            else
            {
                this.activeTextBlock.Foreground = new SolidColorBrush(Colors.White);
                string s = this.activeTextBlock.Text;
                TreeViewItem tvi = this.F(this.Objects_TreeView.Items, s);

                TreeViewItem t1 = tvi, t2;
                while ((t2 = t1.Parent as TreeViewItem) != null)
                {
                    t2.IsExpanded = true;
                    t1 = t2;
                }


                tvi.IsSelected = true;
                tvi.BringIntoView();
            }
            

            
        }

        /// <summary>
        /// редактирование при одиночном клике
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Data_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            if (e.AddedCells.Count == 0) return;
            var currentCell = e.AddedCells[0];
            if (currentCell.Column == this.Data.Columns[3] || currentCell.Column == this.Data.Columns[4] || currentCell.Column == this.Data.Columns[10])
            {
                this.Data.BeginEdit();
            }

        }

        private void Input_MouseDown(object sender, MouseButtonEventArgs e)
        {
            DateTime date_now = this.Date_DatePicker.DisplayDate.Date;

            if (this.ObjectChecking() && this.ValueChecking(this.Value_TextBox.Text) && this.DateChecking(date_now))
            {
                Expense ex = this.FindExpense(this.Object_TextBox.Text);
                double now = Double.Parse(this.Value_TextBox.Text);

                DateTime date_min = DateTime.MinValue, date_max = DateTime.MaxValue;
                foreach (DateTime dt in this.Dates)
                {
                    if (dt < date_now && date_now - dt < date_now - date_min)
                    {
                        date_min = dt;
                    }
                    if (dt > date_now && dt - date_now < date_max - date_now)
                    {
                        date_max = dt;
                    }
                }

                if (this.Dates.Count == 0)//Добавление данных в пустой список
                {
                    this.Add_Record(ex, now, 0, this.Date_DatePicker.DisplayDate.Date, this.Notes_TextBox.Text);
                }
                else
                {
                    //Добавление элемента с наименьшей датой
                    if (date_min.Equals(DateTime.MinValue))
                    {
                        int n = (date_max.Date - date_now.Date).Days;

                        double max = 0.0;
                        string note = "", id = "", date = "";


                        foreach (DataRow dr in this.activeTable.Rows)
                        {
                            if (date_max.Equals(DateTime.Parse(dr["Дата"].ToString()).Date))
                            {
                                max = Double.Parse(dr["Показания"].ToString());
                                note = dr["Примечание"].ToString();
                                id = dr["ID"].ToString();
                                date = dr["Дата"].ToString();
                                break;
                            }
                        }

                        double del = (max - now) / n;

                        if (max < now)
                        {
                            MessageBox.Show("Значение слишком велико!");
                            this.Value_TextBox.Text = "";
                            this.Value_TextBox.Focus();
                            return;
                        }

                        if (n > 1 && MessageBox.Show("Добавить " + n + " записи?", "", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                        {
                            this.Add_Record(ex, now, 0, this.Date_DatePicker.DisplayDate.Date, this.Notes_TextBox.Text);
                            for (int i = 1; i < n; i++)
                            {
                                now = Math.Round(now + del, 2);
                                this.Add_Record(ex, now, Math.Round(del * ex.k, 2), date_now.AddDays(i), "");
                            }
                            this.Add_Record(ex, max, Math.Round(del * ex.k, 2), this.Date_DatePicker.DisplayDate.Date, note);

                            this.ConnectDB.Open();
                            string sql = string.Format("DELETE FROM Данные WHERE ({0} = {1});", "ID", id);
                            this.insert(sql);
                            this.ConnectDB.Close();
                        }
                        if (n == 1)
                        {
                            this.Add_Record(ex, now, 0, this.Date_DatePicker.DisplayDate.Date, note);
                            this.Add_Record(ex, max, Math.Round(del * ex.k, 2), DateTime.Parse(date), note);
                            
                            this.ConnectDB.Open();
                            string sql = string.Format("DELETE FROM Данные WHERE ({0} = {1});", "ID", id);
                            this.insert(sql);
                            this.ConnectDB.Close();
                        }
                    }



                    //Добавление элемента с наибольшей датой
                    if (date_max.Equals(DateTime.MaxValue))
                    {
                        int n = (date_now.Date - date_min.Date).Days;

                        double min = 0.0;

                        foreach (DataRow dr in this.activeTable.Rows)
                        {
                            if (date_min.Equals(DateTime.Parse(dr["Дата"].ToString()).Date))
                            {
                                min = Double.Parse(dr["Показания"].ToString());
                                break;
                            }
                        }

                        double del = (now - min) / n;

                        if (now < min)
                        {
                            MessageBox.Show("Значение слишком мало!");
                            this.Value_TextBox.Text = "";
                            this.Value_TextBox.Focus();
                            return;
                        }


                        if (n > 1 && MessageBox.Show("Добавить " + n + " записи?", "", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                        {

                            for (int i = 1; i < n; i++)
                            {
                                min = Math.Round(min + del, 2);
                                this.Add_Record(ex, min, Math.Round(del * ex.k, 2), date_min.AddDays(i), "");
                            }
                            this.Add_Record(ex, now, Math.Round(del * ex.k, 2), this.Date_DatePicker.DisplayDate.Date, this.Notes_TextBox.Text);
                        }
                        if (n == 1)
                        {
                            this.Add_Record(ex, now, Math.Round(del * ex.k, 2), this.Date_DatePicker.DisplayDate.Date, this.Notes_TextBox.Text);
                        }
                    }



                    if (this.Dates.Contains(date_min) && this.Dates.Contains(date_max))
                    {
                        double min = 0.0, max = 0.0;

                        foreach (DataRow dr in this.activeTable.Rows)
                        {
                            if (date_min.Equals(DateTime.Parse(dr["Дата"].ToString())))
                            {
                                min = Double.Parse(dr["Показания"].ToString());
                            }
                            if (date_max.Equals(DateTime.Parse(dr["Дата"].ToString())))
                            {
                                max = Double.Parse(dr["Показания"].ToString());
                            }
                        }

                        if (now < min || now > max)
                        {
                            MessageBox.Show("Значение не соответствует ограничениям!");
                            this.Value_TextBox.Text = "";
                            this.Value_TextBox.Focus();
                            return;
                        }

                        Add_Record(ex, now, now, date_now, this.Notes_TextBox.Text);
                    }
                }

                t_Selected(this.Objects_TreeView.SelectedItem, null);

                this.Data.UpdateLayout();

                this.Objects_TreeView.Focus();
            }
        }
        /// <summary>
        /// добавление записи в таблицу
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="val"></param>
        /// <param name="?"></param>
        private void Add_Record(Expense ex, double val, double fact, DateTime dt, string note)
        {
            string col1 = "Наименование";
            string col2 = "Дата";
            string col3 = "Показания";
            string col4 = "РасходФакт";
            string col5 = "Расход_план";
            string col6 = "Разница";
            string col7 = "Расход_деньги";
            string col8 = "Примечание";
            string sql = string.Format("INSERT INTO Данные ({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}) VALUES ('{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}')", col1, col2, col3, col4, col5, col6, col7, col8, ex.Name, dt.ToString(), val.ToString(), fact.ToString(), "", "", (ex.Price * fact).ToString(), note);

            this.ConnectDB.Open();
            this.insert(sql);
            this.ConnectDB.Close();
        }

        private void Input_MouseEnter(object sender, MouseEventArgs e)
        {
            this.Input.Background = this.Process.Background;
        }

        private void Input_MouseLeave(object sender, MouseEventArgs e)
        {
            this.Input.Background = null; ;
        }
    }
}
