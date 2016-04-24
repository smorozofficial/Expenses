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
    public partial class MainWindow : Window
    {
        /// <summary>
        /// асинхронное удаление записей из базы данных
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_DoWork_2(object sender, DoWorkEventArgs e)
        {
            List<string> ids = e.Argument as List<string>;

            this.ConnectDB.Open();
            foreach (string id in ids)
            {
                string sql = string.Format("DELETE FROM Данные WHERE ({0} = {1});", "ID", id);
                this.insert(sql);
            }

        }

        private void BackgroundWorker_RunWorkerCompleted_2(object sender, RunWorkerCompletedEventArgs e)
        {
            this.ConnectDB.Close();
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message, "Произошла ошибка");
            }
            else
            {
                t_Selected(this.Objects_TreeView.SelectedItem, null);
                this.Objects_TreeView.Focus();
            }
        }
    }
}