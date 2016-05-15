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
    /// 
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// асинхронное чтение базы данных
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message, "Произошла ошибка");
            }
            else
            {
                DataSet dtSet = (DataSet)e.Result;



                foreach (DataRow dtRow in dtSet.Tables[6].Rows)//групп
                {
                    string groupp = dtRow["Наименование"].ToString();

                    TreeViewItem tvi = this.CreateTreeViewItem(groupp);
                    this.Objects_TreeView.Items.Add(tvi);

                    foreach (DataRow dtRow1 in dtSet.Tables[3].Rows)//по объектам (складам)
                    {
                        string obj = dtRow1["Наименование"].ToString();
                        string gru = dtRow1["Группа"].ToString();

                        if (groupp.Equals(gru))
                        {
                            TreeViewItem tvi1 = this.CreateTreeViewItem(obj);
                            tvi.Items.Add(tvi1);

                            TreeViewItem tvi2 = this.CreateTreeViewItem(dtSet.Tables[1].TableName);
                            tvi1.Items.Add(tvi2);

                            SortedDictionary<int, string> sd = new SortedDictionary<int, string>();

                            foreach (DataRow dtRow2 in dtSet.Tables[1].Rows)//по материалам
                            {
                                string sklad = dtRow2["Склад"].ToString();

                                if (sklad.Equals(obj))
                                {
                                    string matr = dtRow2["Наименование"].ToString();

                                    this.exs.Add(new Expense(dtRow2["Наименование"].ToString(), dtRow2["Единицы_измерения"].ToString(), dtRow2["Цена"].ToString(), sklad, "Материал", "", dtRow2["k"].ToString()));
                                    this.names.Add(matr);

                                    try
                                    {
                                        sd.Add(Int32.Parse(dtRow2["i"].ToString()), matr);
                                    }
                                    catch (Exception except)
                                    {
                                        MessageBox.Show(except.Message);
                                    }
                                }
                            }

                            foreach (int i in sd.Keys)
                            {
                                TreeViewItem tvi3 = this.CreateTreeViewItem(sd[i]);
                                tvi2.Items.Add(tvi3);
                            }



                            foreach (DataRow dtRow2 in dtSet.Tables[5].Rows)//по типам счетчиков
                            {
                                string type = dtRow2["Наименование"].ToString();

                                TreeViewItem tvi3 = this.CreateTreeViewItem(type);
                                tvi1.Items.Add(tvi3);

                                sd.Clear();

                                foreach (DataRow dtRow3 in dtSet.Tables[4].Rows)//по счетчикам
                                {

                                    if (obj.Equals(dtRow3["Склад"].ToString()) && type.Equals(dtRow3["Тип_счетчика"].ToString()))
                                    {

                                        string name = dtRow3["Наименование"].ToString();

                                        this.exs.Add(new Expense(dtRow3["Наименование"].ToString(), dtRow3["Единицы_измерения"].ToString(), dtRow2["Тариф"].ToString(), dtRow3["Склад"].ToString(), dtRow2["Наименование"].ToString(), dtRow3["Описание"].ToString(), dtRow3["k"].ToString()));
                                        this.names.Add(name);

                                        try
                                        {
                                            sd.Add(Int32.Parse(dtRow3["i"].ToString()), name);
                                        }
                                        catch (Exception except)
                                        {
                                            MessageBox.Show(except.Message);
                                        }

                                    }
                                }

                                foreach (int i in sd.Keys)
                                {
                                    TreeViewItem tvi4 = this.CreateTreeViewItem(sd[i]);
                                    tvi3.Items.Add(tvi4);
                                }

                            }
                            tvi1.IsExpanded = true;

                        }
                    }
                    tvi.IsExpanded = true;

                }

                (this.Objects_TreeView.Items[0] as TreeViewItem).Focus();

            }
        }
    }
}