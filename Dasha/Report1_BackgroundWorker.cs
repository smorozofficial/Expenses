using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Collections.Generic;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;

namespace Dasha
{
    /// <summary>
    /// 
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Report1_DoWork(object sender, DoWorkEventArgs e)
        {
            DataTable Summ = new DataTable();//итоговая таблица

            Summ.Columns.Add(new DataColumn("Наименование"));
            Summ.Columns.Add(new DataColumn("Цена"));
            Summ.Columns.Add(new DataColumn("План"));//кол-во по плану
            Summ.Columns.Add(new DataColumn("СтП"));//стоимость по плану
            Summ.Columns.Add(new DataColumn("Факт"));//кол-во по факту
            Summ.Columns.Add(new DataColumn("СтФ"));//стоимость по факту
            Summ.Columns.Add(new DataColumn("%"));//процент выполнения

            ConnectDB.Open();

            foreach (Expense ex in this.exs)
            {
                if (ex.Type.Equals("Материал"))
                {
                    string query = string.Format("SELECT Дата, РасходФакт, Расход_план FROM Данные WHERE {0}='{1}'", "Наименование", ex.Name);
                    OleDbDataAdapter cmd = new OleDbDataAdapter(query, ConnectDB);

                    DataSet dt = new DataSet();
                    cmd.Fill(dt, "Данные");//0

                    double fsummary = 0.0, psummary = 0.0;
                    foreach (DataRow dr in dt.Tables[0].Rows)
                    {
                        fsummary += Double.Parse(dr.ItemArray[1].ToString().Replace('.', ','));
                        if (dr.ItemArray[2].ToString().Length == 0)
                            continue;
                        psummary += Double.Parse(dr.ItemArray[2].ToString().Replace('.', ','));
                    }
                    Summ.Rows.Add(ex.Name, ex.Price, psummary, (ex.Price * psummary).ToString(), fsummary, (ex.Price * fsummary).ToString(), "0");
                }
            }

            ConnectDB.Close();

            this.Report1_Excel(Summ);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Report1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Report1.IsEnabled = true;

            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message, "Произошла ошибка");
            }
            else
            {

            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Summ"></param>
        private void Report1_Excel(DataTable Summ)
        {
            excelapp = new Excel.Application();
            excelapp.SheetsInNewWorkbook = 1;
            excelappworkbook = excelapp.Workbooks.Add(Type.Missing);
            excelsheets = excelappworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            excelworksheet.Activate();

            excelcells = this.Merge("A1", "G1");
            excelcells.Value2 = "Анализ затрат по материалам";

            excelcells = this.Merge("A2", "A3");
            excelcells.Value2 = "Наименование материалов";
            excelcells.ColumnWidth = 50;

            excelcells = excelworksheet.get_Range("B2", "G3");
            excelcells.Select();
            excelcells.ColumnWidth = 10;

            excelcells = this.Merge("B2", "B3");
            excelcells.Value2 = "Цена";

            excelcells = this.Merge("B2", "B3");
            excelcells = excelworksheet.get_Range("C2", "D2");
            excelcells.Select();
            ((Excel.Range)(excelapp.Selection)).Merge(Type.Missing);
            excelcells.Value2 = "План";

            ((Excel.Range)excelworksheet.Cells[3, 3]).Value2 = "Кол-во";
            ((Excel.Range)excelworksheet.Cells[3, 4]).Value2 = "Стоимость";

            excelcells = this.Merge("E2", "F2");
            excelcells.Value2 = "Факт";

            ((Excel.Range)excelworksheet.Cells[3, 5]).Value2 = "Кол-во";
            ((Excel.Range)excelworksheet.Cells[3, 6]).Value2 = "Стоимость";

            excelcells = this.Merge("G2", "G3");
            excelcells.Value2 = "%";

            excelcells = excelworksheet.get_Range("A1", "G3");
            excelcells.Select();
            excelcells.Font.Bold = true;
            excelcells.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcells.VerticalAlignment = Excel.Constants.xlCenter;

            int I = 4;
            foreach (DataRow dr in Summ.Rows)
            {
                ((Excel.Range)excelworksheet.Cells[I, 1]).Value2 = dr.ItemArray[0];
                ((Excel.Range)excelworksheet.Cells[I, 2]).Value2 = dr.ItemArray[1];
                ((Excel.Range)excelworksheet.Cells[I, 3]).Value2 = dr.ItemArray[2];
                ((Excel.Range)excelworksheet.Cells[I, 4]).Value2 = dr.ItemArray[3];
                ((Excel.Range)excelworksheet.Cells[I, 5]).Value2 = dr.ItemArray[4];
                ((Excel.Range)excelworksheet.Cells[I, 6]).Value2 = dr.ItemArray[5];
                ((Excel.Range)excelworksheet.Cells[I, 7]).Value2 = dr.ItemArray[6];
                I++;
            }

            //excelcells = excelworksheet.get_Range("A1", "G" + (I - 1));
            //excelcells.Select();
            //excelcells.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
            excelapp.Visible = true;
        }
    }
}