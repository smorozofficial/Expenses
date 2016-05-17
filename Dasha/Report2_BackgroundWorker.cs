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
        private void Report2_DoWork(object sender, DoWorkEventArgs e)
        {
            SortedSet<DateTime> dates = e.Argument as SortedSet<DateTime>;

            DataTable Summ = new DataTable();//итоговая таблица
            List<string> Names = new List<string>();//будет хранить имена счетчиков

            foreach (Expense ex in this.exs)
            {
                if (ex.Type.Equals("Электричество"))
                {
                    Summ.Columns.Add(new DataColumn(ex.Name));
                    Names.Add(ex.Name);
                }
            }


            ConnectDB.Open();
            foreach (DateTime dti in dates)
            {
                string query = string.Format("SELECT Наименование, Показания, РасходФакт, Расход_план FROM Данные WHERE {0}='{1}'", "Дата", dti.ToString());
                OleDbDataAdapter cmd = new OleDbDataAdapter(query, ConnectDB);
                DataSet dt = new DataSet();
                cmd.Fill(dt, "Данные");//0

                DataRow dr = Summ.NewRow();
                foreach (DataRow row in dt.Tables[0].Rows)
                {
                    string name = row.ItemArray[0].ToString();
                    if (Names.Contains(name))
                    {
                        dr[name] = row.ItemArray[2].ToString();
                    }
                }
                Summ.Rows.Add(dr);


            }
            ConnectDB.Close();
            
            this.Report2_Excel(Summ, dates);
        }


        private void Report2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Report2.IsEnabled = true;

            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message, "Произошла ошибка");
            }
            else
            {

            }
        }

        private void Report2_Excel(DataTable Summ, SortedSet<DateTime> dates)
        {
            excelapp = new Excel.Application();
            excelapp.SheetsInNewWorkbook = 1;
            excelappworkbook = excelapp.Workbooks.Add(Type.Missing);
            excelsheets = excelappworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            excelworksheet.Activate();
            //заголовок название отчета


            excelcells = ((Excel.Range)excelworksheet.Cells[2, 1]);
            excelcells.Value2 = "Наименование счетчиков";
            excelcells.ColumnWidth = 40;
            
            int I = 3, J = 2;
            foreach (DateTime dti in dates)
            {
                ((Excel.Range)excelworksheet.Cells[I - 1, J++]).Value2 = dti.Date.ToShortDateString();
            }

            //заголовок название отчета
            excelcells = excelapp.Range[excelworksheet.Cells[1, 1], excelworksheet.Cells[1, J - 1]];
            excelcells.Select();
            ((Excel.Range)(excelapp.Selection)).Merge(Type.Missing);
            excelcells.Value2 = string.Format("Анализ затрат по счетчикам ({0} - {1})", dates.Min.ToShortDateString(), dates.Max.ToShortDateString());

            excelcells = excelapp.Range[excelworksheet.Cells[1, 2], excelworksheet.Cells[I, J - 1]];
            excelcells.Select();
            excelcells.ColumnWidth = 10;

            excelcells = excelapp.Range[excelworksheet.Cells[1, 1], excelworksheet.Cells[2, J - 1]];
            excelcells.Select();
            excelcells.Font.Bold = true;
            excelcells.HorizontalAlignment = Excel.Constants.xlCenter;
            excelcells.VerticalAlignment = Excel.Constants.xlCenter;

            foreach (DataColumn dc in Summ.Columns)
            {
                J = 2;
                ((Excel.Range)excelworksheet.Cells[I, 1]).Value2 = dc.ColumnName;
                foreach (DataRow dr in Summ.Rows)
                {
                    ((Excel.Range)excelworksheet.Cells[I, J]).Value2 = dr.ItemArray[I - 3];
                    J++;
                }
                I++;
            }

            //excelcells = excelworksheet.get_Range("A1", "G" + (I - 1));
            //excelcells.Select();
            //excelcells.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
            excelapp.Visible = true;
        }
    }
}