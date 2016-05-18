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
        private void Report3_DoWork(object sender, DoWorkEventArgs e)
        {
            SortedSet<DateTime> dates = e.Argument as SortedSet<DateTime>;

            DataTable Summ = new DataTable();//итоговая таблица
            List<string> Names = new List<string>();//будет хранить имена счетчиков

            foreach (Expense ex in this.exs)
            {
                if (ex.Type.Equals("Вода"))
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

                foreach (DataColumn dc in Summ.Columns)
                {
                    dr[dc.ColumnName] = "";
                }

                foreach (DataRow row in dt.Tables[0].Rows)
                {
                    string name = row.ItemArray[0].ToString();
                    if (Names.Contains(name))
                    {
                        //расходФАкт, показания, расходплан
                        dr[name] += row.ItemArray[2] + ";" + row.ItemArray[1] + ";" + row.ItemArray[3];
                    }
                }
                Summ.Rows.Add(dr);
            }
            ConnectDB.Close();

            this.Report3_Excel(Summ, dates);
        }


        private void Report3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
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

        private void Report3_Excel(DataTable Summ, SortedSet<DateTime> dates)
        {
            excelapp = new Excel.Application();
            excelapp.SheetsInNewWorkbook = 2;
            excelappworkbook = excelapp.Workbooks.Add(Type.Missing);
            excelsheets = excelappworkbook.Worksheets;
            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(2);
            excelworksheet.Name = "Показания";
            excelworksheet.Activate();


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
                ((Excel.Range)excelworksheet.Cells[I + 1, 1]).Value2 = dc.ColumnName + " показания";
                ((Excel.Range)excelworksheet.Cells[I + 1, 1]).Font.ColorIndex = 5;//синий цвет для показаний
                foreach (DataRow dr in Summ.Rows)
                {
                    string s = dr.ItemArray[(I - 3) / 2].ToString();
                    if (s.Length > 0)
                    {
                        ((Excel.Range)excelworksheet.Cells[I, J]).Value2 = s.Split(';')[0];
                        ((Excel.Range)excelworksheet.Cells[I + 1, J]).Value2 = s.Split(';')[1];
                        ((Excel.Range)excelworksheet.Cells[I + 1, J]).Font.ColorIndex = 5;//синий цвет для показаний
                    }

                    J++;
                }
                I += 2;
            }

            excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
            excelworksheet.Name = "Расход";
            excelworksheet.Activate();


            excelcells = ((Excel.Range)excelworksheet.Cells[2, 1]);
            excelcells.Value2 = "Наименование счетчиков";
            excelcells.ColumnWidth = 40;

            I = 3; J = 2;
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
                ((Excel.Range)excelworksheet.Cells[I + 1, 1]).Value2 = dc.ColumnName + " план";
                ((Excel.Range)excelworksheet.Cells[I + 1, 1]).Font.ColorIndex = 5;//синий цвет для плана
                foreach (DataRow dr in Summ.Rows)
                {
                    string s = dr.ItemArray[(I - 3) / 2].ToString();
                    if (s.Length > 0)
                    {
                        ((Excel.Range)excelworksheet.Cells[I, J]).Value2 = s.Split(';')[0];
                        ((Excel.Range)excelworksheet.Cells[I + 1, J]).Value2 = s.Split(';')[2];
                        ((Excel.Range)excelworksheet.Cells[I + 1, J]).Font.ColorIndex = 5;//синий цвет для плана
                    }

                    J++;
                }
                I += 2;
            }

            //excelcells = excelworksheet.get_Range("A1", "G" + (I - 1));
            //excelcells.Select();
            //excelcells.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic);
            excelapp.Visible = true;
        }
    }
}