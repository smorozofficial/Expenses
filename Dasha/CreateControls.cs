using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Dasha
{
    public partial class MainWindow : Window
    {
        #region Создание контролов
        /// <summary>
        /// 
        /// </summary>
        /// <param name="header"></param>
        /// <returns></returns>
        private TreeViewItem CreateTreeViewItem(string header)
        {
            TreeViewItem t = new TreeViewItem();
            t.Header = header;
            t.PreviewMouseDoubleClick += t_PreviewMouseDoubleClick;
            t.Selected += t_Selected;
            t.PreviewKeyDown += t_PreviewKeyDown;
            return t;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="header"></param>
        /// <param name="readOnly"></param>
        /// <returns></returns>
        private DataColumn CreateColumn(string header, bool readOnly)
        {
            DataColumn col = new DataColumn();
            col.ColumnName = header;
            col.ReadOnly = readOnly;
            return col;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private TextBlock CreateTextBlock(string text)
        {
            TextBlock t = new TextBlock();
            t.Padding = new Thickness(5, 0, 5, 0);
            t.Margin = new Thickness(0, 0, 0, 0);
            t.Text = text;
            t.PreviewMouseDown += T_PreviewMouseDown;
            t.Width = 315;

            return t;
        }


        
        #endregion
    }
}