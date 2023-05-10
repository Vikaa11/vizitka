using Microsoft.Office.Interop.Excel;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace визитка
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow 
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application(); Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet; //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(@"C:\Users\frenb\OneDrive\Рабочий стол\цвцв.xlsx"); //Таблица.
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            ExcelWorkSheet.Cells[1, "B"] = Surname.Text;
            ExcelWorkSheet.Cells[2, "B"] = Name.Text;
            ExcelWorkSheet.Cells[3, "B"] = phone.Text;
            ExcelWorkSheet.Cells[4, "B"] = Email.Text;
            ExcelWorkSheet.Cells[5, "B"] = name1.Text;
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;

        }
    }
}
