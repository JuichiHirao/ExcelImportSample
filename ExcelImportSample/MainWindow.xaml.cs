using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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

namespace ExcelImportSample
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // NPOI
            // （WorkbookFactory.Create()を使ってinterfaceで受け取れば、xls, xlsxの両方に対応できます）
            //IWorkbook workbook = WorkbookFactory.Create(@"C:\SHARE\TV-RECORD.xlsx");
            // IWorkbook workbook = WorkbookFactory.Create(@"tv.xlsx");
            IWorkbook workbook = WorkbookFactory.Create(@"C:\Users\JuuichiHirao\Dropbox\Interest\BD番組録画.xlsx");
            // C:\Users\JuuichiHirao\Dropbox\Interest\BD番組録画.xlsx
            ISheet worksheet = workbook.GetSheetAt(0);
            int lastRow = worksheet.LastRowNum;
            Debug.Print(workbook.NumberOfSheets.ToString());
            Debug.Print(worksheet.SheetName);
            /*
            for (int i = 0; i <= lastRow; i++)
            {
                IRow row = worksheet.GetRow(i);
                ICell cell = row?.GetCell(0);
                Debug.Print(cell.ToString());
                //Console.WriteLine(cell?.StringCellValue);
            }
             */

            // ClosedXML
            /*
            var path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "tv.xlsx");

            //var workbook = new XLWorkbook(path, XLEventTracking.Disabled);
            //var workbook = new XLWorkbook("file:///c:/SHARE/tv.xlsx");
            var workbook = new XLWorkbook(@"C:\SHARE\TV-RECORD.xlsx");

            var worksheets = workbook.Worksheets;

            foreach(var worksheet in worksheets)
            {
                Debug.Print("Name " + worksheet.Name);
            }
             */
        }
    }
}
