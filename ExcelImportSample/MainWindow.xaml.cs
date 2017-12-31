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
        enum CellName { DiskNo, Seq, RipStatus, OnAirDate, DayOfWeek, ProgramId, ProgramDisplay, StartTime, Duration, Detail }

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
            IWorkbook workbook = WorkbookFactory.Create(@"C:\Users\JuuichiHirao\Dropbox\Interest\BD番組録画TEST.xlsx");
            // C:\Users\JuuichiHirao\Dropbox\Interest\BD番組録画.xlsx

            for(int idx=0; idx < 10; idx++)
            {
                try
                {
                    Debug.Print(idx + " " + workbook.GetSheetName(idx));
                }
                catch (Exception)
                {
                    break;
                }
            }

            ISheet worksheet = workbook.GetSheet("TV録画V2");
            int lastRow = worksheet.LastRowNum;
            Debug.Print(workbook.NumberOfSheets.ToString());
            Debug.Print("lastRow " + lastRow);

            // enum CellName { DiskNo, Seq, RipStatus, OnAirDate, DayOfWeek, ProgramId, Duration, Detail }

            for (int i = 900; i <= 903; i++)
            {
                IRow row = worksheet.GetRow(i);
                string diskNo = GetStringCellData(CellName.DiskNo, row?.GetCell((int)CellName.DiskNo));
                string diskSeq = GetStringCellData(CellName.Seq, row?.GetCell((int)CellName.Seq));
                string ripStatus = GetStringCellData(CellName.RipStatus, row?.GetCell((int)CellName.RipStatus));
                string onAirDate = GetStringCellData(CellName.OnAirDate, row?.GetCell((int)CellName.OnAirDate));
                string programId = GetStringCellData(CellName.ProgramId, row?.GetCell((int)CellName.ProgramId));
                string programName = GetStringCellData(CellName.ProgramDisplay, row?.GetCell((int)CellName.ProgramDisplay));
                string startTime = GetStringCellData(CellName.StartTime, row?.GetCell((int)CellName.StartTime));
                string duration = GetStringCellData(CellName.Duration, row?.GetCell((int)CellName.Duration));
                string detail = GetStringCellData(CellName.Detail, row?.GetCell((int)CellName.Detail));

                if (startTime.Trim().Length > 0)
                    onAirDate = onAirDate + " " + startTime + ":00";

                //ICell cell = row?.GetCell((int)CellName.DiskNo);
                Debug.Print(i + "  " + diskNo + "  Seq:" + diskSeq + " Rip:" + ripStatus + "  onAirDate:" + onAirDate + "  ProgramId:" + programId + "  programName:" + programName);
                Debug.Print("    startTime:" + startTime + " duration:" + duration + "  detail:" + detail);
                //Console.WriteLine(cell?.StringCellValue);
            }

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
        private string GetStringCellData(CellName myCellName, ICell myCell)
        {
            string cellStr = "";
            if (myCell == null)
            {
                Debug.Print("myCell null");
            }
            else
            {
                if (myCell.CellType == CellType.String)
                    return myCell?.StringCellValue;
                if (myCell.CellType == CellType.Numeric)
                {
                    // セルが日付情報が単なる数値かを判定
                    if (DateUtil.IsCellDateFormatted(myCell))
                    {
                        // 日付型
                        // 本来はスタイルに合わせてフォーマットすべきだが、
                        // うまく表示できないケースが若干見られたので固定のフォーマットとして取得
                        cellStr = myCell.DateCellValue.ToString("yyyy/MM/dd");
                    }
                    else
                    {
                        // 数値型
                        cellStr = myCell.NumericCellValue.ToString();
                    }
                }
                if (myCell.CellType == CellType.Formula)
                {
                    // 下記で数式の文字列が取得される
                    //cellStr = cell.CellFormula.ToString();

                    // 数式の元となったセルの型を取得して同様の処理を行う
                    // コメントは省略
                    switch (myCell.CachedFormulaResultType)
                    {
                        case CellType.String:
                            cellStr = myCell.StringCellValue;
                            break;
                        case CellType.Numeric:

                            if (DateUtil.IsCellDateFormatted(myCell))
                            {
                                cellStr = myCell.DateCellValue.ToString("yyyy/MM/dd HH:mm:ss");
                            }
                            else
                            {
                                cellStr = myCell.NumericCellValue.ToString();
                            }
                            break;
                        case CellType.Boolean:
                            cellStr = myCell.BooleanCellValue.ToString();
                            break;
                        case CellType.Blank:
                            break;
                        case CellType.Error:
                            cellStr = myCell.ErrorCellValue.ToString();
                            break;
                        case CellType.Unknown:
                            break;
                        default:
                            break;
                    }
                }
                if (cellStr.Length <= 0)
                    Debug.Print(myCellName.ToString() + " 変換できず、 " + myCell.CellType);
            }

            return cellStr;

        }
    }
}
