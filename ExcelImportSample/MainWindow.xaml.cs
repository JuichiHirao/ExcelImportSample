using ExcelImportSample.data;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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
        enum CellNameV1 { DiskNo, Seq, RipStatus, OnAirDate, BeforeRip, Kind, Channel, ProgramId, ProgramName, ProgramDisplay, Detail, StartTime, Duration }
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

            DbClear(new DbConnection());

            TvV1(workbook);
            TvV2(workbook);
        }

        public void TvV1(IWorkbook myWorkbook)
        {
            ISheet worksheet = myWorkbook.GetSheet("TV録画");
            int lastRow = worksheet.LastRowNum;
            Debug.Print(myWorkbook.NumberOfSheets.ToString());
            Debug.Print("lastRow " + lastRow);

            for (int i = 1; i <= lastRow; i++)
            {
                IRow row = worksheet.GetRow(i);

                Record record = new Record();
                record.DiskNo = GetStringCellData(CellNameV1.DiskNo, row?.GetCell((int)CellNameV1.DiskNo));
                record.Seq = GetStringCellData(CellNameV1.Seq, row?.GetCell((int)CellNameV1.Seq));
                record.RipStatus = GetStringCellData(CellNameV1.RipStatus, row?.GetCell((int)CellNameV1.RipStatus));
                string onAirDate = GetStringCellData(CellNameV1.OnAirDate, row?.GetCell((int)CellNameV1.OnAirDate));
                record.ProgramId = GetStringCellData(CellNameV1.ProgramId, row?.GetCell((int)CellNameV1.ProgramId));
                string programName = GetStringCellData(CellNameV1.ProgramDisplay, row?.GetCell((int)CellNameV1.ProgramDisplay));
                string startTime = GetStringCellData(CellNameV1.StartTime, row?.GetCell((int)CellNameV1.StartTime));
                record.Duration = GetStringCellData(CellNameV1.Duration, row?.GetCell((int)CellNameV1.Duration));
                record.Detail = GetStringCellData(CellNameV1.Detail, row?.GetCell((int)CellNameV1.Detail));

                record.SetOnAirDate(onAirDate);

                DbExport(record, new DbConnection());

                //Debug.Print(i + "  " + record.DiskNo + "  Seq:" + record.Seq + " Rip:" + record.RipStatus + "  onAirDate:" + record.OnAirDate + "  ProgramId:" + record.ProgramId + "  programName:" + programName);
                //Debug.Print("    startTime:" + startTime + " duration:" + record.Duration + "  detail:" + record.Detail);
            }
        }

        public void TvV2(IWorkbook myWorkbook)
        {
            ISheet worksheet = myWorkbook.GetSheet("TV録画V2");
            int lastRow = worksheet.LastRowNum;
            Debug.Print(myWorkbook.NumberOfSheets.ToString());
            Debug.Print("lastRow " + lastRow);

            for (int i = 800; i <= 803; i++)
            {
                IRow row = worksheet.GetRow(i);
                Record record = new Record();
                record.DiskNo = GetStringCellData(CellName.DiskNo, row?.GetCell((int)CellName.DiskNo));
                record.Seq = GetStringCellData(CellName.Seq, row?.GetCell((int)CellName.Seq));
                record.RipStatus = GetStringCellData(CellName.RipStatus, row?.GetCell((int)CellName.RipStatus));
                string onAirDate = GetStringCellData(CellName.OnAirDate, row?.GetCell((int)CellName.OnAirDate));
                record.ProgramId = GetStringCellData(CellName.ProgramId, row?.GetCell((int)CellName.ProgramId));
                string programName = GetStringCellData(CellName.ProgramDisplay, row?.GetCell((int)CellName.ProgramDisplay));
                string startTime = GetStringCellData(CellName.StartTime, row?.GetCell((int)CellName.StartTime));
                record.Duration = GetStringCellData(CellName.Duration, row?.GetCell((int)CellName.Duration));
                record.Detail = GetStringCellData(CellName.Detail, row?.GetCell((int)CellName.Detail));

                record.SetOnAirDate(onAirDate, startTime);

                DbExport(record, new DbConnection());

                //Debug.Print(i + "  " + record.DiskNo + "  Seq:" + record.Seq + " Rip:" + record.RipStatus + "  onAirDate:" + record.OnAirDate + "  ProgramId:" + record.ProgramId + "  programName:" + programName);
                //Debug.Print("    startTime:" + startTime + " duration:" + record.Duration + "  detail:" + record.Detail);
            }
        }

        private string GetStringCellData(CellName myCellName, ICell myCell)
        {
            return GetStringCellDataCore(myCellName.ToString(), myCell);
        }
        private string GetStringCellData(CellNameV1 myCellName, ICell myCell)
        {
            return GetStringCellDataCore(myCellName.ToString(), myCell);
        }

        private string GetStringCellDataCore(string myCellName, ICell myCell)
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
                                cellStr = myCell.DateCellValue.ToString("yyyy/MM/dd");
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

        public void DbClear(DbConnection myDbCon)
        {
            string sqlCommand = "DELETE FROM RECORD ";

            SqlCommand command = new SqlCommand();

            command = new SqlCommand(sqlCommand, myDbCon.getSqlConnection());

            myDbCon.execSqlCommand(sqlCommand);
        }

        public void DbExport(Record myRecord, DbConnection myDbCon)
        {
            string sqlCommand = "INSERT INTO RECORD ";
            sqlCommand += "( DISK, SEQ, STATUS, ON_AIR_DATE, PROGRAM_ID, DURATION, DETAIL ) ";
            sqlCommand += "VALUES( @Disk, @Seq, @Status, @OnAirDate, @ProgramId, @Duration, @Detail )";

            SqlCommand command = new SqlCommand();

            command = new SqlCommand(sqlCommand, myDbCon.getSqlConnection());

            List<SqlParameter> sqlparamList = new List<SqlParameter>();

            SqlParameter sqlParam = new SqlParameter("@Disk", SqlDbType.VarChar);
            sqlParam.Value = myRecord.DiskNo;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@Seq", SqlDbType.VarChar);
            sqlParam.Value = myRecord.Seq;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@Status", SqlDbType.VarChar);
            sqlParam.Value = myRecord.RipStatus;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@OnAirDate", SqlDbType.DateTime);
            sqlParam.Value = myRecord.OnAirDate;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@ProgramId", SqlDbType.VarChar);
            sqlParam.Value = myRecord.ProgramId;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@Duration", SqlDbType.VarChar);
            sqlParam.Value = myRecord.Duration;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@Detail", SqlDbType.VarChar);
            sqlParam.Value = myRecord.Detail;
            sqlparamList.Add(sqlParam);

            myDbCon.SetParameter(sqlparamList.ToArray());
            myDbCon.execSqlCommand(sqlCommand);
        }
    }
}
