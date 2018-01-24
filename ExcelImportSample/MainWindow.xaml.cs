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
        enum CellNameProgram { Id, ChannelName, Name, AbbreviationName, Kind, RelationId, DateKind, OnAirStart, OnAirEnd, Detail }
        enum CellNameChannel { Channel, Name, BroadcastKind, RipId, Remark, VideoRate, VoiceRate }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        public void Program(IWorkbook myWorkbook)
        {
            ISheet worksheet = myWorkbook.GetSheet("番組名");
            int lastRow = worksheet.LastRowNum;
            Debug.Print(myWorkbook.NumberOfSheets.ToString());
            Debug.Print("lastRow " + lastRow);

            for (int i = 1; i <= lastRow; i++)
            {
                IRow row = worksheet.GetRow(i);

                Program program = new Program();
                program.ChannelId = GetStringCellData(CellNameProgram.Id, row?.GetCell((int)CellNameProgram.Id));
                program.Name = GetStringCellData(CellNameProgram.Name, row?.GetCell((int)CellNameProgram.Name));
                program.AbbreviationName = GetStringCellData(CellNameProgram.Name, row?.GetCell((int)CellNameProgram.Name));
                program.Kind = GetStringCellData(CellNameProgram.Kind, row?.GetCell((int)CellNameProgram.Kind));
                program.RelationId = GetStringCellData(CellNameProgram.RelationId, row?.GetCell((int)CellNameProgram.RelationId));
                program.DateKind = GetStringCellData(CellNameProgram.DateKind, row?.GetCell((int)CellNameProgram.DateKind));
                string onAirStart = GetStringCellData(CellNameProgram.OnAirStart, row?.GetCell((int)CellNameProgram.OnAirStart));
                string onAirEnd = GetStringCellData(CellNameProgram.OnAirEnd, row?.GetCell((int)CellNameProgram.OnAirEnd));
                program.Detail = GetStringCellData(CellNameProgram.Detail, row?.GetCell((int)CellNameProgram.Detail));

                program.SetOnAirStart(onAirStart);
                program.SetOnAirEnd(onAirEnd);

                DbExport(program, new DbConnection());

                Debug.Print(i + "  " + program.ChannelId + "  Name:" + program.Name + " AbbreviationName:" + program.AbbreviationName + "  Kind:" + program.Kind + "  DateKind:" + program.DateKind);
                Debug.Print("    RelationId:" + program.RelationId + " OnAirDuration:" + program.OnAirStart + "～" + program.OnAirEnd);
                Debug.Print("    Detail:" + program.Detail);
            }
        }

        public void Channel(IWorkbook myWorkbook)
        {
            ISheet worksheet = myWorkbook.GetSheet("CHANNEL");
            int lastRow = worksheet.LastRowNum;
            Debug.Print(myWorkbook.NumberOfSheets.ToString());
            Debug.Print("lastRow " + lastRow);

            for (int i = 1; i <= lastRow; i++)
            {
                IRow row = worksheet.GetRow(i);

                ChannelData channel = new ChannelData();
                channel.Channel = Convert.ToInt32(GetStringCellData(CellNameV1.DiskNo, row?.GetCell((int)CellNameChannel.Channel)));
                channel.Name = GetStringCellData(CellNameChannel.Name, row?.GetCell((int)CellNameChannel.Name));
                channel.BroadcastKind = GetStringCellData(CellNameChannel.BroadcastKind, row?.GetCell((int)CellNameChannel.BroadcastKind));
                channel.RipId = GetStringCellData(CellNameChannel.RipId, row?.GetCell((int)CellNameChannel.RipId));
                channel.VideoRate = GetStringCellData(CellNameChannel.VideoRate, row?.GetCell((int)CellNameChannel.VideoRate));
                channel.VoiceRate = GetStringCellData(CellNameChannel.VoiceRate, row?.GetCell((int)CellNameChannel.VoiceRate));
                channel.Remark = GetStringCellData(CellNameChannel.Remark, row?.GetCell((int)CellNameChannel.Remark));

                DbExport(channel, new DbConnection());

                //Debug.Print(i + "  " + channel.Channel + "  Name:" + channel.Name + " Kind:" + channel.BroadcastKind + "  RipId:" + channel.RipId + "  VideoRate:" + channel.VideoRate + "  VoiceRate:" + channel.VoiceRate);
                //Debug.Print("    Remark:" + channel.Remark);
            }
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
            ISheet worksheet = myWorkbook.GetSheet("TV録画2");
            int lastRow = worksheet.LastRowNum;
            Debug.Print(myWorkbook.NumberOfSheets.ToString());
            Debug.Print("lastRow " + lastRow);

            string beforeDiskNo = "";
            for (int i = 1; i <= lastRow; i++)
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

                if (record.ProgramId.Length <= 0 && onAirDate.Length <= 0)
                    continue;

                if (record.DiskNo == null || record.DiskNo.Length <= 0)
                {
                    if (beforeDiskNo.Length > 0)
                        record.DiskNo = beforeDiskNo;
                }
                else
                    beforeDiskNo = "";

                DbExport(record, new DbConnection());

                if (beforeDiskNo.Length <= 0)
                    beforeDiskNo = record.DiskNo + "(TEMP)";
                //Debug.Print(i + "  " + record.DiskNo + "  Seq:" + record.Seq + " Rip:" + record.RipStatus + "  onAirDate:" + record.OnAirDate + "  ProgramId:" + record.ProgramId + "  programName:" + programName);
                //Debug.Print("    startTime:" + startTime + " duration:" + record.Duration + "  detail:" + record.Detail);
            }
        }

        private string GetStringCellData(CellNameChannel myCellName, ICell myCell)
        {
            return GetStringCellDataCore(myCellName.ToString(), myCell);
        }

        private string GetStringCellData(CellName myCellName, ICell myCell)
        {
            return GetStringCellDataCore(myCellName.ToString(), myCell);
        }
        private string GetStringCellData(CellNameV1 myCellName, ICell myCell)
        {
            return GetStringCellDataCore(myCellName.ToString(), myCell);
        }

        private string GetStringCellData(CellNameProgram myCellName, ICell myCell)
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

        public void DbClearProgram(DbConnection myDbCon)
        {
            string sqlCommand = "DELETE FROM PROGRAM ";

            SqlCommand command = new SqlCommand();

            command = new SqlCommand(sqlCommand, myDbCon.getSqlConnection());

            myDbCon.execSqlCommand(sqlCommand);
        }

        public void DbClearChannel(DbConnection myDbCon)
        {
            string sqlCommand = "DELETE FROM CHANNEL ";

            SqlCommand command = new SqlCommand();

            command = new SqlCommand(sqlCommand, myDbCon.getSqlConnection());

            myDbCon.execSqlCommand(sqlCommand);
        }

        public void DbExport(ChannelData myChannel, DbConnection myDbCon)
        {
            string sqlCommand = "INSERT INTO CHANNEL ";
            sqlCommand += "( CHANNEL, NAME, BROADCAST_KIND, RIP_ID, VIDEO_RATE, VOICE_RATE, REMARK ) ";
            sqlCommand += "VALUES( @Channel, @Name, @BroadcastKind, @RipId, @VideoRate, @VoiceRate, @Remark )";

            SqlCommand command = new SqlCommand();

            command = new SqlCommand(sqlCommand, myDbCon.getSqlConnection());

            List<SqlParameter> sqlparamList = new List<SqlParameter>();

            SqlParameter sqlParam = new SqlParameter("@Channel", SqlDbType.Int);
            sqlParam.Value = myChannel.Channel;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@Name", SqlDbType.VarChar);
            sqlParam.Value = myChannel.Name;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@BroadcastKind", SqlDbType.VarChar);
            sqlParam.Value = myChannel.BroadcastKind;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@RipId", SqlDbType.VarChar);
            sqlParam.Value = myChannel.RipId;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@VideoRate", SqlDbType.VarChar);
            sqlParam.Value = myChannel.VideoRate;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@VoiceRate", SqlDbType.VarChar);
            sqlParam.Value = myChannel.VoiceRate;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@Remark", SqlDbType.VarChar);
            sqlParam.Value = myChannel.Remark;
            sqlparamList.Add(sqlParam);

            myDbCon.SetParameter(sqlparamList.ToArray());
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

        public void DbExport(Program myProgram, DbConnection myDbCon)
        {
            string sqlCommand = "INSERT INTO PROGRAM ";
            sqlCommand += "( CHANNEL_ID, NAME, ABBREVIATION_NAME, RELATION_ID, ON_AIR_START, ON_AIR_END, DETAIL, REMARK ) ";
            sqlCommand += "VALUES( @Id, @Name, @AbbeviationName, @RelationId, @OnAirStart, @OnAirEnd, @Detail, @Remark )";

            SqlCommand command = new SqlCommand();

            command = new SqlCommand(sqlCommand, myDbCon.getSqlConnection());

            List<SqlParameter> sqlparamList = new List<SqlParameter>();

            SqlParameter sqlParam = new SqlParameter("@Id", SqlDbType.Int);
            sqlParam.Value = myProgram.ChannelId;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@Name", SqlDbType.VarChar);
            sqlParam.Value = myProgram.Name;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@AbbeviationName", SqlDbType.VarChar);
            sqlParam.Value = myProgram.AbbreviationName;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@RelationId", SqlDbType.VarChar);
            sqlParam.Value = myProgram.RelationId;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@OnAirStart", SqlDbType.DateTime);
            sqlParam.Value = myProgram.OnAirStart;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@OnAirEnd", SqlDbType.DateTime);
            sqlParam.Value = myProgram.OnAirEnd;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@Detail", SqlDbType.VarChar);
            sqlParam.Value = myProgram.Detail;
            sqlparamList.Add(sqlParam);

            sqlParam = new SqlParameter("@Remark", SqlDbType.VarChar);
            sqlParam.Value = myProgram.Remark;
            sqlparamList.Add(sqlParam);

            myDbCon.SetParameter(sqlparamList.ToArray());
            myDbCon.execSqlCommand(sqlCommand);
        }

        private void OnImportExecute(object sender, RoutedEventArgs e)
        {
            // NPOI
            // （WorkbookFactory.Create()を使ってinterfaceで受け取れば、xls, xlsxの両方に対応できます）
            //IWorkbook workbook = WorkbookFactory.Create(@"C:\SHARE\TV-RECORD.xlsx");
            // IWorkbook workbook = WorkbookFactory.Create(@"tv.xlsx");
            IWorkbook workbook = WorkbookFactory.Create(@"C:\Users\JuuichiHirao\Dropbox\Interest\BD番組録画.xlsx");

            for (int idx = 0; idx < 10; idx++)
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

        private void OnImportProgramExecute(object sender, RoutedEventArgs e)
        {
            IWorkbook workbook = WorkbookFactory.Create(@"C:\Users\JuuichiHirao\Dropbox\Interest\BD番組録画.xlsx");

            for (int idx = 0; idx < 10; idx++)
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

            DbClearProgram(new DbConnection());

            Program(workbook);
        }

        private void OnImportChannelExecute(object sender, RoutedEventArgs e)
        {
            IWorkbook workbook = WorkbookFactory.Create(@"C:\Users\JuuichiHirao\Dropbox\Interest\BD番組録画.xlsx");

            for (int idx = 0; idx < 10; idx++)
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

            DbClearChannel(new DbConnection());

            Channel(workbook);
        }
    }
}
