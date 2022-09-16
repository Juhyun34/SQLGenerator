using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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

namespace SQLGenerator
{
    public class TableInfo
    {
        public int StartRowNum { get; set; }
        public int RowCount { get; set; }
        public string TableName_KOR { get; set; }
        public string TableName_ENG { get; set; }
    }

    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        Worksheet ExcelSheet { get; set; }
        Workbook ExcelDoc { get; set; }
        Microsoft.Office.Interop.Excel.Application Application { get; set; }
        List<TableInfo> Tables { get; set; } = new List<TableInfo>();
        BackgroundWorker Worker = new BackgroundWorker();
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwThreadId);
        uint processId = 0;

        #region Excel 파일 경로 버튼 클릭 이벤트
        private void OpenExcelPathBtn_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.Filter = "엑셀 파일 (*.xls)|*.xls|엑셀 파일 (*.xlsx)|*.xlsx";

            var dr = ofd.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                xExcelFilePath.Text = ofd.FileName;
            }
            ofd.Dispose();

        }
        #endregion

        #region Generate 버튼 클릭 이벤트
        private void GenBtn_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(xExcelFilePath.Text))
            {
                System.Windows.MessageBox.Show("엑셀 파일을 선택하세요.");
            }
            if (string.IsNullOrEmpty(xSheetNumber.Text))
            {
                MessageBox.Show("시트 번호를 입력하세요.");
            }
            if (string.IsNullOrEmpty(xDBName.Text))
            {
                MessageBox.Show("데이터베이스 이름을 입력하세요.");
            }
            else
            {
                xSqlText.Text = String.Empty;
                Application = new Microsoft.Office.Interop.Excel.Application();
                ExcelDoc = Application.Workbooks.Open(xExcelFilePath.Text);
                ExcelSheet = ExcelDoc.Sheets[int.Parse(xSheetNumber.Text)];

                GetWindowThreadProcessId(new IntPtr(Application.Hwnd), out processId);

                Worker = new BackgroundWorker();
                Worker.DoWork += new DoWorkEventHandler(Run);
                Worker.ProgressChanged += worker_ProgressChanged;
                Worker.RunWorkerCompleted += worker_RunWorkerCompleted;
                Worker.WorkerReportsProgress = true;
                Worker.RunWorkerAsync();
            }
        }
        #endregion

        #region Write to DB 버튼 클릭 이벤트
        private void WriteDBBtn_Click(object sender, RoutedEventArgs e)
        {
            string connectString = $"Server={xServerIP.Text};Port={xPort.Text};Database={xDBName.Text};Uid={xUser.Text};Pwd={xPwd.Password};";
            try
            {
                using (MySqlConnection conn = new MySqlConnection(connectString))
                {
                    conn.Open();
                    CheckTableExist(conn);

                    MySqlCommand cmd = new MySqlCommand(xSqlText.Text, conn);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("쿼리 입력 완료");
                }
            }
            catch (MySqlException ex)
            {
                switch (ex.Number)
                {
                    case 0:
                        MessageBox.Show("Cannot connect to server.  Contact administrator");
                        break;
                    case 1045:
                        MessageBox.Show("Invalid username/password, please try again");
                        break;
                }
            }
        }
        #endregion

        #region 복사 버튼 클릭 이벤트
        private void CopyBtn_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(xSqlText.Text);
            MessageBox.Show("클립보드에 복사되었습니다.");
        }
        #endregion

        #region 테이블 존재 여부 조회 및 삭제 (종속 테이블 포함)
        private void CheckTableExist(MySqlConnection conn)
        {
            foreach (var table in Tables)
            {
                if (!string.IsNullOrEmpty(table.TableName_ENG))
                {
                    if (xSqlText.Text.Contains(table.TableName_ENG))
                    {
                        var sql = $"SHOW TABLES LIKE '{table.TableName_ENG}'";
                        MySqlCommand cmd = new MySqlCommand(sql, conn);
                        List<string> reftables = new List<string>();
                        using (MySqlDataReader reader = cmd.ExecuteReader())
                        {
                            reader.Read();
                            if (reader.HasRows)
                            {
                                reader.Close();
                                sql = $"SELECT table_name from information_schema.key_column_usage where referenced_table_name='{table.TableName_ENG}';";
                                cmd = new MySqlCommand(sql, conn);
                                using (MySqlDataReader reader2 = cmd.ExecuteReader())
                                {
                                    while (reader2.Read())
                                    {
                                        if (reader2.HasRows)
                                        {
                                            var readtable = reader2[0].ToString();
                                            reftables.Add(readtable);
                                        }
                                    }
                                }
                                foreach (var tb in reftables)
                                {
                                    sql = $"DROP TABLE {tb}";
                                    cmd = new MySqlCommand(sql, conn);
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        };
                    }
                }
            }
        }
        #endregion

        #region 제너레이트 수행
        private void Run(object sender, DoWorkEventArgs e)
        {
            DeleteBlank();
            GetMergedCellInfo();
            Generate();
        }
        #endregion

        #region 빈 행/열 삭제
        private void DeleteBlank()
        {
            var LastRow = ExcelSheet.UsedRange.Rows.Count;
            LastRow = LastRow + ExcelSheet.UsedRange.Row - 1;

            var LastColumn = ExcelSheet.UsedRange.Columns.Count;
            LastColumn = LastColumn + ExcelSheet.UsedRange.Column - 1;

            int i = 0;
            for (i = 1; i <= LastRow; i++)
            {
                if (Application.WorksheetFunction.CountA(ExcelSheet.Rows[i]) == 0)
                {
                    (ExcelSheet.Rows[i] as Microsoft.Office.Interop.Excel.Range).Delete();
                    if (i < LastRow - 1)
                    {
                        i -= 1;
                    }
                }
            }

            for (i = 1; i <= LastColumn; i++)
            {
                if (Application.WorksheetFunction.CountA(ExcelSheet.Columns[i]) == 0)
                {
                    (ExcelSheet.Columns[i] as Microsoft.Office.Interop.Excel.Range).Delete();
                    if (i < LastColumn - 1)
                    {
                        i -= 1;
                    }
                }
            }
            ExcelDoc.Save();

        }
        #endregion

        #region 병합 셀 정보 얻기
        private void GetMergedCellInfo()
        {
            var index = 2;
            var cnt = 0;

            while (index < ExcelSheet.UsedRange.Rows.Count)
            {

                Range range = ExcelSheet.Range[$"A{index}"];
                cnt = range.MergeArea.Count;
                Tables.Add(new TableInfo()
                {
                    StartRowNum = index,
                    RowCount = cnt,
                    TableName_KOR = range[1, 1].Value,
                    TableName_ENG = ExcelSheet.Range[$"B{index}"][1, 1].Value,
                });
                index += cnt;
            }
        }
        #endregion

        #region 제너레이트
        private void Generate()
        {
            var totalsql = string.Empty;
            foreach (var table in Tables)
            {
                string sql = $"CREATE TABLE {table.TableName_ENG} (\n\t";
                string primarykey = string.Empty;
                List<object> row = new List<object>();
                for (int rownum = table.StartRowNum; rownum < table.StartRowNum + table.RowCount; rownum++)
                {
                    row.Clear();
                    for (int columnnum = 3; columnnum <= ExcelSheet.UsedRange.Columns.Count; columnnum++)
                    {
                        var obj = (ExcelSheet.UsedRange.Cells[rownum, columnnum]).Value;
                        row.Add(obj);
                    }
                    var rowarray = row.ToArray();
                    /*컬럼명(영문) 빈칸 오류*/
                    //if (rowarray[1] == null)
                    //{
                    //    //SetErrorText($"D{rownum}");
                    //  CloseExcel();
                    //    return;
                    //}
                    /*타입 빈칸 오류*/
                    //if (rowarray[2] == null)
                    //{
                    //    //SetErrorText($"E{rownum}");
                    //  CloseExcel();
                    //    return;
                    //}
                    /*길이 빈칸 오류(datetime 제외)*/
                    //if (rowarray[3] == null)
                    //{
                    //    if (rowarray[2].ToString() != "datetime")
                    //    {
                    //        //SetErrorText($"F{rownum}");
                    //  CloseExcel();
                    //        return;
                    //    }
                    //}
                    if (rowarray[6] != null)
                    {
                        rowarray[6] = " AUTO_INCREMENT";
                    }
                    if (rowarray[7] != null)
                    {
                        /*NOT NULL일 때 Default값 빈칸 오류*/
                        //if (rowarray[8] == null)
                        //{
                        //    //SetErrorText($"K{rownum}");
                        //  CloseExcel();
                        //    return;
                        //}
                        rowarray[7] = " NOT NULL";
                    }
                    if (rowarray[8] != null)
                    {
                        var temp = rowarray[8].ToString();
                        rowarray[8] = $" DEFAULT {temp}";
                    }
                    if (rowarray[4] != null)
                    {
                        /*기본키일 때 NOT NULL 조건 빈칸 오류*/
                        //if (rowarray[7] == null)
                        //{
                        //    //SetErrorText($"J{rownum}");
                        //  CloseExcel();
                        //    return;
                        //}
                        rowarray[4] = $" PRIMARY KEY ({rowarray[1]})\n";
                        primarykey = rowarray[4].ToString();
                    }
                    if (rowarray[3] != null)
                    {
                        var temp = rowarray[3];
                        rowarray[3] = $"({rowarray[3]})";
                    }
                    sql += $"{rowarray[1]} {rowarray[2]}{rowarray[3]}{rowarray[7]}{rowarray[6]}{rowarray[8]} COMMENT '{rowarray[0]}',\n\t";
                }
                if (string.IsNullOrEmpty(primarykey))
                {
                    sql = sql.Substring(0, sql.Length - 3);
                    sql += "\n";
                }
                sql += $"{primarykey}) COMMENT='{table.TableName_KOR}';\n";
                totalsql += sql;
                Worker.ReportProgress(Tables.IndexOf(table) * 100 / Tables.Count);
            }
            App.Current.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new System.Action(delegate
            {
                xSqlText.Text = totalsql;
            }));

            CloseExcel();

        }
        #endregion

        private void SetErrorText(string cell)
        {
            CloseExcel();
            App.Current.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new System.Action(delegate
            {
                xResultText.Text += $"Error: [{cell}] 값이 null입니다. \nSQL구문 생성을 중단합니다.\n";
            }));
        }

        private void CloseExcel()
        {
            if (processId != 0)
            {
                System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById((int)processId);
                excelProcess.CloseMainWindow();
                excelProcess.Refresh();
                excelProcess.Kill();
            }
            //ExcelDoc.Close();
            //Application.Quit();
        }

        #region Background Worker 관련 함수
        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            xProgressBar.Value = e.ProgressPercentage;
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            xProgressBar.Value = 100;
        }
        #endregion

    }
}
