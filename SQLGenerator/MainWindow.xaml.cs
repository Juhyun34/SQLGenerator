using Microsoft.Office.Interop.Excel;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    public class FileSet : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// 이벤트 발생
        /// </summary>
        /// <param name="propertyName">속성 명</param>
        protected virtual void OnPropertyChanged(string propertyName)
        {
            /// 이벤트 발생
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #region FileName 
        private string _FileName = string.Empty;
        public string FileName
        {
            get { return this._FileName; }
            set
            {
                if (this._FileName != value)
                {
                    this._FileName = value;
                    this.OnPropertyChanged("FileName");
                }
            }
        }
        #endregion

        #region Code 
        private string _Code;
        public string Code
        {
            get { return this._Code; }
            set
            {
                if (this._Code != value)
                {
                    this._Code = value;
                    this.OnPropertyChanged("Code");
                }
            }
        }
        #endregion
    }

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
            xExcelFilePath.Text = Properties.Settings.Default.ExcelFilePath;
            xSheetNumber.Text = Properties.Settings.Default.SheetNumber;
            xServerIP.Text = Properties.Settings.Default.ServerIP;
            xDBName.Text = Properties.Settings.Default.DatabaseName;
            xUser.Text = Properties.Settings.Default.UserID;
            xPwd.Password = Properties.Settings.Default.PassWord;
            xPort.Text = Properties.Settings.Default.Port;
            xFileList.ItemsSource = FileList;
        }
        public ObservableCollection<FileSet> FileList = new ObservableCollection<FileSet>();
        Worksheet ExcelSheet { get; set; }
        Workbook ExcelDoc { get; set; }
        Microsoft.Office.Interop.Excel.Application Application { get; set; }
        List<TableInfo> Tables { get; set; } = new List<TableInfo>();
        List<String> Seqs { get; set; } = new List<string>();
        BackgroundWorker Worker = new BackgroundWorker();
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwThreadId);
        uint processId = 0;

        #region Excel 파일 경로 버튼 클릭 이벤트
        private void OpenExcelPathBtn_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new System.Windows.Forms.OpenFileDialog();
            ofd.Filter = "엑셀 파일 (*.xlsx)|*.xlsx|엑셀 파일 (*.xls)|*.xls";

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
            if (!File.Exists(xExcelFilePath.Text))
            {
                MessageBox.Show("파일이 존재하지 않습니다.");
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
                xFileList.SelectedIndex = 0;
            }
        }
        #endregion

        #region Save Config 버튼 클릭 이벤트
        private void SaveConfBtn_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.ExcelFilePath = xExcelFilePath.Text;
            Properties.Settings.Default.SheetNumber = xSheetNumber.Text;
            Properties.Settings.Default.ServerIP = xServerIP.Text;
            Properties.Settings.Default.DatabaseName = xDBName.Text;
            Properties.Settings.Default.UserID = xUser.Text;
            Properties.Settings.Default.PassWord = xPwd.Password;
            Properties.Settings.Default.Port = xPort.Text;
            Properties.Settings.Default.Save();
            MessageBox.Show("저장 완료");
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
                    CheckSeqExist(conn);
                    var sql = FileList.First(i => i.FileName == "SQL Query");
                    MySqlCommand cmd = new MySqlCommand(sql.Code, conn);
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
                                reftables.Add(table.TableName_ENG);
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

        #region 시퀀스 존재 여부 조회 및 삭제
        private void CheckSeqExist(MySqlConnection conn)
        {
            foreach(var seq in Seqs)
            {
                if (xSqlText.Text.Contains(seq))
                {
                    var sql = $"SHOW TABLES LIKE '{seq}'";
                    MySqlCommand cmd = new MySqlCommand(sql, conn);
                    List<string> existseqs = new List<string>();
                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        reader.Read();
                        if (reader.HasRows)
                        {
                            reader.Close();
                            existseqs.Add(seq);
                            foreach (var eseq in existseqs)
                            {
                                sql = $"DROP SEQUENCE {eseq}";
                                cmd = new MySqlCommand(sql, conn);
                                cmd.ExecuteNonQuery();
                            }
                        }
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
            for (i = 1; i <= LastRow+1; i++)
            {
                if (Application.WorksheetFunction.CountA(ExcelSheet.Rows[i]) == 0)
                {
                    (ExcelSheet.Rows[i] as Microsoft.Office.Interop.Excel.Range).Delete();
                    //if (i < LastRow - 1)
                    //{
                    //    i -= 1;
                    //}
                }
            }

            for (i = 1; i <= LastColumn; i++)
            {
                if (Application.WorksheetFunction.CountA(ExcelSheet.Columns[i]) == 0)
                {
                    (ExcelSheet.Columns[i] as Microsoft.Office.Interop.Excel.Range).Delete();
                    //if (i < LastColumn - 1)
                    //{
                    //    i -= 1;
                    //}
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

            while (index < ExcelSheet.UsedRange.Rows.Count-1)
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
            if (FileList.Any())
            {
                FileList = new ObservableCollection<FileSet>();
            }
            var totalsql = string.Empty;
            var totalscript = string.Empty;
            var isoptioned = false;
            var seqsql = string.Empty;
            foreach (var table in Tables)
            {
                string sql = $"CREATE TABLE {table.TableName_ENG} (\n\t";
                string script = "TABLE "+table.TableName_ENG+" {\n";
                string foreignScript = string.Empty;
                string primarykey = string.Empty;
                string foreignkey = string.Empty;
                seqsql = string.Empty;
                List<object> row = new List<object>();
                /*
                 * rowarray[0]: 컬럼명(한글)
                 * rowarray[1]: 컬럼명(영문)
                 * rowarray[2]: 타입
                 * rowarray[3]: 길이
                 * rowarray[4]: PK
                 * rowarray[5]: FK
                 * rowarray[6]: SQ
                 * rowarray[7]: NOTNULL
                 * rowarray[8]: Default      
                 * rowarray[12]: FK테이블명
                 * rowarray[13]: FK컬럼명
                 */
                for (int rownum = table.StartRowNum; rownum < table.StartRowNum + table.RowCount; rownum++)
                {
                    row.Clear();
                    for (int columnnum = 3; columnnum <= ExcelSheet.UsedRange.Columns.Count; columnnum++)
                    {
                        var obj = (ExcelSheet.UsedRange.Cells[rownum, columnnum]).Value;
                        row.Add(obj);
                    }
                    var rowarray = row.ToArray();
                    var scriptarray = new string[4];
                    //var scriptoption = string.Empty;
                    /*컬럼명(영문) 빈칸 오류*/
                    if (rowarray[1] == null)
                    {
                        SetErrorText($"D{rownum}");
                        return;
                    }
                    /*타입 빈칸 오류*/
                    if (rowarray[2] == null)
                    {
                        SetErrorText($"E{rownum}");
                        return;
                    }
                    /*길이 빈칸 오류(datetime 제외)*/
                    if (rowarray[3] == null)
                    {
                        if (rowarray[2].ToString() != "timestamp" && rowarray[2].ToString() != "text" && rowarray[2].ToString() != "double")
                        {
                            SetErrorText($"F{rownum}");
                            return;
                        }
                    }
                    if (rowarray[3] != null)
                    {
                        var temp = rowarray[3];
                        rowarray[3] = $"({rowarray[3]})";
                    }
                    if (rowarray[4] != null)
                    {
                        /*기본키일 때 NOT NULL 조건 빈칸 오류*/
                        if (rowarray[7] == null)
                        {
                            SetErrorText($"J{rownum}");
                            return;
                        }
                        rowarray[4] = $"PRIMARY KEY ({rowarray[1]})\n";
                        primarykey = rowarray[4].ToString();
                        scriptarray[0] = "pk";
                        isoptioned = true;
                    }
                    if(rowarray[5] != null)
                    {
                        /*외래키일 때 FK테이블명, FK컬럼명 빈칸 오류*/
                        if(rowarray[12] == null)
                        {
                            SetErrorText($"O{rownum}");
                            return;
                        }
                        else if(rowarray[13] == null)
                        {
                            SetErrorText($"P{rownum}");
                            return;
                        }
                        rowarray[5] = $"FOREIGN KEY ({rowarray[1]}) REFERENCES {rowarray[12]}({rowarray[13]})\n";
                        if(string.IsNullOrEmpty(primarykey))
                        {
                            foreignkey += "\t";
                        }
                        foreignkey += rowarray[5].ToString();
                        foreignScript += $"Ref: {rowarray[12]}.{rowarray[13]} > {table.TableName_ENG}.{rowarray[1]}\n";
                    }
                    if (rowarray[6] != null)
                    {
                        rowarray[6] = string.Empty;
                        seqsql += $"CREATE SEQUENCE {table.TableName_ENG}_SEQ START WITH 1 INCREMENT BY 1;\n";
                        Seqs.Add($"{table.TableName_ENG}_SEQ");
                        //CREATE SEQUENCE MATL_FLIGHT_INFO_SEQ START WITH 1 INCREMENT BY 1;
                    }
                    if (rowarray[7] != null)
                    {
                        rowarray[7] = " NOT NULL";
                        scriptarray[2] = "not null";
                        isoptioned = true;
                    }
                    else
                    {
                        if(rowarray[2].ToString() == "timestamp")
                        {
                            rowarray[7] = " NULL DEFAULT NULL";
                        }
                    }
                    if (rowarray[8] != null)
                    {
                        var temp = rowarray[8].ToString();
                        rowarray[8] = $" DEFAULT {temp}"; 
                        if(temp == "current_timestamp()")
                        {
                            temp = $"'{temp}'";
                        }
                        scriptarray[3] = $"default: {temp}";
                        isoptioned = true;
                    }

                    sql += $"{rowarray[1]} {rowarray[2]}{rowarray[3]}{rowarray[7]}{rowarray[8]} COMMENT '{rowarray[0]}',\n\t";
                    //script += $"{scriptarray[1]} {scriptarray[2]}{scriptarray[3]} [{scriptarray[4]} {scriptarray[6]} {scriptarray[7]}]";
                    script += $"\t{rowarray[1]} {rowarray[2]}{rowarray[3]}";
                    if (isoptioned)
                    {
                        script += " [";
                    }
                    if (scriptarray[0] != null)
                    {
                        script += $"{scriptarray[0]}";
                    }
                    if (scriptarray[1] != null)
                    {
                        if (script.Substring(script.Length - 1, 1) != "[")
                        {
                            script += ", ";
                        }
                        script += $"{scriptarray[1]}";
                    }
                    if (scriptarray[2] != null)
                    {
                        if (script.Substring(script.Length - 1, 1) != "[")
                        {
                            script += ", ";
                        }
                        script += $"{scriptarray[2]}";
                    }
                    if (scriptarray[3] != null)
                    {
                        if (script.Substring(script.Length - 1, 1) != "[")
                        {
                            script += ", ";
                        }
                        script += $"{scriptarray[3]}";
                    }
                    if (isoptioned)
                    {
                        script += "]";
                    }
                    script += "\n";
                    isoptioned = false;
                }
                if (string.IsNullOrEmpty(primarykey))
                {
                    sql = sql.Substring(0, sql.Length - 3);
                    sql += "\n";
                }
                sql += $"{primarykey}{foreignkey}) COMMENT='{table.TableName_KOR}';\n";
                script += "}\n";
                script += foreignScript;
                totalsql += seqsql;
                totalsql += sql;
                totalscript += script;
                Worker.ReportProgress(Tables.IndexOf(table) * 100 / Tables.Count);
            }
            
            App.Current.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new System.Action(delegate
            {
                FileList.Add(new FileSet()
                {
                    FileName = "SQL Query",
                    Code = totalsql
                });
                FileList.Add(new FileSet()
                {
                    FileName = "Script",
                    Code = totalscript
                });
                xSqlText.Text = totalsql;
            }));

            CloseExcel();

        }
        #endregion

        #region 에러 텍스트 추가
        private void SetErrorText(string cell)
        {
            CloseExcel();
            App.Current.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new System.Action(delegate
            {
                xResultText.Text += $"Error: [{cell}] 값이 null입니다. \nSQL구문 생성을 중단합니다.\n";
            }));
        }
        #endregion

        #region 엑셀 프로세스 Kill 함수
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
        #endregion

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

        private void xFileList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (xFileList.SelectedItem == null)
                return;

            xSqlText.Text = (xFileList.SelectedItem as FileSet).Code;
        }

        private void xSqlText_TextChanged(object sender, TextChangedEventArgs e)
        {
            (xFileList.SelectedItem as FileSet).Code = xSqlText.Text;
        }
    }
}
