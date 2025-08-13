using System;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace count_dead_sign
{
    public class MainForm : Form
    {
        private ComboBox comboBox;
        private Button loadFileButton;
        private Button resultButton;
        private List<string> selectedFilePaths = new List<string>();

        private List<int> morningCounts = new List<int>();
        private List<int> afternoonCounts = new List<int>();
        private List<string> room= new List<string>();

        private Dictionary<string, Dictionary<string, List<int>>> hospitalMorningCounts = new Dictionary<string, Dictionary<string, List<int>>>();
        private Dictionary<string, Dictionary<string, List<int>>> hospitalAfternoonCounts = new Dictionary<string, Dictionary<string, List<int>>>();
        private Dictionary<string, Dictionary<string, int>> hospitalTotalCounts = new Dictionary<string, Dictionary<string, int>>();
        private ProgressBar progressBar;
        private readonly object logLock = new object();
        private FlowLayoutPanel mainPanel; 
        private CircularProgressBar cpb;

        public MainForm()
        {
            SetupUI();
        }
        private void SetupUI()
    {
        // 폼 설정
        this.Text = "사망위험군 분석 툴";
        this.Width = 300;
        this.Height = 400;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = Color.WhiteSmoke;

        // 세로 정렬 패널
        mainPanel = new FlowLayoutPanel();
        mainPanel.Dock = DockStyle.Fill;
        mainPanel.FlowDirection = FlowDirection.TopDown;
        mainPanel.WrapContents = false; 
        mainPanel.Padding = new Padding(20);
        mainPanel.AutoScroll = true;
        this.Controls.Add(mainPanel);

        // 콤보박스
        comboBox = new ComboBox();
        comboBox.Width = 240;
        comboBox.Font = new Font("Segoe UI", 10, FontStyle.Regular);
        comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
        comboBox.Margin = new Padding(0, 0, 0, 10);
        comboBox.DropDownHeight = 220; // 리스트 높이를 200px로
        comboBox.IntegralHeight = false; // DropDownHeight를 정확히 적용하려면 false
        mainPanel.Controls.Add(comboBox);

        // CircularProgressBar 별도 Panel에 넣기 (가로 중앙 정렬)
        Panel cpbPanel = new Panel();
        cpbPanel.Width = mainPanel.ClientSize.Width; // 패널 폭 = mainPanel 폭
        cpbPanel.Height = 140;                        // cpb 높이 + 여백
        cpbPanel.Margin = new Padding(0, 10, 0, 70);

        cpb = new CircularProgressBar();
        cpb.Size = new Size(120, 120);
        cpb.Maximum = 100;
        cpb.Value = 0;

        // 중앙 위치 계산
        cpb.Location = new Point((cpbPanel.Width - cpb.Width) -210 / 2, 25);
        cpbPanel.Controls.Add(cpb);
        cpb.Visible = false;

        mainPanel.Controls.Add(cpbPanel);

        // 버튼 패널 (가로 배치)
        FlowLayoutPanel buttonPanel = new FlowLayoutPanel();
        buttonPanel.FlowDirection = FlowDirection.LeftToRight;
        buttonPanel.Width = mainPanel.ClientSize.Width;
        buttonPanel.Height = 50;
        buttonPanel.WrapContents = false;
        buttonPanel.Margin = new Padding(0, 0, 0, 0);

        // 폴더 선택 버튼
        loadFileButton = new Button();
        loadFileButton.Text = "폴더 선택";
        loadFileButton.Width = 110;
        loadFileButton.Height = 40;
        loadFileButton.Font = new Font("Segoe UI", 10, FontStyle.Bold);
        loadFileButton.BackColor = Color.LightSkyBlue;
        loadFileButton.ForeColor = Color.White;
        loadFileButton.FlatStyle = FlatStyle.Flat;
        loadFileButton.FlatAppearance.BorderSize = 0;
        loadFileButton.Click += LoadFileButton_Click;
        buttonPanel.Controls.Add(loadFileButton);

        // 분석 시작 버튼
        resultButton = new Button();
        resultButton.Text = "분석 시작";
        resultButton.Width = 110;
        resultButton.Height = 40;
        resultButton.Font = new Font("Segoe UI", 10, FontStyle.Bold);
        resultButton.BackColor = Color.LightSkyBlue;
        resultButton.ForeColor = Color.White;
        resultButton.FlatStyle = FlatStyle.Flat;
        resultButton.FlatAppearance.BorderSize = 0;
        resultButton.Click += ResultButton_Click;
        buttonPanel.Controls.Add(resultButton);

        mainPanel.Controls.Add(buttonPanel);
    }

        private void LoadFileButton_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "엑셀 파일이 있는 폴더를 선택하세요.";

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFolder = folderDialog.SelectedPath;

                    // 폴더 내 모든 .xlsx 파일 검색
                    selectedFilePaths = Directory.GetFiles(selectedFolder, "*.xlsx").ToList();

                    comboBox.Items.Clear();

                    foreach (string filePath in selectedFilePaths)
                    {
                        comboBox.Items.Add(Path.GetFileName(filePath));
                    }

                    if (selectedFilePaths.Count > 0)
                    {
                        comboBox.SelectedIndex = 0;
                        Log($"폴더 선택됨: {selectedFolder}");
                        Log($"엑셀 파일 {selectedFilePaths.Count}개를 불러왔습니다.");
                    }
                    else
                    {
                        Log($"선택한 폴더에 엑셀 파일이 없습니다.");
                    }
                }
            }
        }


        private async void ResultButton_Click(object sender, EventArgs e)
        {
            if (selectedFilePaths.Count == 0)
            {
                Log("먼저 엑셀 파일을 선택하세요.");
                return;
            }
            // 데이터 초기화
            hospitalMorningCounts.Clear();
            hospitalAfternoonCounts.Clear();
            hospitalTotalCounts.Clear();

            cpb.Visible = true;

            await CheckDeadSignAsync();
            var summaryForm = new SaveExcelForm(hospitalMorningCounts, hospitalAfternoonCounts, hospitalTotalCounts);
            summaryForm.ShowDialog();
        }

        private async Task CheckDeadSignAsync()
        {
            List<string> checkRoomList = new List<string> { "311", "312", "313", "315", "316", "317", "318", "319", "320" };

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string exePath = AppDomain.CurrentDomain.BaseDirectory;
            string todayDate = DateTime.Now.ToString("yyyy-MM-dd");
            string saveFolderPath = Path.Combine(exePath, "EDSD_Data", todayDate);
            if (!Directory.Exists(saveFolderPath))
                Directory.CreateDirectory(saveFolderPath);

            int totalFiles = selectedFilePaths.Count;
            int processedFiles = 0;

             // ProgressBar 초기화
            this.Invoke(new Action(() =>
            {
                cpb.Maximum = 100;
                cpb.Value = 0;
            }));

            // 각 파일별로 Task.Run으로 비동기 작업 생성
            var tasks = selectedFilePaths.Select(file => Task.Run(() =>
            {
                try
                {
                    string fileName = Path.GetFileName(file);

                    var match = Regex.Match(fileName, @"^\d{8}_(\d+_\d+)(?:_([a-zA-Z]+))?\.xlsx$");

                    if (!match.Success)
                    {
                        Log($"[실패] 파일명 패턴이 맞지 않아 처리하지 않음: {fileName}");
                        return;
                    }

                    string currentRoom = match.Groups[1].Value;
                    string hospitalCode = match.Groups[2].Success ? match.Groups[2].Value : "yn";

                    room.Add(currentRoom);

                    FileInfo fileInfo = new FileInfo(file);
                    using (var package = new ExcelPackage(fileInfo))
                    {
                        var ws = package.Workbook.Worksheets[0];
                        var headers = ws.Cells[1, 1, 1, ws.Dimension.End.Column]
                                        .Select(c => c.Text).ToList();

                        string[] requiredCols = { "breath", "heart_detection", "alive_count", "range", "sp02", "radar_rssi" };
                        if (!requiredCols.All(col => headers.Contains(col)))
                        {
                            Log($"[무시됨] 필수 열 누락: {fileName}");

                            lock (hospitalMorningCounts)
                            {
                                if (!hospitalMorningCounts.ContainsKey(hospitalCode))
                                    hospitalMorningCounts[hospitalCode] = new Dictionary<string, List<int>>();
                                if (!hospitalMorningCounts[hospitalCode].ContainsKey(currentRoom))
                                    hospitalMorningCounts[hospitalCode][currentRoom] = new List<int>();
                                hospitalMorningCounts[hospitalCode][currentRoom].Add(0);
                            }

                            lock (hospitalAfternoonCounts)
                            {
                                if (!hospitalAfternoonCounts.ContainsKey(hospitalCode))
                                    hospitalAfternoonCounts[hospitalCode] = new Dictionary<string, List<int>>();
                                if (!hospitalAfternoonCounts[hospitalCode].ContainsKey(currentRoom))
                                    hospitalAfternoonCounts[hospitalCode][currentRoom] = new List<int>();
                                hospitalAfternoonCounts[hospitalCode][currentRoom].Add(0);
                            }

                            lock (hospitalTotalCounts)
                            {
                                if (!hospitalTotalCounts.ContainsKey(hospitalCode))
                                    hospitalTotalCounts[hospitalCode] = new Dictionary<string, int>();
                                hospitalTotalCounts[hospitalCode][currentRoom] = 0;
                            }

                            return;
                        }

                        int totalRows = ws.Dimension.End.Row;
                        int dataRows = totalRows - 1;

                        lock (hospitalTotalCounts)
                        {
                            if (!hospitalTotalCounts.ContainsKey(hospitalCode))
                                hospitalTotalCounts[hospitalCode] = new Dictionary<string, int>();
                            hospitalTotalCounts[hospitalCode][currentRoom] = dataRows;
                        }

                        int col_breath = headers.IndexOf("breath") + 1;
                        int col_heart = headers.IndexOf("heart_detection") + 1;
                        int col_range = headers.IndexOf("range") + 1;
                        int col_spo2 = headers.IndexOf("sp02") + 1;
                        int col_rssi = headers.IndexOf("radar_rssi") + 1;

                        int col_timestamp = 1;

                        int col_sign1 = headers.Count + 1;
                        int col_sign2 = headers.Count + 2;
                        int col_sign3 = headers.Count + 3;

                        ws.Cells[1, col_sign1].Value = "sign1";
                        ws.Cells[1, col_sign2].Value = "sign2";
                        ws.Cells[1, col_sign3].Value = "sign3";

                        int sign3Count = 0;
                        int morningCount = 0;
                        int afternoonCount = 0;

                        for (int row = 2; row <= totalRows; row++)
                        {
                            if (row < 21) continue;

                            double meanBreath = 0;
                            int count = 0;

                            for (int i = row - 19; i <= row; i++)
                            {
                                var eVal = ws.Cells[i, col_heart].GetValue<int>();
                                var dVal = ws.Cells[i, col_breath].GetValue<double>();
                                if (eVal == 1)
                                {
                                    meanBreath += dVal;
                                    count++;
                                }
                            }

                            if (count > 0)
                                meanBreath = Math.Round(meanBreath / count, 2);

                            ws.Cells[row, col_sign1].Value = meanBreath;
                            ws.Cells[row, col_sign1].Style.Numberformat.Format = "0";

                            try
                            {
                                var d = ws.Cells[row, col_breath].GetValue<double>();
                                var e = ws.Cells[row, col_heart].GetValue<int>();
                                var g = ws.Cells[row, col_range].GetValue<double>();
                                var h = ws.Cells[row, col_spo2].GetValue<double>();
                                var j = ws.Cells[row, col_rssi].GetValue<double>();
                                var l = ws.Cells[row - 1, col_sign1].GetValue<double>();

                                int sign2 = 0;
                                string baseRoom = currentRoom.Split('_')[0];

                                if (checkRoomList.Contains(baseRoom))
                                {
                                    sign2 = (
                                        d > 0 &&
                                        e == 1 &&
                                        Math.Round(d, 2) <= 0.85 * l &&
                                        Math.Abs(g - 1.66) < 0.0001 &&
                                        (h < 95 || d < 11) &&
                                        j > 200000
                                    ) ? 1 : 0;
                                }
                                else
                                {
                                    sign2 = (
                                        d > 0 &&
                                        e == 1 &&
                                        Math.Round(d, 2) <= 0.85 * l &&
                                        Math.Abs(g - 1.77) < 0.0001 &&
                                        (h < 95 || d < 11) &&
                                        j > 200000
                                    ) ? 1 : 0;
                                }

                                ws.Cells[row, col_sign2].Value = sign2;

                                if (sign2 == 1)
                                {
                                    sign3Count++;
                                    ws.Cells[row, col_sign3].Value = sign3Count;

                                    ws.Cells[row, col_sign2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    ws.Cells[row, col_sign2].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                }
                                else
                                {
                                    ws.Cells[row, col_sign3].Value = 0;
                                }
                            }
                            catch
                            {
                                ws.Cells[row, col_sign2].Value = 0;
                                ws.Cells[row, col_sign3].Value = 0;
                            }

                            string timestampText = ws.Cells[row, col_timestamp].Text;
                            if (DateTime.TryParse(timestampText, out DateTime timestamp))
                            {
                                int sign2Val = ws.Cells[row, col_sign2].GetValue<int>();
                                if (sign2Val == 1)
                                {
                                    if (timestamp.TimeOfDay < new TimeSpan(12, 0, 0))
                                    {
                                        morningCount++;
                                    }
                                    else
                                    {
                                        afternoonCount++;
                                    }
                                }
                            }
                        }

                        string lastColLetter = ws.Cells[1, col_sign3].Address.Substring(0, 1);
                        ws.Cells[$"A1:{lastColLetter}1"].AutoFilter = true;

                        string saveFilePath = Path.Combine(saveFolderPath, fileName);
                        package.SaveAs(new FileInfo(saveFilePath));

                        Log($"[완료] {fileName} 처리 완료, 저장 경로: {saveFilePath}");
                        Log($"[{fileName}] 오전 사망위험군 총 개수: {morningCount}, 오후 사망위험군 총 개수: {afternoonCount}");

                        processedFiles++;
                         // ProgressBar 초기화
                        this.Invoke(new Action(() =>
                        {
                            //progressBar.Maximum = totalFiles;
                            cpb.Value = (int)Math.Ceiling((double)processedFiles / totalFiles * 100);
                            if (cpb.Value == 100)
                            { 
                                cpb.Visible = false;
                            }
                        }));

                        lock (hospitalMorningCounts)
                        {
                            if (!hospitalMorningCounts.ContainsKey(hospitalCode))
                                hospitalMorningCounts[hospitalCode] = new Dictionary<string, List<int>>();
                            if (!hospitalMorningCounts[hospitalCode].ContainsKey(currentRoom))
                                hospitalMorningCounts[hospitalCode][currentRoom] = new List<int>();
                            hospitalMorningCounts[hospitalCode][currentRoom].Add(morningCount);
                        }

                        lock (hospitalAfternoonCounts)
                        {
                            if (!hospitalAfternoonCounts.ContainsKey(hospitalCode))
                                hospitalAfternoonCounts[hospitalCode] = new Dictionary<string, List<int>>();
                            if (!hospitalAfternoonCounts[hospitalCode].ContainsKey(currentRoom))
                                hospitalAfternoonCounts[hospitalCode][currentRoom] = new List<int>();
                            hospitalAfternoonCounts[hospitalCode][currentRoom].Add(afternoonCount);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"[에러] {Path.GetFileName(file)} 처리 실패: {ex.Message}");
                }
            }));
            // 모든 작업 완료 후 ProgressBar 100% 유지 후 잠시 대기하고 숨기기

            Log("[완료] 모든 파일 처리 완료!");

            // 모든 작업 완료 대기
            await Task.WhenAll(tasks);
        }


        private void Log(string message)
        {
            lock (logLock)
            {
                string exePath = AppDomain.CurrentDomain.BaseDirectory;
                string todayDate = DateTime.Now.ToString("yyyy-MM-dd");

                string logFolder = Path.Combine(exePath, "log");
                if (!Directory.Exists(logFolder))
                    Directory.CreateDirectory(logFolder);

                string logFilePath = Path.Combine(logFolder, $"{todayDate}.log");

                string logMessage = $"{DateTime.Now:HH:mm:ss} - {message}";

                try
                {
                    using (StreamWriter sw = new StreamWriter(logFilePath, append: true, encoding: System.Text.Encoding.UTF8))
                    {
                        sw.WriteLine(logMessage);
                    }
                }
                catch (IOException ex)
                {
                    Console.WriteLine($"로그 파일 저장 실패: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"알 수 없는 오류 발생: {ex.Message}");
                }
            }
        }



    }
}
