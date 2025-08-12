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

        public MainForm()
        {
            // 폼 설정
            this.Text = "사망위험군 분석 툴";
            this.Width = 1000;
            this.Height = 400;

            // 콤보박스 설정
            comboBox = new ComboBox();
            comboBox.Location = new System.Drawing.Point(30, 30);
            comboBox.Width = 200;
            this.Controls.Add(comboBox);

            // 엑셀 파일 불러오기 버튼
            loadFileButton = new Button();
            loadFileButton.Text = "폴더 선택";
            loadFileButton.Location = new System.Drawing.Point(250, 30);
            loadFileButton.Click += LoadFileButton_Click;
            this.Controls.Add(loadFileButton);

            // 결과 처리 버튼
            resultButton = new Button();
            resultButton.Text = "시작";
            resultButton.Location = new System.Drawing.Point(330, 30);
            resultButton.AutoSize = true;
            resultButton.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            resultButton.Click += ResultButton_Click;
            this.Controls.Add(resultButton);

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

            // 모든 작업 완료 대기
            await Task.WhenAll(tasks);
        }

        private void Log(string message)
        {
            string exePath = AppDomain.CurrentDomain.BaseDirectory; // 실행 경로
            string todayDate = DateTime.Now.ToString("yyyy-MM-dd");

            // 폴더 경로
            string logFolder = Path.Combine(exePath, "log");
            if (!Directory.Exists(logFolder))
                Directory.CreateDirectory(logFolder);

            // 로그 파일 경로
            string logFilePath = Path.Combine(logFolder, $"{todayDate}.log");

            // 파일에도 저장
            try
            {
                File.AppendAllText(logFilePath, $"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                // 파일 쓰기 실패 시 콘솔 출력 (혹은 무시 가능)
                Console.WriteLine($"로그 파일 저장 실패: {ex.Message}");
            }
        }

    }
}
