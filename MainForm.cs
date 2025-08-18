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
using System.Drawing.Drawing2D;


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
        private List<string> room = new List<string>();

        private Dictionary<string, Dictionary<string, List<int>>> hospitalMorningCounts = new Dictionary<string, Dictionary<string, List<int>>>();
        private Dictionary<string, Dictionary<string, List<int>>> hospitalAfternoonCounts = new Dictionary<string, Dictionary<string, List<int>>>();
        private Dictionary<string, Dictionary<string, int>> hospitalTotalCounts = new Dictionary<string, Dictionary<string, int>>();
        private ProgressBar progressBar;
        private readonly object logLock = new object();
        private FlowLayoutPanel mainPanel;
        private CircularProgressBar cpb;
        private FlowLayoutPanel horizontalPanel;
        private RoundedButton roundButton;
        private string fileDatename;

        public MainForm()
        {
            SetupUI();
        }
        private void SetupUI()
        {
            // 폼 기본 설정
            this.Text = "사망위험군 분석 툴";
            this.Width = 320;
            this.Height = 350;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(70, 70, 70);

            // 1. 세로 정렬 메인 패널
            mainPanel = new FlowLayoutPanel();
            mainPanel.Dock = DockStyle.Fill;
            mainPanel.FlowDirection = FlowDirection.TopDown;
            mainPanel.WrapContents = false;
            mainPanel.Padding = new Padding(10);
            mainPanel.AutoScroll = true;
            mainPanel.AutoScrollMinSize = new Size(0, 0); // 수평 스크롤 강제 비활성화
            this.Controls.Add(mainPanel);

            // 폼이 로드된 후 실제 크기 계산
            this.Load += (s, e) =>
            {
                int availableWidth = mainPanel.ClientSize.Width - mainPanel.Padding.Horizontal;

                // 2. 콤보박스 + 폴더 버튼 가로 패널
                Panel comboButtonWrapper = new Panel();
                comboButtonWrapper.Height = 50;
                comboButtonWrapper.Width = availableWidth;
                comboButtonWrapper.Margin = new Padding(0, 0, 0, 0);

                horizontalPanel = new FlowLayoutPanel();
                horizontalPanel.FlowDirection = FlowDirection.LeftToRight;
                horizontalPanel.WrapContents = false;
                horizontalPanel.AutoSize = false; // AutoSize 비활성화
                horizontalPanel.Size = new Size(250, 50); // 고정 크기

                // 콤보박스
                comboBox = new ComboBox();
                comboBox.Width = 140;
                comboBox.Font = new Font("Segoe UI", 9, FontStyle.Regular);
                comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                comboBox.DropDownHeight = 200;
                comboBox.IntegralHeight = false;
                comboBox.Margin = new Padding(0, 10, 5, 0);
                horizontalPanel.Controls.Add(comboBox);

                // 폴더 선택 버튼
                RoundedButton loadFileButton = new RoundedButton();
                loadFileButton.Text = "폴더 선택";
                loadFileButton.Width = 100;   // 충분히 넓게
                loadFileButton.Height = 40;
                loadFileButton.ButtonBackColor = Color.DarkGray; // 버튼 배경
                loadFileButton.TextColor = Color.White;       // 글자색
                loadFileButton.BorderColor = Color.FromArgb(70, 70, 70);
                //loadFileButton.HoverBorderColor = Color.Red;
                loadFileButton.BorderSize = 1;
                loadFileButton.CornerRadius = 10; // 덜 둥글게
                loadFileButton.ShadowOffset = 1;
                loadFileButton.DepthOffset = 1;
                loadFileButton.Click += LoadFileButton_Click;

                horizontalPanel.Controls.Add(loadFileButton);



                // horizontalPanel 중앙 배치
                horizontalPanel.Location = new Point(
                    Math.Max(0, (comboButtonWrapper.Width - horizontalPanel.Width) / 2),
                    (comboButtonWrapper.Height - horizontalPanel.Height) / 2
                );

                comboButtonWrapper.Controls.Add(horizontalPanel);
                mainPanel.Controls.Add(comboButtonWrapper);

                // 3. CircularProgressBar 패널
                Panel cpbPanel = new Panel();
                cpbPanel.Height = 140;
                cpbPanel.Width = availableWidth;
                cpbPanel.Margin = new Padding(0, 10, 0, 30);

                cpb = new CircularProgressBar();
                cpb.Size = new Size(120, 120);
                cpb.Maximum = 100;
                cpb.Value = 0;
                cpb.Visible = false;
                cpb.Location = new Point(
                    (cpbPanel.Width - cpb.Width) / 2,
                    (cpbPanel.Height - cpb.Height) / 2
                );

                cpbPanel.Controls.Add(cpb);
                mainPanel.Controls.Add(cpbPanel);

                // 4. 분석 시작 버튼 패널
                Panel resultPanel = new Panel();
                resultPanel.Height = 50;
                resultPanel.Width = availableWidth;

                RoundedButton resultButton = new RoundedButton();
                resultButton.Text = "분석 시작";
                resultButton.Width = 100;
                resultButton.Height = 35;
                resultButton.Font = new Font("Segoe UI", 9, FontStyle.Bold);

                // 버튼 색상
                resultButton.ButtonBackColor = Color.DarkGray;;
                resultButton.TextColor = Color.White;

                // border 설정
                resultButton.BorderSize = 1; // 원하는 굵기
                resultButton.BorderColor = Color.FromArgb(70, 70, 70);

                // 모서리 둥글기
                resultButton.CornerRadius = 10;

                // 그림자 옵션 (원하면 조정 가능)
                resultButton.ShadowOffset = 2;
                resultButton.DepthOffset = 1;
                resultButton.ShadowBlur = 6;
                resultButton.HoverShadowExpand = 0; // hover 시 올라가는 효과 제거
                resultButton.ShadowColor = Color.FromArgb(40, 0, 0, 0);

                // 위치
                resultButton.Location = new Point(
                    (resultPanel.Width - resultButton.Width) / 2,
                    (resultPanel.Height - resultButton.Height) / 2
                );

                // 클릭 이벤트
                resultButton.Click += ResultButton_Click;

                // 패널에 추가
                resultPanel.Controls.Add(resultButton);

                mainPanel.Controls.Add(resultPanel);

                // 크기 변경 이벤트 추가
                this.SizeChanged += (sender, args) =>
                {
                    int newAvailableWidth = mainPanel.ClientSize.Width - mainPanel.Padding.Horizontal;
                    comboButtonWrapper.Width = newAvailableWidth;
                    cpbPanel.Width = newAvailableWidth;
                    resultPanel.Width = newAvailableWidth;

                    // 위치 재조정
                    horizontalPanel.Location = new Point(
                        Math.Max(0, (comboButtonWrapper.Width - horizontalPanel.Width) / 2),
                        (comboButtonWrapper.Height - horizontalPanel.Height) / 2
                    );

                    cpb.Location = new Point(
                        (cpbPanel.Width - cpb.Width) / 2,
                        (cpbPanel.Height - cpb.Height) / 2
                    );

                    resultButton.Location = new Point(
                        (resultPanel.Width - resultButton.Width) / 2,
                        (resultPanel.Height - resultButton.Height) / 2
                    );
                };
            };



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
            var summaryForm = new SaveExcelForm(hospitalMorningCounts, hospitalAfternoonCounts, hospitalTotalCounts, fileDatename);
            summaryForm.ShowDialog();
        }

        private async Task CheckDeadSignAsync()
        {
            List<string> checkRoomList = new List<string> { "311", "312", "313", "315", "316", "317", "318", "319", "320" };

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string exePath = AppDomain.CurrentDomain.BaseDirectory;
            string todayDate = DateTime.Now.ToString("yyyy-MM-dd");



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

                    // 날짜를 분리하기 위한 정규표현식
                    string pattern = @"^(\d{8})";
                    var filedate = Regex.Match(fileName, pattern);

                    string datePart = filedate.Groups[1].Value;

                    // yyyyMMdd 형식으로 파싱
                    DateTime date = DateTime.ParseExact(datePart, "yyyyMMdd", null);

                    // yyyy-MM-dd 형태로 변환
                    fileDatename = date.ToString("yyyy-MM-dd");

                    string saveFolderPath = Path.Combine(exePath, "EDSD_Data", fileDatename);


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

                        int totalRows = ws.Dimension.End.Row; // Excel 시트의 총 행 수 확인 (헤더 포함)
                        int dataRows = totalRows - 1; // 실제 데이터 행 수 (헤더 제외)

                        lock (hospitalTotalCounts)
                        {
                            if (!hospitalTotalCounts.ContainsKey(hospitalCode))
                                hospitalTotalCounts[hospitalCode] = new Dictionary<string, int>();
                            hospitalTotalCounts[hospitalCode][currentRoom] = dataRows;
                        }


                        // 각 컬럼의 위치(Index) 찾기
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


                        // 각 컬럼의 위치(Index) 찾기
                        int sign3Count = 0;
                        int morningCount = 0;
                        int afternoonCount = 0;

                        for (int row = 2; row <= totalRows; row++)
                        {
                            if (row < 21) continue; // 처음 20행을 계산하기 위해 21행 이전은 건너뜀

                            double meanBreath = 0; // 평균 호흡수 계산용
                            int count = 0; // // 유효 데이터 개수

                            for (int i = row - 19; i <= row; i++) // 현재 행 포함 최근 20행 반복
                            {
                                var eVal = ws.Cells[i, col_heart].GetValue<int>();
                                var dVal = ws.Cells[i, col_breath].GetValue<double>();

                                if (eVal == 1) // 심박 감지 == 1 인 경우만 평균에 포함
                                {
                                    meanBreath += dVal;
                                    count++;
                                }
                            }

                            if (count > 0)
                                meanBreath = Math.Round(meanBreath / count, 2); //heartdetection 1인 경우 호흡 평균 계산

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

                                if (checkRoomList.Contains(baseRoom)) // yn 3층인 경우
                                {
                                    sign2 = (
                                        d > 0 &&
                                        e == 1 &&
                                        Math.Round(d, 2) <= 0.85 * l &&
                                        Math.Abs(g - 1.66) < 0.0001 &&
                                        (h < 95 || d < 10) &&
                                        j > 200000
                                    ) ? 1 : 0;
                                }else if(hospitalCode == "h" || hospitalCode == "jj" || hospitalCode == "gj"){
                                    sign2 = (
                                        d > 0 &&
                                        e == 1 &&
                                        Math.Round(d, 2) <= 0.85 * l &&
                                        Math.Abs(g - 1.44) < 0.0001 &&
                                        (h < 95 || d < 10) &&
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
                                        (h < 95 || d < 10) &&
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

                        if (!Directory.Exists(saveFolderPath))
                            Directory.CreateDirectory(saveFolderPath);

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
