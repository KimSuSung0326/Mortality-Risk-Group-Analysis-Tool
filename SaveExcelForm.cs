using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;
using System.Threading.Tasks;
using System.Diagnostics;

namespace count_dead_sign
{
    public partial class SaveExcelForm : Form
    {
        // 멤버 변수 (병원코드 -> (호실 -> List<int>))
        private Dictionary<string, Dictionary<string, List<int>>> hospitalMorningCounts;
        private Dictionary<string, Dictionary<string, List<int>>> hospitalAfternoonCounts;
        private Dictionary<string, Dictionary<string, int>> hospitalTotalCounts;

        // UI 컨트롤
        private Label lblDate;
        private TreeView treeViewStats;
        private Button btnSave;

        // RoomData 클래스 정의
        class RoomData
        {
            public string Room { get; set; }
            public int MorningSum { get; set; }
            public int AfternoonSum { get; set; }
        }

        public SaveExcelForm(
            Dictionary<string, Dictionary<string, List<int>>> morningCounts,
            Dictionary<string, Dictionary<string, List<int>>> afternoonCounts,
            Dictionary<string, Dictionary<string, int>> totalCounts)
        {
            hospitalMorningCounts = morningCounts;
            hospitalAfternoonCounts = afternoonCounts;
            hospitalTotalCounts = totalCounts;

            InitializeComponent();
            DisplayStats();
        }

        private void InitializeComponent()
        {
            this.Text = "사망 위험군 요약";
            this.Size = new Size(420, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Font = new Font("맑은 고딕", 10);

            lblDate = new Label()
            {
                Left = 0,
                Top = 10,
                Width = this.ClientSize.Width,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("맑은 고딕", 12, FontStyle.Bold)
            };
            this.Controls.Add(lblDate);

            treeViewStats = new TreeView()
            {
                Left = 10,
                Top = 50,
                Width = this.ClientSize.Width - 20,
                Height = this.ClientSize.Height - 120,
                Font = new Font("맑은 고딕", 10),
            };
            this.Controls.Add(treeViewStats);

            btnSave = new Button()
            {
                Text = "엑셀 파일로 저장",
                Width = 180,
                Height = 40,
                Top = this.ClientSize.Height - 50,
                BackColor = Color.LightSteelBlue
            };
            btnSave.Left = (this.ClientSize.Width - btnSave.Width) / 2;
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);
        }

        private void DisplayStats()
        {
            lblDate.Text = $"오늘 날짜: {DateTime.Now:yyyy-MM-dd}";
            treeViewStats.Nodes.Clear();

            // 오전/오후 위험군 건수 노드
            TreeNode morningRoot = new TreeNode("오전 사망위험군 건수");
            TreeNode afternoonRoot = new TreeNode("오후 사망위험군 건수");

            foreach (var hospitalEntry in hospitalMorningCounts)
            {
                string hospitalCode = hospitalEntry.Key;
                var morningDict = hospitalEntry.Value;
                TreeNode morningHospitalNode = new TreeNode(hospitalCode);

                int hospitalMorningDangerSum = 0;
                int hospitalMorningTotalSum = 0;

                foreach (var roomEntry in morningDict)
                {
                    string room = roomEntry.Key;
                    var counts = roomEntry.Value;
                    int dangerCount = counts.Count(x => x > 3);
                    int totalCount = counts.Count();

                    hospitalMorningDangerSum += dangerCount;
                    hospitalMorningTotalSum += totalCount;

                    string dangerYN = dangerCount > 0 ? "O" : "X";
                    morningHospitalNode.Nodes.Add(new TreeNode($"{room} : 위험군 {dangerYN}"));
                }

                morningHospitalNode.Text += $"  (합계: 위험군 {hospitalMorningDangerSum} / 전체 {hospitalMorningTotalSum})";
                morningRoot.Nodes.Add(morningHospitalNode);

                if (hospitalAfternoonCounts.TryGetValue(hospitalCode, out var afternoonDict))
                {
                    TreeNode afternoonHospitalNode = new TreeNode(hospitalCode);
                    int hospitalAfternoonDangerSum = 0;
                    int hospitalAfternoonTotalSum = 0;

                    foreach (var roomEntry in afternoonDict)
                    {
                        string room = roomEntry.Key;
                        var counts = roomEntry.Value;
                        int dangerCount = counts.Count(x => x > 3);
                        int totalCount = counts.Count();

                        hospitalAfternoonDangerSum += dangerCount;
                        hospitalAfternoonTotalSum += totalCount;

                        string dangerYN = dangerCount > 0 ? "O" : "X";
                        afternoonHospitalNode.Nodes.Add(new TreeNode($"{room} : 위험군 {dangerYN}"));
                    }

                    afternoonHospitalNode.Text += $"  (합계: 위험군 {hospitalAfternoonDangerSum} / 전체 {hospitalAfternoonTotalSum})";
                    afternoonRoot.Nodes.Add(afternoonHospitalNode);
                }
                else
                {
                    afternoonRoot.Nodes.Add(new TreeNode($"{hospitalCode} (데이터 없음)"));
                }
            }

            TreeNode morningRateRoot = new TreeNode("오전 사망위험군 비율");
            TreeNode afternoonRateRoot = new TreeNode("오후 사망위험군 비율");

            foreach (var hospitalEntry in hospitalMorningCounts)
            {
                string hospitalCode = hospitalEntry.Key;
                var morningDict = hospitalEntry.Value;

                int hospitalMorningDangerSum = 0;
                int hospitalMorningTotalSum = 0;

                foreach (var roomEntry in morningDict)
                {
                    var counts = roomEntry.Value;
                    hospitalMorningDangerSum += counts.Count(x => x > 3);
                    hospitalMorningTotalSum += counts.Count();
                }

                double morningRate = hospitalMorningTotalSum > 0 ? (double)hospitalMorningDangerSum / hospitalMorningTotalSum * 100 : 0;
                TreeNode morningRateHospitalNode = new TreeNode($"{hospitalCode} : 위험군 비율 {morningRate:F1}%");
                morningRateRoot.Nodes.Add(morningRateHospitalNode);

                if (hospitalAfternoonCounts.TryGetValue(hospitalCode, out var afternoonDict))
                {
                    int hospitalAfternoonDangerSum = 0;
                    int hospitalAfternoonTotalSum = 0;

                    foreach (var roomEntry in afternoonDict)
                    {
                        var counts = roomEntry.Value;
                        hospitalAfternoonDangerSum += counts.Count(x => x > 3);
                        hospitalAfternoonTotalSum += counts.Count();
                    }

                    double afternoonRate = hospitalAfternoonTotalSum > 0 ? (double)hospitalAfternoonDangerSum / hospitalAfternoonTotalSum * 100 : 0;
                    TreeNode afternoonRateHospitalNode = new TreeNode($"{hospitalCode} : 위험군 비율 {afternoonRate:F1}%");
                    afternoonRateRoot.Nodes.Add(afternoonRateHospitalNode);
                }
                else
                {
                    afternoonRateRoot.Nodes.Add(new TreeNode($"{hospitalCode} (데이터 없음)"));
                }
            }

            treeViewStats.Nodes.Add(morningRoot);
            treeViewStats.Nodes.Add(afternoonRoot);
            treeViewStats.Nodes.Add(morningRateRoot);
            treeViewStats.Nodes.Add(afternoonRateRoot);

            treeViewStats.BeginUpdate();
            treeViewStats.Nodes[0].Collapse();
            treeViewStats.Nodes[1].Collapse();
            treeViewStats.Nodes[2].Expand();
            treeViewStats.Nodes[3].Expand();
            treeViewStats.EndUpdate();
        }

        private async void BtnSave_Click(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            try
            {
                await saveExcelFile();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"저장 중 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnSave.Enabled = true;
            }
        }

        private async Task saveExcelFile()
        {
            try
            {
                
                string todayDate = DateTime.Now.ToString("yyyy-MM-dd");
                string folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "사망위험군_요약");
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);

                string filePath = Path.Combine(folder, $"사망위험군_발생_요약({todayDate}).xlsx");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = File.Exists(filePath)
                    ? new ExcelPackage(new FileInfo(filePath))
                    : new ExcelPackage())
                {
                    await Task.Run(() =>
                    {
                        foreach (var hospitalEntry in hospitalMorningCounts)
                        {
                            string hospitalCode = hospitalEntry.Key;

                            var morningDict = hospitalMorningCounts[hospitalCode];
                            var afternoonDict = hospitalAfternoonCounts.ContainsKey(hospitalCode)
                                                ? hospitalAfternoonCounts[hospitalCode]
                                                : new Dictionary<string, List<int>>();

                            ExcelWorksheet ws = package.Workbook.Worksheets[hospitalCode];
                            
                            if (ws == null)
                                ws = package.Workbook.Worksheets.Add(hospitalCode);
                            
                            ws.View.ZoomScale = 70;// 엑셀 파일 퍼센트 70% 설정

                            // 날짜 타이틀 병합
                            ws.Cells[1, 1].Value = $"오늘날짜: {todayDate}";
                            ws.Cells[1, 1, 1, 4].Merge = true;
                            ws.Cells[1, 1, 1, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            ws.Cells[1, 1, 1, 4].Style.Font.Bold = true;
                            ws.Row(1).Height = 24;

                            // 헤더 작성
                            ws.Cells[2, 1].Value = "위험도 구분";
                            ws.Cells[2, 2].Value = "호실";
                            ws.Cells[2, 3].Value = "오전";
                            ws.Cells[2, 4].Value = "오후";
                            ws.Cells[2, 1, 2, 4].Style.Font.Bold = true;
                            ws.Cells[2, 1, 2, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            ws.Row(2).Height = 20;
                            ws.Cells[2, 1, 2, 4].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Cells[2, 1, 2, 4].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);

                            var rooms = morningDict.Keys.OrderBy(r => r).ToList();

                            var roomDataList = rooms.Select(room => new RoomData
                            {
                                Room = room,
                                MorningSum = morningDict[room].Sum(),
                                AfternoonSum = afternoonDict.ContainsKey(room) ? afternoonDict[room].Sum() : 0
                            }).ToList();

                            //var highRisk = roomDataList.Where(r => r.MorningSum >= 10 || r.AfternoonSum >= 10).ToList();
                            //var midRisk = roomDataList.Where(r => (r.MorningSum >= 4 && r.MorningSum <= 9) || (r.AfternoonSum >= 4 && r.AfternoonSum <= 9)).ToList();
                            //var lowRisk = roomDataList.Where(r => r.MorningSum <= 3 && r.AfternoonSum <= 3).ToList();

                            int totalCount = roomDataList.Count(); // 각 병원 별 병실 전체 카운트 값

                            int currentRow = 3;

        int WriteGroup(string groupName, List<RoomData> groupList, int totalCount, int startCol, int startRow)
        {
            if (groupList.Count == 0)
                return startRow;
                
            int currentRow = startRow;
            
            // 위험군별 배경색 설정
            Color bgColor = Color.White;
            if (groupName.Contains("고위험")) bgColor = Color.FromArgb(255, 200, 200);      // 연한 빨강
            else if (groupName.Contains("중위험")) bgColor = Color.FromArgb(255, 255, 200); // 연한 노랑
            else if (groupName.Contains("저위험")) bgColor = Color.FromArgb(200, 255, 200);    // 연한 초록
            
            // 그룹명을 startCol 위치에 설정 (모든 위험군 공통)
            ws.Cells[currentRow, startCol].Value = groupName;
            
            // 데이터 처리
            foreach (var item in groupList)
            {
                // 병실명
                ws.Cells[currentRow, startCol + 1].Value = item.Room;
                
                // 위험군별 조건에 따른 데이터 입력
                if (groupName.Contains("고위험"))
                {
                    ws.Cells[currentRow, startCol + 2].Value = item.MorningSum > 9 ? (object)item.MorningSum : "";
                    ws.Cells[currentRow, startCol + 3].Value = item.AfternoonSum > 9 ? (object)item.AfternoonSum : "";
                }
                else if (groupName.Contains("중위험"))
                {
                    ws.Cells[currentRow, startCol + 2].Value = (3 < item.MorningSum && item.MorningSum < 10) ? (object)item.MorningSum : "";
                    ws.Cells[currentRow, startCol + 3].Value = (3 < item.AfternoonSum && item.AfternoonSum < 10) ? (object)item.AfternoonSum : "";
                }
                else if (groupName.Contains("저위험"))
                {
                    ws.Cells[currentRow, startCol + 2].Value = item.MorningSum < 4 ? (object)item.MorningSum : "";
                    ws.Cells[currentRow, startCol + 3].Value = item.AfternoonSum < 4 ? (object)item.AfternoonSum : "";
                }
                
                // 홀짝 줄 배경색 (줄무늬 효과)
                var rowColor = (currentRow % 2 == 0) ? Color.White : Color.FromArgb(240, 240, 240);
                ws.Cells[currentRow, startCol + 1, currentRow, startCol + 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells[currentRow, startCol + 1, currentRow, startCol + 3].Style.Fill.BackgroundColor.SetColor(rowColor);
                
                // 정렬
                ws.Cells[currentRow, startCol + 2, currentRow, startCol + 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                ws.Cells[currentRow, startCol + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                
                currentRow++;
            }
            
            // 그룹명 병합 및 스타일 (데이터 처리 후에 적용)
            if (groupList.Count > 0)
            {
                ws.Cells[startRow, startCol, currentRow - 1, startCol].Merge = true;
                ws.Cells[startRow, startCol, currentRow - 1, startCol].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
                ws.Cells[startRow, startCol, currentRow - 1, startCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[startRow, startCol, currentRow - 1, startCol].Style.Font.Bold = true;
                ws.Cells[startRow, startCol, currentRow - 1, startCol].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells[startRow, startCol, currentRow - 1, startCol].Style.Fill.BackgroundColor.SetColor(Color.White);
            }
            
            // 통계 계산
            int morningCount = 0, afternoonCount = 0;
            if (groupName.Contains("고위험"))
            {
                morningCount = groupList.Count(x => x.MorningSum > 9);
                afternoonCount = groupList.Count(x => x.AfternoonSum > 9);
            }
            else if (groupName.Contains("중위험"))
            {
                morningCount = groupList.Count(x => 3 < x.MorningSum && x.MorningSum < 10);
                afternoonCount = groupList.Count(x => 3 < x.AfternoonSum && x.AfternoonSum < 10);
            }
            else if (groupName.Contains("저위험"))
            {
                morningCount = groupList.Count(x => x.MorningSum < 4);
                afternoonCount = groupList.Count(x => x.AfternoonSum < 4);
            }
            
            // 발생 횟수 행
            ws.Cells[currentRow, startCol].Value = $"{groupName} 발생 횟수";
            ws.Cells[currentRow, startCol + 2].Value = $"{morningCount} 회";
            ws.Cells[currentRow, startCol + 3].Value = $"{afternoonCount} 회";
            ws.Cells[currentRow, startCol, currentRow, startCol + 3].Style.Font.Bold = true;
            ws.Cells[currentRow, startCol, currentRow, startCol + 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells[currentRow, startCol, currentRow, startCol + 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells[currentRow, startCol, currentRow, startCol + 3].Style.Fill.BackgroundColor.SetColor(bgColor);
            currentRow++;
            
            // 발생 비율 행
            ws.Cells[currentRow, startCol].Value = $"{groupName} 발생 비율 (%)";
            ws.Cells[currentRow, startCol + 2].Value = totalCount > 0 ? (double)morningCount / totalCount : 0;
            ws.Cells[currentRow, startCol + 3].Value = totalCount > 0 ? (double)afternoonCount / totalCount : 0;
            ws.Cells[currentRow, startCol + 2, currentRow, startCol + 3].Style.Numberformat.Format = "0.0%";
            ws.Cells[currentRow, startCol, currentRow, startCol + 3].Style.Font.Bold = true;
            ws.Cells[currentRow, startCol, currentRow, startCol + 3].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells[currentRow, startCol, currentRow, startCol + 3].Style.Fill.BackgroundColor.SetColor(bgColor);
            ws.Cells[currentRow, startCol, currentRow, startCol + 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            currentRow++;
            
            // 그룹 경계 테두리
            var groupRange = ws.Cells[startRow, startCol, currentRow - 1, startCol + 3];
            groupRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            groupRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            groupRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            groupRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            
            currentRow++; // 그룹 간 한 줄 공백
            return currentRow;
        }

                            //int totalCounts = roomDataList.Select(x => x.Room).Distinct().Count();
                            int row = 3;
                            WriteGroup("고위험군 (10회 이상)", roomDataList, totalCount,1,row);
                            WriteGroup("중위험군 (4~9회)", roomDataList, totalCount,6,row);
                            WriteGroup("저위험군 (0~3회)", roomDataList, totalCount,11,row);

                      
                            currentRow++;

                            // 열 너비 자동 조절
                            ws.Cells[ws.Dimension.Address].AutoFitColumns();
                        }

                        package.SaveAs(new FileInfo(filePath));
                    });
                }

                MessageBox.Show($"엑셀 저장 완료!\n\n{filePath}", "저장 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Process.Start("explorer.exe", Path.GetDirectoryName(filePath));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"엑셀 저장 실패: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }
}
