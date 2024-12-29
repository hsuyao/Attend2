using System;
using System.IO;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using CellType = NPOI.SS.UserModel.CellType;
using IndexedColors = NPOI.SS.UserModel.IndexedColors;
using NPOI.HPSF;
using System.Text;
using System.Data;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Drawing;
using NPOI.SS.Formula.Functions;
using NPOI.HSSF.Record;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Text.RegularExpressions;
using System.Reflection;


namespace Attend;

public partial class AttendForm : Form
{
    HSSFWorkbook workbook;
    private List<AttendanceRecord> records;
    private Dictionary<string, Dictionary<string, object>> attendanceSummary;
    public List<string> FileNames { get; private set; }
    ColorDialog colorDialog;
    Size originalFormSize;
    Size originalTabControlSize;
    Size originalDataGridViewSize;
    public AttendForm()
    {
        InitializeComponent();
        // 取得編譯日期
        DateTime compileDate = GetCompileDate();
        // 設定 Form 的標題欄文字，只顯示日期
        this.Text = $"召會點名系統統計小幫手 v1.0 Build {compileDate.ToString("yyyyMMdd")}";

        colorDialog = new ColorDialog();
        colorDialog.FullOpen = true;
        FileNames = new List<string>();
        // 記錄原始的大小
        originalFormSize = this.Size;
        originalTabControlSize = tabControl1.Size;
        originalDataGridViewSize = tabControl1.Size; // dgvResult1.Size;
        LoadAttendanceRecords();
        attendanceSummary = new Dictionary<string, Dictionary<string, object>>();
       // CalculateAttendanceSummary();
    }
    private DateTime GetCompileDate()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var filePath = assembly.Location;
        var fileInfo = new System.IO.FileInfo(filePath);
        return fileInfo.LastWriteTime;
    }

    private void CalculateAttendanceSummary()
    {
        attendanceSummary.Clear();

        foreach (var record in records)
        {
            if (!attendanceSummary.ContainsKey(record.Name))
            {
                attendanceSummary[record.Name] = new Dictionary<string, object>
            {
                { "TotalAttendance", 0 },
                { "LastAttendDate", "無近期紀錄" }
            };
            }

            attendanceSummary[record.Name]["TotalAttendance"] = (int)attendanceSummary[record.Name]["TotalAttendance"] + record.Attendance;

            if (record.Attendance == 1)
            {
                attendanceSummary[record.Name]["LastAttendDate"] = record.Date;
            }
        }
    }
    public void AddNewRecord(AttendanceRecord newRecord)
    {
        AttendanceRecordComparer comparer = new AttendanceRecordComparer();
        int existingIndex = records.FindIndex(r => comparer.Equals(r, newRecord));

        if (existingIndex != -1)
        {
            // 记录已存在，更新 Attendance
            records[existingIndex].Attendance = newRecord.Attendance;
        }
        else
        {
            // 检查相同姓名的记录数量是否少于 52
            int count = records.Count(r => r.Name == newRecord.Name);
            if (count < 52*4000)
            {
                records.Add(newRecord);
            }
            else
            {
                // 找到并移除最舊记录
                var oldestRecord = records.Where(r => r.Name == newRecord.Name)
                    .OrderBy(r => ParseCustomDate(r.Date))
                    .FirstOrDefault();

                if (oldestRecord != null)
                {
                    // 比较 newRecord 和 oldestRecord 的日期
                    if (ParseCustomDate(newRecord.Date) > ParseCustomDate(oldestRecord.Date))
                    {
                        // 仅当 newRecord 的日期较新时才移除和新增
                        records.Remove(oldestRecord);
                        records.Add(newRecord);
                    }
                    // 若 newRecord 较旧，则不进行任何操作
                }
            }
        }

        // 更新 UI 或執行其他操作
    }

    private DateTime ParseCustomDate(string date)
    {
        // Example: "2024年5月第一週"
        var parts = date.Split(new[] { '年', '月' }, StringSplitOptions.RemoveEmptyEntries);
        int year = int.Parse(parts[0]);
        int month = int.Parse(parts[1]);
        string weekPart = parts[2];

        // Determine the week number
        int weekNumber = 1;
        if (weekPart.Contains("第二週"))
        {
            weekNumber = 2;
        }
        else if (weekPart.Contains("第三週"))
        {
            weekNumber = 3;
        }
        else if (weekPart.Contains("第四週"))
        {
            weekNumber = 4;
        }
        else if (weekPart.Contains("第五週"))
        {
            weekNumber = 5;
        }

        // Calculate the date of the first day of the specified week
        DateTime firstDayOfMonth = new DateTime(year, month, 1);
        DateTime targetDate = firstDayOfMonth.AddDays((weekNumber - 1) * 7);

        return targetDate;
    }
    private void SaveAttendanceRecords()
    {
        if (records != null && records.Count > 0)
        {
            string filePath = "attendance.json";
            string json = JsonConvert.SerializeObject(records, Formatting.Indented);
            File.WriteAllText(filePath, json);
        }
    }
    private void LoadAttendanceRecords()
    {
        string filePath = "attendance.json";
        if (File.Exists(filePath))
        {
            try
            {
                string json = File.ReadAllText(filePath);
                records = JsonConvert.DeserializeObject<List<AttendanceRecord>>(json);
            }
            catch (Exception ex)
            {
                // Handle potential exceptions during JSON deserialization (e.g., invalid format)
                MessageBox.Show("Error loading attendance records: " + ex.Message);
                records = new List<AttendanceRecord>(); // Initialize empty list on error
            }
        }
        else
        {
            records = new List<AttendanceRecord>(); // Initialize empty list if file doesn't exist
        }
    }
    public bool ConvertHssfToXssf(string inputFileName, string outputFileName)
    {
        try
        {
            // 檢查輸出檔案名稱是否包含 ".xlsx" 擴展名，如果沒有，則添加
            if (!outputFileName.EndsWith(".xlsx"))
            {
                outputFileName += ".xlsx";
            }

            // 讀取 .xls 文件
            HSSFWorkbook hssfWorkbook;
            using (FileStream file = new FileStream(inputFileName, FileMode.Open, FileAccess.Read))
            {
                hssfWorkbook = new HSSFWorkbook(file);
            }

            // 如果輸出文件已存在，刪除它
            if (File.Exists(outputFileName))
            {
                File.Delete(outputFileName);
            }

            // 創建一個新的 .xlsx 文件
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook();

            // 遍歷所有的工作表
            for (int i = 0; i < hssfWorkbook.NumberOfSheets; i++)
            {
                ISheet hssfSheet = hssfWorkbook.GetSheetAt(i);
                ISheet xssfSheet = xssfWorkbook.CreateSheet(hssfSheet.SheetName);

                // 遍歷所有的行
                for (int j = 0; j <= hssfSheet.LastRowNum; j++)
                {
                    IRow hssfRow = hssfSheet.GetRow(j);
                    IRow xssfRow = xssfSheet.CreateRow(j);

                    if (hssfRow != null)
                    {
                        // 遍歷所有的單元格
                        for (int k = hssfRow.FirstCellNum; k < hssfRow.LastCellNum; k++)
                        {
                            ICell hssfCell = hssfRow.GetCell(k);
                            ICell xssfCell = xssfRow.CreateCell(k);

                            if (hssfCell != null)
                            {
                                // 複製單元格的值
                                switch (hssfCell.CellType)
                                {
                                    case CellType.String:
                                        xssfCell.SetCellValue(hssfCell.StringCellValue);
                                        break;
                                    case CellType.Numeric:
                                        xssfCell.SetCellValue(hssfCell.NumericCellValue);
                                        break;
                                    case CellType.Boolean:
                                        xssfCell.SetCellValue(hssfCell.BooleanCellValue);
                                        break;
                                    case CellType.Formula:
                                        xssfCell.SetCellFormula(hssfCell.CellFormula);
                                        break;
                                    default:
                                        xssfCell.SetCellValue(hssfCell.ToString());
                                        break;
                                }
                            }
                        }
                    }
                }
            }

            // 儲存 .xlsx 文件
            using (FileStream file = new FileStream(outputFileName, FileMode.Create, FileAccess.Write))
            {
                xssfWorkbook.Write(file);
            }
            return true;
        }
        catch (FileNotFoundException e)
        {
            MessageBox.Show("檔案不存在: " + e.Message);
        }
        catch (IOException e)
        {
            MessageBox.Show("檔案無法開啟: " + e.Message);
        }
        catch (Exception e)
        {
            MessageBox.Show("發生錯誤: " + e.Message);
        }
        return false;
    }
    private static XSSFColor HexToXSSFColor(string hexColor)
    {
        // Remove the hash at the front if it's there
        if (hexColor.StartsWith("#"))
        {
            hexColor = hexColor.Substring(1);
        }

        // Convert hex string to Color
        Color color = ColorTranslator.FromHtml("#" + hexColor);

        // Convert Color to byte array
        byte[] rgb = new byte[] { color.R, color.G, color.B };

        // Create XSSFColor from the byte array
        return new XSSFColor(rgb);
    }
    private string ReadFirstCell(string filePath)
    {
        try
        {
            IWorkbook workbook;

            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(file);
            }

            ISheet sheet = workbook.GetSheetAt(0); // 獲取第一個工作表
            IRow row = sheet.GetRow(0); // 獲取第一列
            ICell cell = row.GetCell(0); // 獲取第一欄

            return cell.ToString(); // 讀取單元格的內容
        }
        catch (Exception ex)
        {
            Console.WriteLine("讀取檔案時發生錯誤: " + ex.Message);
            return null;
        }
    }

    private string GetMonthString(ISheet sheet, string input)
    {
        int columnIndex = int.Parse(input.Split('-')[0]);

        IRow row = sheet.GetRow(0);
        ICell cell = row.GetCell(columnIndex);

        return cell.StringCellValue;
    }
    private string GetWeekString(ISheet sheet, string input)
    {
        int columnIndex1 = int.Parse(input.Split('-')[0]);
        int columnIndex2 = int.Parse(input.Split('-')[1]);

        IRow row0 = sheet.GetRow(0);
        IRow row1 = sheet.GetRow(1);
        ICell cell0 = row0?.GetCell(columnIndex1);
        ICell cell1 = row1?.GetCell(columnIndex1);

        string cellValue0 = cell0?.StringCellValue ?? string.Empty;
        string cellValue1 = cell1?.StringCellValue ?? string.Empty;

        // 如果Row(0)內容是空的，往前面的column搜尋以取得該column的Row(0)內容
        if (string.IsNullOrEmpty(cellValue0))
        {
            for (int i = columnIndex1 - 1; i >= 0; i--)
            {
                cellValue0 = row0.GetCell(i)?.StringCellValue ?? string.Empty;
                if (!string.IsNullOrEmpty(cellValue0))
                {
                    break;
                }
            }
        }

        // 將Row(0) 與 Row(1) 組成新的字串
        string newCellValue = cellValue0 + cellValue1;

        if (columnIndex1 != columnIndex2)
        {
            ICell cell2 = row0?.GetCell(columnIndex2);
            ICell cell3 = row1?.GetCell(columnIndex2);

            string cellValue2 = cell2?.StringCellValue ?? string.Empty;
            string cellValue3 = cell3?.StringCellValue ?? string.Empty;

            // 如果Row(0)內容是空的，往前面的column搜尋以取得該column的Row(0)內容
            if (string.IsNullOrEmpty(cellValue2))
            {
                for (int i = columnIndex2 - 1; i >= 0; i--)
                {
                    cellValue2 = row0.GetCell(i)?.StringCellValue ?? string.Empty;
                    if (!string.IsNullOrEmpty(cellValue2))
                    {
                        break;
                    }
                }
            }
            newCellValue = newCellValue + "~" + cellValue2 + cellValue3;
        }
        return ShortenDateRange(newCellValue);
        // return newCellValue;
    }
    private void WriteAvergeResultToSheet(Dictionary<string, double> result, XSSFWorkbook workbook, string sheetName, int startRow, int startColumn)
    {
        ISheet sheet = workbook.GetSheet(sheetName);

        // 如果 column 輸入為 0，則使用輸入 row 的最後一個有資料的 column 的下一個 column 作為指定 column 的值
        if (startColumn == 0)
        {
            IRow row = sheet.GetRow(startRow);
            startColumn = row.LastCellNum + 1;
        }

        var districtSums = new Dictionary<string, double>();
        var districtIdentitySums = new Dictionary<string, double>();

        // 建立顏色
        var ClrStat1 = HexToXSSFColor(lbClrStat1.Text);
        var ClrStat2 = HexToXSSFColor(lbClrStat2.Text);
       

        bool isColor1 = true;

        foreach (var key in result.Keys)
        {
            var parts = key.Split('|');
            var district = parts[0];
            var identity = parts[1];

            // 將 "年長", "中壯", "青壯", "青職" 的數值合併為一個數值，並統一他們的第二資訊為 "青職以上"
            if (identity == "年長" || identity == "中壯" || identity == "青壯" || identity == "青職")
            {
                identity = "青職以上";
                if (!districtIdentitySums.ContainsKey(district + "|" + identity))
                {
                    districtIdentitySums[district + "|" + identity] = 0;
                }
                districtIdentitySums[district + "|" + identity] += result[key];
            }
            else
            {
                // 將其他身份的數值加到 districtIdentitySums 中
                if (!districtIdentitySums.ContainsKey(district + "|" + identity))
                {
                    districtIdentitySums[district + "|" + identity] = 0;
                }

                districtIdentitySums[district + "|" + identity] += result[key];
            }

            // 將所有身份的數值加到 districtSums 中
            if (!districtSums.ContainsKey(district))
            {
                districtSums[district] = 0;
            }
            if (identity != "學齡前")
                districtSums[district] += result[key];
        }

        // 寫入每個第一資訊的 "青職以上" 和總和
        IRow row0 = sheet.GetRow(startRow);
        if (row0 == null)
        {
            row0 = sheet.CreateRow(startRow);
        }
        row0.CreateCell(startColumn).SetCellValue("所屬區");
        row0.CreateCell(startColumn + 1).SetCellValue("身分");
        row0.CreateCell(startColumn + 2).SetCellValue("人數");
        startRow += 1;
        foreach (var district in districtSums.Keys)
        {
            var identities = new List<string> { "青職以上", "大專", "中學", "小學", "學齡前", "總計" };
            foreach (var identity in identities)
            {
                IRow row = sheet.GetRow(startRow++);
                if (row == null)
                {
                    row = sheet.CreateRow(startRow - 1);
                }
                row.CreateCell(startColumn).SetCellValue(district);
                row.CreateCell(startColumn + 1).SetCellValue(identity);
                if (identity == "總計")
                {
                    row.CreateCell(startColumn + 2).SetCellValue(districtSums[district]);
                }
                else
                {
                    row.CreateCell(startColumn + 2).SetCellValue(districtIdentitySums.ContainsKey(district + "|" + identity) ? districtIdentitySums[district + "|" + identity] : 0);
                }

                // 設置儲存格的背景顏色和對齊方式
                XSSFCellStyle style = (XSSFCellStyle)workbook.CreateCellStyle();
                // style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                style.VerticalAlignment = VerticalAlignment.Center;
                if (isColor1 == true)
                {
                    style.SetFillForegroundColor(ClrStat2);
                }
                else
                {
                    style.SetFillForegroundColor(ClrStat1);
                }

                style.FillPattern = FillPattern.SolidForeground;
                for (int i = startColumn; i <= startColumn + 2; i++)
                {
                    row.GetCell(i).CellStyle = style;
                }
            }
            XSSFCellStyle style0 = (XSSFCellStyle)workbook.CreateCellStyle();
            style0.VerticalAlignment = VerticalAlignment.Center;
            style0.SetFillForegroundColor(ClrStat1);
            style0.FillPattern = FillPattern.SolidForeground;
            for (int i = startColumn; i <= startColumn + 2; i++)
            {
                row0.GetCell(i).CellStyle = style0;
            }

            // 切換顏色
            isColor1 = !isColor1;
        }

        MergeColumnCells(sheet, startColumn);
        sheet.SetColumnWidth(startColumn, 14 * 256);
        sheet.SetColumnWidth(startColumn + 1, 10 * 256);
        sheet.SetColumnWidth(startColumn + 2, 4 * 256);
    }

    private void RemoveZeroColumns(ISheet sheet, int startRow)
    {
        int lastColumnNum = sheet.GetRow(startRow).LastCellNum;

        for (int columnNum = lastColumnNum - 1; columnNum >= 0; columnNum--)
        {
            bool allZeros = true;

            for (int rowNum = startRow; rowNum <= sheet.LastRowNum; rowNum++)
            {
                IRow row = sheet.GetRow(rowNum);
                if (row != null)
                {
                    ICell cell = row.GetCell(columnNum);
                    if (cell != null && cell.NumericCellValue != 0)
                    {
                        allZeros = false;
                        break;
                    }
                }
            }

            if (allZeros)
            {
                RemoveColumn(sheet, columnNum);
            }
            else
                return;
        }
    }

    private void RemoveColumn(ISheet sheet, int columnIndex)
    {
        for (int r = 0; r <= sheet.LastRowNum; r++)
        {
            IRow row = sheet.GetRow(r);
            if (row != null)
            {
                if (columnIndex < row.LastCellNum) // 檢查要移除的列是否存在
                {
                    for (int c = columnIndex; c < row.LastCellNum - 1; c++)
                    {
                        ICell cell = row.GetCell(c);
                        if (cell != null)
                            row.RemoveCell(cell);

                        ICell nextCell = row.GetCell(c + 1);
                        if (nextCell != null)
                        {
                            ICell newCell = row.CreateCell(c, nextCell.CellType);
                            newCell.CellStyle = nextCell.CellStyle;
                            if (nextCell.CellType == CellType.Formula)
                            {
                                newCell.SetCellFormula(nextCell.CellFormula);
                            }
                            else
                            {
                                newCell.SetCellValue(nextCell.ToString());
                            }
                        }
                    }

                    ICell lastCell = row.GetCell(row.LastCellNum - 1);
                    if (lastCell != null)
                        row.RemoveCell(lastCell);
                }
            }
        }
    }

    public static void SetTabText(DataGridView dataGridView, string s)
    {
        // Check if the DataGridView is contained within a TabPage
        TabPage tabPage = FindParentTabPage(dataGridView);

        // If a TabPage is found, set its Text property
        if (tabPage != null)
        {
            tabPage.Text = s;
        }
    }

    private static TabPage FindParentTabPage(System.Windows.Forms.Control control)
    {
        // Loop through the parent hierarchy of the control to find a TabPage
        System.Windows.Forms.Control parent = control.Parent;
        while (parent != null)
        {
            if (parent is TabPage)
            {
                return (TabPage)parent;
            }
            parent = parent.Parent;
        }
        // Return null if no TabPage is found
        return null;
    }
    private void OpenExcelFile(string inputFilePath, string outputFilePath, DataGridView dbvResult)
    {
        try
        {
            string startColumnLetter = txtBoxStartColumn.Text; // 使用者輸入的開始列名
            int startColumnIndex = startColumnLetter.ToUpper()[0] - 'A'; // 轉換列名為索引
            if (false == ConvertHssfToXssf(inputFilePath, outputFilePath)) return;

            var sheetName = new List<string>();
            using (var fs = new FileStream(outputFilePath, FileMode.Open, FileAccess.Read))
            {
                var workbook = new XSSFWorkbook(fs);
                var sheet = workbook.GetSheetAt(0); // 選擇第一個工作表
                var row = sheet.GetRow(1); // 選擇第五行
                int lastColumnWithData = startColumnIndex;
                if (cbIgnoreNoData.Checked)
                {
                    RemoveZeroColumns(sheet, 2);
                }
                lastColumnWithData = GetLastColumnWithData(sheet, 1, startColumnIndex);
                if (rbWeek.Checked == true)
                {
                    lastColumnWithData = GetLastColumnWithData(sheet, 1, startColumnIndex); // 分析第二列
                    List<string> result = GroupNumbers(startColumnIndex, lastColumnWithData, 1, ckbFwdBwd.Checked);
                    foreach (string s in result)
                    {
                        // MessageBox.Show(s);
                        var dict = ClassifyAttendancy(s, sheet, CatagoryArray(inputFilePath));
                        var week = GetWeekString(sheet, s);
                        sheetName.Add(week);
                        FillSheetWithDict(dict, week, workbook, false);
                        var byIdentity = AttendanceCountByIdentity(s, sheet);
                        var calculateAverageResult = CalculateAverage(byIdentity);
                        WriteAvergeResultToSheet(calculateAverageResult, workbook, week, 0, 0);
                    }

                }
                if (rbMonth.Checked == true)
                {
                    List<string> result = GroupByMonth(sheet);
                    foreach (string s in result)

                    {
                        // MessageBox.Show(s);
                        var dict = ClassifyAttendancy(s, sheet, CatagoryArray(inputFilePath));
                        var month = GetMonthString(sheet, s);
                        sheetName.Add(month);
                        FillSheetWithDict(dict, month, workbook, false);
                        var byIdentity = AttendanceCountByIdentity(s, sheet);
                        var calculateAverageResult = CalculateAverage(byIdentity);
                        WriteAvergeResultToSheet(calculateAverageResult, workbook, month, 0, 0);
                    }
                }

                if (rbHalfYear.Checked == true)
                {
                    lastColumnWithData = GetLastColumnWithData(sheet, 1, startColumnIndex); // 分析第二列
                    List<string> result = GroupNumbers(startColumnIndex, lastColumnWithData, 26, ckbFwdBwd.Checked);

                    foreach (string s in result)
                    {
                        var dict = ClassifyAttendancy(s, sheet, CatagoryArray(inputFilePath));
                        var week = GetWeekString(sheet, s);
                        sheetName.Add(week);
                        FillSheetWithDict(dict, week, workbook, false);
                        var byIdentity = AttendanceCountByIdentity(s, sheet);
                        var calculateAverageResult = CalculateAverage(byIdentity);
                        WriteAvergeResultToSheet(calculateAverageResult, workbook, week, 0, 0);
                    }
                }
                if (rbSelfDef.Checked == true)
                {
                    lastColumnWithData = GetLastColumnWithData(sheet, 1, startColumnIndex); // 分析第二列
                    List<string> result = GroupNumbers(startColumnIndex, lastColumnWithData, int.Parse(tbSelfDefWeek.Text), ckbFwdBwd.Checked);

                    foreach (string s in result)
                    {
                        var dict = ClassifyAttendancy(s, sheet, CatagoryArray(inputFilePath));
                        var week = GetWeekString(sheet, s);
                        sheetName.Add(week);
                        FillSheetWithDict(dict, week, workbook, false);
                        var byIdentity = AttendanceCountByIdentity(s, sheet);
                        var calculateAverageResult = CalculateAverage(byIdentity);
                        WriteAvergeResultToSheet(calculateAverageResult, workbook, week, 0, 0);
                    }
                }
                for (int i = 1; i < sheetName.Count; i++)
                {
                    // 取得當前表單和前一個表單的名稱
                    string currentSheetName = sheetName[i];
                    string previousSheetName = sheetName[i - 1];

                    CompareSheets(workbook, currentSheetName, previousSheetName);
                }

                var sheetToShow = workbook.GetSheet(sheetName[sheetName.Count - 1]);
                DisplayExcelInDataGridView(sheetToShow, dbvResult);

                IRow row0 = sheet.GetRow(0); // 取得第一列
                ICell cell = row0.GetCell(0); // 取得第一欄
                for (int i = 0; i < sheetName.Count; i++)
                {
                    ISheet minor_sheet = workbook.GetSheet(sheetName[i]);
                    FillSheetNameAndDataName(minor_sheet, sheetName[i] + " " + cell.ToString());
                }

                SetTabText(dbvResult, sheetToShow.SheetName + " " + cell.ToString());
                using (FileStream file = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }


            }
        }
        catch (FileNotFoundException e)
        {
            MessageBox.Show("檔案不存在: " + e.Message);
        }
        catch (IOException e)
        {
            MessageBox.Show("檔案無法開啟: " + e.Message);
        }
        catch (Exception e)
        {
            MessageBox.Show("發生錯誤: " + e.Message);
        }
    }

    private void DisplayExcelInDataGridView(ISheet sheet, DataGridView dataGridView)
    {
        var dt = new DataTable();
        var headerRow = sheet.GetRow(0);
        int cellCount = headerRow.LastCellNum;

        // Add columns with numeric names
        for (int i = 0; i < cellCount; i++)
        {
            DataColumn column = new DataColumn(i.ToString());
            dt.Columns.Add(column);
        }

        // Add rows
        for (int i = 0; i <= sheet.LastRowNum; i++)
        {
            var row = sheet.GetRow(i);
            DataRow dataRow = dt.NewRow();
            for (int j = 0; j < cellCount; j++)
            {
                if (row.GetCell(j) != null)
                {
                    dataRow[j] = RemoveParentheses(row.GetCell(j).ToString());
                }
            }
            dt.Rows.Add(dataRow);
        }

        // Set the font size according to tbFontSize.Text
        float fontSize;
        if (float.TryParse(tbFontSize.Text, out fontSize))
        {
            dataGridView.DefaultCellStyle.Font = new System.Drawing.Font(dataGridView.DefaultCellStyle.Font.FontFamily, fontSize);
        }
        else
        {
            // Handle invalid font size
            MessageBox.Show("Invalid font size.");
        }

        dataGridView.DataSource = dt;
        dataGridView.ColumnHeadersVisible = false;
        dataGridView.RowHeadersVisible = false;
        // dataGridView.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

        // Set the column captions to the values from the second row
        if (dt.Rows.Count > 0)
        {
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                dataGridView.Columns[i].HeaderText = dt.Rows[i].ToString();
            }
        }

        // Set cell colors and handle merged cells
        for (int i = 0; i <= sheet.LastRowNum; i++)
        {
            var row = sheet.GetRow(i);
            for (int j = 0; j < cellCount; j++)
            {
                if (row.GetCell(j) != null)
                {
                    // Get the color of the cell in Excel
                    var color = row.GetCell(j).CellStyle.FillForegroundColorColor;
                    if (color != null)
                    {
                        // Convert the color to System.Drawing.Color
                        var drawingColor = System.Drawing.Color.FromArgb(color.RGB[0], color.RGB[1], color.RGB[2]);

                        // Set the background color of the corresponding cell in DataGridView
                        dataGridView.Rows[i].Cells[j].Style.BackColor = drawingColor;
                    }

                    // Handle merged cells
                    foreach (var range in sheet.MergedRegions)
                    {
                        if (range.IsInRange(i, j))
                        {
                            dataGridView.Rows[i].Cells[j].Value = RemoveParentheses(sheet.GetRow(range.FirstRow).GetCell(range.FirstColumn).ToString());
                            break;
                        }
                    }
                }
            }
        }

        SortDataGridViewByDictionary(dataGridView, attendanceSummary, 2);
        if (cbDisplayLastAttend.Checked ==  true)
            AppendLastAttendDate(sheet, dataGridView, attendanceSummary, 2);
        AdjustDataGridViewByColor(dataGridView, 2);
        dataGridView.AutoResizeColumns();
        dataGridView.AutoResizeRows();
    }

    public void AppendLastAttendDate(ISheet sheet, DataGridView dgv, Dictionary<string, Dictionary<string, object>> attendanceSummary, int startRow)
    {
        // 取得 DataGridView 的字體大小
        float originalFontSize = dgv.DefaultCellStyle.Font.Size;

        // 計算新的字體大小
        float newFontSize = originalFontSize * 0.75f;

        // 建立新的字體物件
        Font newFont = new Font(dgv.DefaultCellStyle.Font.FontFamily, newFontSize);


        // 遍歷 DataGridView 的每個儲存格
        foreach (DataGridViewRow row in dgv.Rows)
        {
            foreach (DataGridViewCell cell in row.Cells)
            {
                if (cell.Value != null && attendanceSummary.TryGetValue(cell.Value.ToString(), out var dictValue))
                {
                    // 從巢狀字典中取得 LastAttendDate
                    string lastAttendDate = dictValue["LastAttendDate"].ToString();

                    // 加入判斷：如果 LastAttendDate 被包含在 TabPageText 內，則不附加日期
                    if (!sheet.SheetName.Contains(lastAttendDate) || cbAllDisplayLastAttend.Checked == true)
                    {
                        // 轉換日期格式
                        lastAttendDate = ConvertDateFormat(lastAttendDate);

                        // 將 LastAttendDate 附加到儲存格的當前值
                        cell.Value = $"{cell.Value} {lastAttendDate}";

                        // 設定儲存格的字體
                        cell.Style.Font = newFont;
                    }
                }
            }
        }
    }
    static string ConvertDateFormat(string input)
    {
        // 使用正則表達式解析年份、月份和週次（週次為中文數字）
        var match = Regex.Match(input, @"(\d{4})年(\d{1,2})月第([一二三四五六七八九十]+)週");
        if (match.Success)
        {
            int year = int.Parse(match.Groups[1].Value);
            int month = int.Parse(match.Groups[2].Value);
            string weekChinese = match.Groups[3].Value;

            // 將中文數字轉換成阿拉伯數字
            int week = ConvertChineseNumberToInteger(weekChinese);

            // 取得年份的最後兩位
            string yearShort = (year % 100).ToString("D2");
            // 取得月份的兩位數格式
            string monthShort = month.ToString("D2");
            // 將週數以 'W' 為前綴
            string weekFormat = $"W{week}";

            return $"{yearShort}.{monthShort}.{weekFormat}";
        }
        else
        {
            return input;
        }
    }

    // 中文數字轉換為阿拉伯數字的方法
    static int ConvertChineseNumberToInteger(string chineseNumber)
    {
        switch (chineseNumber)
        {
            case "一": return 1;
            case "二": return 2;
            case "三": return 3;
            case "四": return 4;
            case "五": return 5;
            case "六": return 6;
            case "七": return 7;
            case "八": return 8;
            case "九": return 9;
            case "十": return 10;
            default: return -1; // 錯誤情況
        }
    }
    private string RemoveParentheses(string input)
    {
        // 使用正則表達式移除全形和半形括號及其中的內容
        return System.Text.RegularExpressions.Regex.Replace(input, @"[（(][^）)]*[）)]", string.Empty);
    }

    public void SortDataGridViewByDictionary(DataGridView dgv, Dictionary<string, Dictionary<string, object>> attendanceSummary, int startRow)
    {
        // 迭代每個欄位
        foreach (DataGridViewColumn column in dgv.Columns)
        {
            // 提取儲存格值及其對應的字典值
            var cellValuesWithDictValues = new List<Tuple<string, int, DataGridViewCellStyle>>();
            for (int rowIndex = startRow; rowIndex <= dgv.Rows.Count - 1; rowIndex++) // 包含最後一行
            {
                var cell = dgv.Rows[rowIndex].Cells[column.Index];
                var cellValue = cell.Value != null ? cell.Value.ToString() : null;
                var cellStyle = cell.Style.Clone(); // 複製儲存格樣式

                if (cellValue != null && attendanceSummary.TryGetValue(cellValue, out var dictValue))
                {
                    // 從巢狀字典中取得 TotalAttendance
                    int attendanceValue = (int)dictValue["TotalAttendance"];
                    cellValuesWithDictValues.Add(Tuple.Create(cellValue, attendanceValue, cellStyle));
                }
                else
                {
                    // 處理儲存格值為 null 或在字典中找不到的情況
                    cellValuesWithDictValues.Add(Tuple.Create(cellValue, int.MinValue, cellStyle));
                }
            }

            // 根據字典值以降序排序
            var sortedValues = cellValuesWithDictValues.OrderByDescending(t => t.Item2).ToList();

            // 將排序後的數值應用回 DataGridView
            for (int rowIndex = startRow; rowIndex <= dgv.Rows.Count - 1; rowIndex++) // 包含最後一行
            {
                dgv.Rows[rowIndex].Cells[column.Index].Value = sortedValues[rowIndex - startRow].Item1;
                dgv.Rows[rowIndex].Cells[column.Index].Style = sortedValues[rowIndex - startRow].Item3; // 設定儲存格樣式
            }
        }
    }

    public void AdjustDataGridViewByColor(DataGridView dgv, int startRow)
    {
        for (int col = 0; col < dgv.Columns.Count; col++)
        {
            List<object> coloredCells = new List<object>();
            List<System.Drawing.Color> coloredCellColors = new List<System.Drawing.Color>();
            List<object> nonColoredCells = new List<object>();

            // 先備份所有的格子
            for (int row = startRow; row < dgv.Rows.Count; row++)
            {
                DataGridViewCell cell = dgv.Rows[row].Cells[col];
                if (cell.Style.BackColor != System.Drawing.Color.Empty)
                {
                    coloredCells.Add(cell.Value);
                    coloredCellColors.Add(cell.Style.BackColor);
                }
                else
                {
                    nonColoredCells.Add(cell.Value);
                }
            }

            // 把有顏色的格子先填上去
            int insertIndex = startRow;
            for (int i = 0; i < coloredCells.Count; i++)
            {
                dgv.Rows[insertIndex].Cells[col].Value = coloredCells[i];
                dgv.Rows[insertIndex].Cells[col].Style.BackColor = coloredCellColors[i];
                insertIndex++;
            }

            // 把沒顏色的格子填在後面
            for (int i = 0; i < nonColoredCells.Count; i++)
            {
                dgv.Rows[insertIndex].Cells[col].Value = nonColoredCells[i];
                dgv.Rows[insertIndex].Cells[col].Style.BackColor = System.Drawing.Color.Empty;
                insertIndex++;
            }
        }
    }



    private List<string> GroupByMonth(ISheet sheet)
    {
        var result = new List<string>();
        int groupCount = 0;
        int i = 0;

        while ((sheet.GetRow(1).GetCell(i) != null) && (sheet.GetRow(1).GetCell(i).ToString() != ""))
        {
            if (sheet.GetRow(1).GetCell(i).ToString() == "第一週")
            {
                int groupStart = i;
                int j = 0;
                while (sheet.GetRow(1).GetCell(i + j + 1) != null && (sheet.GetRow(1).GetCell(i + j).ToString() == "第一週" || sheet.GetRow(1).GetCell(i + j).ToString() == "第二週" ||
                    sheet.GetRow(1).GetCell(i + j).ToString() == "第三週" || sheet.GetRow(1).GetCell(i + j).ToString() == "第四週" || sheet.GetRow(1).GetCell(i + j).ToString() == "第五週"))
                {
                    if (sheet.GetRow(1).GetCell(i + j + 1) == null || sheet.GetRow(1).GetCell(i + j + 1).ToString() == "" || sheet.GetRow(1).GetCell(i + j + 1).ToString() == "第一週")
                    {
                        break;
                    }
                    j++;
                }
                int groupEnd = i + j;
                result.Add(groupStart + "-" + groupEnd);
                groupCount++;
            }
            i++;
        }
        return result;
    }


    private int GetLastColumnWithData(ISheet sheet, int rowIndex, int startColumnIndex)
    {
        var row = sheet.GetRow(rowIndex);

        int lastColumnWithData = startColumnIndex;

        for (int i = startColumnIndex; i <= row.LastCellNum; i++)
        {
            if (row.GetCell(i) != null && !string.IsNullOrEmpty(row.GetCell(i).ToString()))
            {
                lastColumnWithData = i;
            }
        }

        return lastColumnWithData;
    }

    private List<string> GroupNumbers(long startNum, long endNum, long groupSize, bool startFromEnd)
    {
        var groups = new List<string>();
        long total = endNum - startNum + 1;
        long numGroups = (long)Math.Ceiling((double)total / groupSize);

        if (startFromEnd)
        {
            for (long i = 0; i < numGroups; i++)
            {
                long groupStartNum = endNum - groupSize + 1;
                if (groupStartNum < startNum) groupStartNum = startNum;
                groups.Insert(0, groupStartNum + "-" + endNum);
                endNum = groupStartNum - 1;
            }
        }
        else
        {
            for (long i = numGroups - 1; i >= 0; i--)
            {
                long groupEndNum = startNum + groupSize - 1;
                if (groupEndNum > endNum) groupEndNum = endNum;
                groups.Add(startNum + "-" + groupEndNum);
                startNum = groupEndNum + 1;
            }
        }

        return groups;
    }

    private Dictionary<string, List<string>> ClassifyAttendancy(string columnRange, ISheet sheet, string[] categories)
    {
        var dict = new Dictionary<string, List<string>>();

        // 使用 Split 函數來分割字串
        string[] parts = columnRange.Split('-');
        int startColumn = int.Parse(parts[0]);
        int endColumn = int.Parse(parts[1]);
        IRow row0 = sheet.GetRow(0); // 獲取第一列
        ICell cell0 = row0.GetCell(0); // 獲取第一欄
        var type = cell0.ToString(); // 讀取單元格的內容
        for (int i = 2; i <= sheet.LastRowNum; i++)
        {
            var row = sheet.GetRow(i);
            string groupName = row.GetCell(0)?.ToString() + row.GetCell(1)?.ToString();

            if (string.IsNullOrEmpty(groupName))
                continue;

            foreach (var _category in categories)
            {
                string _key = groupName + "_" + _category;
                if (!dict.ContainsKey(_key))
                    dict[_key] = new List<string>();
            }

            string name = row.GetCell(3)?.ToString();
            int attendanceCount = 0;
            int weekCount = 0;

            for (int j = startColumn; j <= endColumn; j++)
            {
                ICell cell = row.GetCell(j);
                int attendancy = 0;
                if (cell != null)
                {
                    attendancy = 0;
                    switch (cell.CellType)
                    {
                        case CellType.Numeric:

                            if (cell.NumericCellValue == 1)
                            {
                                attendanceCount++;
                                attendancy = 1;
                            }

                            break;
                        case CellType.String:
                            if (cell.StringCellValue == "1")
                            {
                                attendanceCount++;
                                attendancy = 1;
                            }
                            break;
                            // 你可以根據需要添加更多的 case
                    }
                    AttendanceRecord newRecord = new AttendanceRecord
                    {
                        Name = name,
                        Date = GetWeekString(sheet, string.Format("{0}-{1}", j, j)),
                        Attendance = attendancy
                    };

                    // Call the AddNewRecord function
                    AddNewRecord(newRecord);
                }
                weekCount++;
            }
            double attendanceRate = (double)attendanceCount / weekCount;
            string category;

            if (categories.Length == 2)
                category = attendanceRate > double.Parse(txtBoxStable.Text) / 100 ? categories[0] : categories[1];
            else
                category = attendanceRate > double.Parse(txtBoxStable.Text) / 100 ? categories[0] : attendanceRate > 0 ? categories[1] : categories[2];


            string key = groupName + "_" + category;

            if (!dict.ContainsKey(key))
                dict[key] = new List<string>();

            dict[key].Add(name + "_" + attendanceRate.ToString("P"));
        }
        CalculateAttendanceSummary();
        return dict;
    }

    private Dictionary<string, List<int>> AttendanceCountByIdentity(string columnRange, ISheet sheet)
    {
        var dict = new Dictionary<string, List<int>>();
        var dict2 = new Dictionary<string, List<int>>();

        // 使用 Split 函數來分割字串
        string[] parts = columnRange.Split('-');
        int startColumn = int.Parse(parts[0]);
        int endColumn = int.Parse(parts[1]);

        for (int i = 2; i <= sheet.LastRowNum; i++)
        {
            var row = sheet.GetRow(i);
            string district = row.GetCell(0)?.ToString() + row.GetCell(1)?.ToString();
            string upper_district = row.GetCell(0)?.ToString();
            if (string.IsNullOrEmpty(district))
                continue;

            string identity = row.GetCell(5)?.ToString();

            for (int j = startColumn; j <= endColumn; j++)
            {
                ICell cell = row.GetCell(j);
                int weekAttendance = 0;

                if (cell != null)
                {
                    switch (cell.CellType)
                    {
                        case CellType.Numeric:
                            weekAttendance = (int)cell.NumericCellValue;
                            break;
                        case CellType.String:
                            if (int.TryParse(cell.StringCellValue, out int numericValue))
                            {
                                weekAttendance = numericValue;
                            }
                            break;
                            // 你可以根據需要添加更多的 case
                    }
                }

                if (weekAttendance == 1)
                {
                    string key = district + "|" + identity + "|" + (j - 7); // 小區|身分|週次
                    string key2 = upper_district + "|" + identity + "|" + (j - 7);
                    if (dict.ContainsKey(key))
                    {
                        dict[key][0] += 1;
                    }
                    else
                    {
                        dict[key] = new List<int> { 1 };
                    }
                    if (dict2.ContainsKey(key2))
                    {
                        dict2[key2][0] += 1;
                    }
                    else
                    {
                        dict2[key2] = new List<int> { 1 };
                    }
                }
            }

        }

        // 把 dict2 的內容添加到 dict 的最後面
        foreach (var item in dict2)
        {
            dict[item.Key] = item.Value;
        }

        return dict;
    }

    private Dictionary<string, double> CalculateAverage(Dictionary<string, List<int>> dict)
    {
        var districtIdentityDict = new Dictionary<string, List<int>>();

        // 分類每個小區和身份的數字
        // 身份的固定順序
        string[] identities = new string[] { "年長", "中壯", "青壯", "青職", "大專", "中學", "小學", "學齡前" };

        foreach (var key in dict.Keys)
        {
            var parts = key.Split('|');
            var district = parts[0];
            var identity = parts[1];
            var districtIdentityKey = district + "|" + identity;

            // 如果 districtIdentityKey 不存在，則為每個 identity 創建一個新的 key
            if (!districtIdentityDict.ContainsKey(districtIdentityKey))
            {
                foreach (var id in identities)
                {
                    var newKey = district + "|" + id;
                    if (!districtIdentityDict.ContainsKey(newKey))
                    {
                        districtIdentityDict[newKey] = new List<int>();
                    }
                }
            }

            districtIdentityDict[districtIdentityKey].AddRange(dict[key]);
        }


        var result = new Dictionary<string, double>();

        // 計算每個小區和身份的中位數平均數
        foreach (var key in districtIdentityDict.Keys)
        {
            var values = districtIdentityDict[key];

            // 如果 values 為空，則將 result[key] 設為 0 並跳過此次循環
            if (values.Count == 0)
            {
                result[key] = 0;
                continue;
            }

            // 使用 OrderBy 對 List 進行排序
            values.Sort();

            // 計算中位數
            double median = (values.Count % 2 != 0) ? values[values.Count / 2] : (values[(values.Count - 1) / 2] + values[values.Count / 2]) / 2.0;

            // 計算低於中位數XX %的數字的平均數
            double lowerBound = median * double.Parse(txtbIgnoreLevel.Text) / 100;
            double sum = 0;
            int count = 0;

            foreach (var value in values)
            {
                if (value >= lowerBound)
                {
                    sum += value;
                    count++;
                }
                else
                {
                    count = count;// debug purpose
                }
            }

            // 計算平均數並更新字典
            result[key] = Math.Round(sum / count, 0);
        }


        return result;
    }

    private void FillSheetWithDict(Dictionary<string, List<string>> dict, string sheetName, XSSFWorkbook workbook, bool replaceUnderscore)
    {
        // Remove the sheet with the same name if it exists
        int sheetIndex = workbook.GetSheetIndex(sheetName);
        if (sheetIndex != -1)
        {
            workbook.RemoveSheetAt(sheetIndex);
        }

        ISheet sheet = workbook.CreateSheet(sheetName);

        int columnIndex = 0;
        foreach (var kvp in dict)
        {
            // Split the key into two parts
            string[] parts = kvp.Key.Split('_');
            string info1 = parts[0];
            string info2 = parts.Length > 1 ? parts[1] : "";

            // Create a row and fill the first and second cells with the two parts of the key
            IRow row1 = sheet.GetRow(0) ?? sheet.CreateRow(0);
            IRow row2 = sheet.GetRow(1) ?? sheet.CreateRow(1);
            row1.CreateCell(columnIndex).SetCellValue(info1);
            row2.CreateCell(columnIndex).SetCellValue(info2);

            // Fill the rest of the column with the list values
            for (int i = 0; i < kvp.Value.Count; i++)
            {
                string value = kvp.Value[i];
                if (replaceUnderscore)
                {
                    if (value.Contains("_"))
                    {
                        // Replace "_" with "(" ")"
                        value = value.Replace("_", "(") + ")";
                    }
                }
                else
                {
                    if (value.Contains("_"))
                    {
                        // Remove everything after and including the first "_"
                        value = value.Substring(0, value.IndexOf("_"));
                    }
                }
                IRow row = sheet.GetRow(i + 2) ?? sheet.CreateRow(i + 2);
                row.CreateCell(columnIndex).SetCellValue(value);
            }

            columnIndex++;  // Move to the next column
        }
        int rowIndex = 1;  // Row index for the second row
        IRow row3 = sheet.GetRow(rowIndex);
        if (row3 != null)
        {
            ICellStyle style = workbook.CreateCellStyle();
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            for (int i = row3.FirstCellNum; i < row3.LastCellNum; i++)
            {
                ICell cell = row3.GetCell(i);
                if (cell != null)
                {
                    int length = cell.ToString().Length;
                    sheet.SetColumnWidth(i, ((length + 1) * 2) * 256);  // Set the column width based on the length of the cell content
                    cell.CellStyle = style;
                }
            }
        }
        ColorRow(sheet, 0, lbClrHdr1.Text);
        ColorRow(sheet, 1, lbClrHdr2.Text);
        MergeCells(sheet, 0);
    }
    public void ColorRow(ISheet sheet, int rowNum, string colorCode)
    {
        try
        {
            // 獲取行
            IRow row = sheet.GetRow(rowNum);
            if (row == null)
            {
                throw new ArgumentException($"行號 {rowNum} 無效");
            }

            // 轉換色碼
            System.Drawing.Color drawingColor = ColorTranslator.FromHtml(colorCode);
            byte[] rgb = new byte[3] { drawingColor.R, drawingColor.G, drawingColor.B };
            XSSFColor color = new XSSFColor(rgb);

            // 創建單元格樣式
            ICellStyle cellStyle = sheet.Workbook.CreateCellStyle();
            cellStyle.FillPattern = FillPattern.SolidForeground;
            ((XSSFCellStyle)cellStyle).SetFillForegroundColor(color);

            // 著色有值的單元格
            for (int i = 0; i < row.LastCellNum; i++)
            {
                ICell cell = row.GetCell(i);
                if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
                {
                    cell.CellStyle = cellStyle;
                }
            }
        }
        catch (Exception e)
        {
            MessageBox.Show("發生錯誤: " + e.Message);
        }
    }
    private void FillSheetNameAndDataName(ISheet sheet, string inputString)
    {
        IRow row = sheet.GetRow(0); // 取得第一列

        sheet.ShiftRows(0, sheet.LastRowNum, 1); // 將所有列向下移動一列

        IRow newRow = sheet.CreateRow(0); // 在第一列插入新的列

        int firstCellNum = row.FirstCellNum; // 取得第一列的第一個資料格子的索引
        int lastCellNum = row.LastCellNum; // 取得第一列的最後一個資料格子的索引

        sheet.AddMergedRegion(new CellRangeAddress(0, 0, firstCellNum, lastCellNum - 1)); // 合併新插入列的資料格子

        ICell cell = newRow.CreateCell(firstCellNum); // 在新插入的列中創建一個資料格子
        cell.SetCellValue(inputString); // 將輸入的字串設定為資料格子的值

        ICellStyle style = sheet.Workbook.CreateCellStyle(); // 創建一個新的樣式
        style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; // 設定樣式的對齊方式為水平置中

        cell.CellStyle = style; // 將資料格子的樣式設定為剛創建的樣式
    }


    private void MergeCells(ISheet sheet, int rowNumber)
    {
        IRow row = sheet.GetRow(rowNumber);
        string previousValue = null;
        int startMergeIndex = -1;

        for (int i = 0; i < row.LastCellNum; i++)
        {
            ICell cell = row.GetCell(i);
            if (cell != null)
            {
                if (previousValue == null)
                {
                    previousValue = cell.StringCellValue;
                    startMergeIndex = i;
                }
                else if (previousValue == cell.StringCellValue)
                {
                    continue;
                }
                else
                {
                    if (i - startMergeIndex > 1)
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(rowNumber, rowNumber, startMergeIndex, i - 1));
                        row.GetCell(startMergeIndex).SetCellValue(previousValue);
                    }
                    previousValue = cell.StringCellValue;
                    startMergeIndex = i;
                }
            }
        }

        if (row.LastCellNum - startMergeIndex > 1)
        {
            sheet.AddMergedRegion(new CellRangeAddress(rowNumber, rowNumber, startMergeIndex, row.LastCellNum - 1));
            row.GetCell(startMergeIndex).SetCellValue(previousValue);
        }
    }

    private void MergeColumnCells(ISheet sheet, int columnNumber)
    {
        string previousValue = null;
        int startMergeIndex = -1;

        for (int i = 0; i <= sheet.LastRowNum; i++)
        {
            IRow row = sheet.GetRow(i);
            if (row != null)
            {
                ICell cell = row.GetCell(columnNumber);
                if (cell != null)
                {
                    if (previousValue == null)
                    {
                        previousValue = cell.StringCellValue;
                        startMergeIndex = i;
                    }
                    else if (previousValue == cell.StringCellValue)
                    {
                        continue;
                    }
                    else
                    {
                        if (i - startMergeIndex > 1 && startMergeIndex >= 0)
                        {
                            sheet.AddMergedRegion(new CellRangeAddress(startMergeIndex, i - 1, columnNumber, columnNumber));
                        }
                        previousValue = cell.StringCellValue;
                        startMergeIndex = i;
                    }
                }
                else if (previousValue != null)
                {
                    if (i - startMergeIndex > 1 && startMergeIndex >= 0)
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(startMergeIndex, i - 1, columnNumber, columnNumber));
                    }
                    previousValue = null;
                    startMergeIndex = -1;
                }
            }
        }
    }
    private void CompareSheets(XSSFWorkbook workbook, string mainSheetName, string compareSheetName)
    {
        ISheet mainSheet = workbook.GetSheet(mainSheetName);
        ISheet compareSheet = workbook.GetSheet(compareSheetName);

        // Define the colors
        XSSFColor lightGreen = HexToXSSFColor(lbClrPos.Text);
        XSSFColor lightRed = HexToXSSFColor(lbClrNeg.Text);

        // Loop through columns
        for (int col = 0; col < mainSheet.GetRow(0).LastCellNum; col++)
        {
            // Compare column
            for (int row = 1; row <= mainSheet.LastRowNum; row++)
            {
                ICell mainCell = mainSheet.GetRow(row)?.GetCell(col);
                string mainCellValue = mainCell?.ToString() ?? string.Empty;

                int bracketPos = mainCellValue.IndexOf("(");
                string removeAfterBracket = bracketPos > 0 ? mainCellValue.Substring(0, bracketPos) : mainCellValue;

                if (!string.IsNullOrEmpty(mainCellValue) && !CellExistsInColumn(compareSheet, col, removeAfterBracket))
                {
                    // Check if the cell already has a color
                    var existingColor = mainCell.CellStyle.FillForegroundColorColor;
                    if (existingColor == null || (existingColor.RGB[0] == 255 && existingColor.RGB[1] == 255 && existingColor.RGB[2] == 255))
                    {
                        // Create a new cell style and set the fill foreground color
                        XSSFCellStyle style = (XSSFCellStyle)workbook.CreateCellStyle();

                        // Highlight the cell with light green or light red color
                        if (CellExistsInColumn(compareSheet, col + 1, removeAfterBracket) || CellExistsInColumn(compareSheet, col + 2, removeAfterBracket))
                        {
                            // Green, if attendance getting better
                            style.SetFillForegroundColor(lightGreen); // 使用自訂顏色
                            style.FillPattern = FillPattern.SolidForeground;
                        }
                        else
                        {
                            // Red, if attendance getting worse
                            style.SetFillForegroundColor(lightRed); // 使用自訂顏色
                            style.FillPattern = FillPattern.SolidForeground;
                        }

                        // Apply the style to the cell
                        mainCell.CellStyle = style;
                    }
                }
            }
        }
    }
    private bool CellExistsInColumn(ISheet sheet, int col, string value)
    {
        for (int row = 2; row <= sheet.LastRowNum; row++)
        {
            ICell cell = sheet.GetRow(row).GetCell(col);
            if (cell != null && cell.ToString().Contains(value))
            {
                return true;
            }
        }
        return false;
    }
    private void AttendForm_SizeChanged(object sender, EventArgs e)
    {
        // 計算 Form 大小的變化比例
        float widthRatio = (float)this.Width / originalFormSize.Width;
        float heightRatio = (float)this.Height / originalFormSize.Height;

        // 根據比例調整 TabControl 的大小
        tabControl1.Width = (int)(originalTabControlSize.Width * widthRatio);
        tabControl1.Height = (int)(originalTabControlSize.Height * heightRatio);

        // 根據比例調整 DataGridView 的大小
        foreach (TabPage tabPage in tabControl1.TabPages)
        {
            foreach (System.Windows.Forms.Control control in tabPage.Controls)
            {
                if (control is DataGridView)
                {
                    control.Width = (int)(originalDataGridViewSize.Width * widthRatio);
                    control.Height = (int)(originalDataGridViewSize.Height * heightRatio);
                }
            }
        }
    }
    private void btnCalculateAllExcel_Click(object sender, EventArgs e)
    {
        // 確保 FileNames 和 lbFileInfo 有相同數量的項目
        if (FileNames.Count != lbFileInfo.Items.Count)
        {
            MessageBox.Show("FileNames 和 lbFileInfo 的項目數量不一致");
            return;
        }

        // 如果沒有選擇任何檔案，則顯示提示訊息
        if (FileNames.Count == 0)
        {
            MessageBox.Show("請選擇檔案");
            return;
        }

        // 為每個檔案創建一個 tabPage 和一個 dataGridView
        for (int i = 0; i < FileNames.Count; i++)
        {
            string fileName = FileNames[i];
            string fileInfo = lbFileInfo.Items[i].ToString();
            float fontSize;

            // 嘗試將 tbFontSize.Text 轉換為浮點數
            if (!float.TryParse(tbFontSize.Text, out fontSize))
            {
                // 如果轉換失敗，可以設置一個默認的字體大小，這裡假設為 10
                fontSize = 10f;
            }

            // 如果 tabPage 已經存在，則選擇該 tabPage
            // 否則，創建一個新的 tabPage 和一個新的 dataGridView
            TabPage tabPage;
            DataGridView dataGridView;
            if (tabControl1.TabPages.Count > i)
            {
                tabPage = tabControl1.TabPages[i];
                dataGridView = (DataGridView)tabPage.Controls[0];
            }
            else
            {
                tabPage = new TabPage($"Tab {i + 1}");
                tabPage.Text = "Computing";
                dataGridView = new DataGridView
                {
                    Dock = DockStyle.Fill,
                };
                tabPage.Controls.Add(dataGridView);
                tabControl1.TabPages.Add(tabPage);
            }

            // 設置 DataGridView 的字體大小
            dataGridView.Font = new System.Drawing.Font(dataGridView.Font.FontFamily, fontSize);

            // 選擇 tabPage
            tabControl1.SelectedTab = tabPage;

            // 開啟 Excel 檔案
            OpenExcelFile(fileName, fileInfo + ".xlsx", dataGridView);
        }
    }
    private string[] CatagoryArray(string input)
    {
        // 获取 ListBox 中项目的索引
        int index = FileNames.IndexOf(input);
        if (index == -1)
        {
            return new string[] { };
        }

        // 获取指定索引的项目
        string fileInfo = lbFileInfo.Items[index].ToString();

        // 根据 fileInfo 和 rbWeek 的选中状态返回相应的数组
        if (fileInfo == "主日")
        {
            if (rbWeek.Checked)
            {
                return new string[] { tbSheet1WeekCat1.Text, tbSheet1WeekCat2.Text };
            }
            else
            {
                return new string[] { tbSheet1Cat1.Text, tbSheet1Cat2.Text, tbSheet1Cat3.Text };
            }
        }
        else if (fileInfo == "禱告")
        {
            if (rbWeek.Checked)
            {
                return new string[] { tbSheet2WeekCat1.Text, tbSheet2WeekCat2.Text };
            }
            else
            {
                return new string[] { tbSheet2Cat1.Text, tbSheet2Cat2.Text, tbSheet2Cat3.Text };
            }
        }
        else if (fileInfo == "小排")
        {
            if (rbWeek.Checked)
            {
                return new string[] { tbSheet3WeekCat1.Text, tbSheet3WeekCat2.Text };
            }
            else
            {
                return new string[] { tbSheet3Cat1.Text, tbSheet3Cat2.Text, tbSheet3Cat3.Text };
            }
        }
        else if (fileInfo == "晨興")
        {
            if (rbWeek.Checked)
            {
                return new string[] { tbSheet4WeekCat1.Text, tbSheet4WeekCat2.Text };
            }
            else
            {
                return new string[] { tbSheet4Cat1.Text, tbSheet4Cat2.Text, tbSheet4Cat3.Text };
            }
        }
        else
        {
            if (rbWeek.Checked)
            {
                return new string[] { tbSheet1WeekCat1.Text, tbSheet1WeekCat2.Text };
            }
            else
            {
                return new string[] { tbSheet1Cat1.Text, tbSheet1Cat2.Text, tbSheet1Cat3.Text };
            }
        }
    }


    private void AttendForm_Load(object sender, EventArgs e)
    {
        ControlState controlState;
        if (File.Exists("controlState.json"))
        {
            var json = File.ReadAllText("controlState.json");
            controlState = JsonConvert.DeserializeObject<ControlState>(json);
            rbWeek.Checked = controlState.RbWeek;
            rbHalfYear.Checked = controlState.RbHalfYear;
            rbMonth.Checked = controlState.RbMonth;
            rbSelfDef.Checked = controlState.RbSelfDef;
            txtbIgnoreLevel.Text = controlState.TxtbIgnoreLevel;
            txtBoxStable.Text = controlState.TxtBoxStable;
            txtBoxStartColumn.Text = controlState.TxtBoxStartColumn;
            tbSheet1WeekCat2.Text = controlState.TbSheet1WeekCat2;
            tbSheet1WeekCat1.Text = controlState.TbSheet1WeekCat1;
            tbSheet4Cat3.Text = controlState.TbSheet4Cat3;
            tbSheet3Cat3.Text = controlState.TbSheet3Cat3;
            tbSheet2Cat3.Text = controlState.TbSheet2Cat3;
            tbSheet1Cat3.Text = controlState.TbSheet1Cat3;
            tbSheet4Cat2.Text = controlState.TbSheet4Cat2;
            tbSheet4Cat1.Text = controlState.TbSheet4Cat1;
            tbSheet3Cat2.Text = controlState.TbSheet3Cat2;
            tbSheet3Cat1.Text = controlState.TbSheet3Cat1;
            tbSheet2Cat2.Text = controlState.TbSheet2Cat2;
            tbSheet2Cat1.Text = controlState.TbSheet2Cat1;
            tbSheet1Cat2.Text = controlState.TbSheet1Cat2;
            tbSheet1Cat1.Text = controlState.TbSheet1Cat1;
            tbSheet4WeekCat2.Text = controlState.TbSheet4WeekCat2;
            tbSheet4WeekCat1.Text = controlState.TbSheet4WeekCat1;
            tbSheet3WeekCat2.Text = controlState.TbSheet3WeekCat2;
            tbSheet3WeekCat1.Text = controlState.TbSheet3WeekCat1;
            tbSheet2WeekCat2.Text = controlState.TbSheet2WeekCat2;
            tbSheet2WeekCat1.Text = controlState.TbSheet2WeekCat1;
            tbSelfDefWeek.Text = controlState.TbSelfDefWeek;
            cbIgnoreNoData.Checked = controlState.CbIgnoreNoData;
            cbIgnoreElementarySchool.Checked = controlState.CbIgnoreElementarySchool;
            cbDisplayLastAttend.Checked = controlState.CbDisplayLastAttend;
            cbAllDisplayLastAttend.Checked = controlState.CbAllDisplayLastAttend;
            ckbCompare.Checked = controlState.CkbCompare;
            ckbFwdBwd.Checked = controlState.CkbFwdBwd;
            tbFontSize.Text = controlState.TbFontSize;
            lbClrStat1.Text = controlState.LbClrStat1;
            lbClrStat1.BackColor = ColorTranslator.FromHtml(lbClrStat1.Text);
            lbClrStat2.Text = controlState.LbClrStat2;
            lbClrStat2.BackColor = ColorTranslator.FromHtml(lbClrStat2.Text);
            lbClrPos.Text = controlState.LbClrPos;
            lbClrPos.BackColor = ColorTranslator.FromHtml(lbClrPos.Text);
            lbClrNeg.Text = controlState.LbClrNeg;
            lbClrNeg.BackColor = ColorTranslator.FromHtml(lbClrNeg.Text);
            lbClrHdr1.Text = controlState.LbClrHdr1;
            lbClrHdr1.BackColor = ColorTranslator.FromHtml(lbClrHdr1.Text);
            lbClrHdr2.Text = controlState.LbClrHdr2;
            lbClrHdr2.BackColor = ColorTranslator.FromHtml(lbClrHdr2.Text);

            // ... 其他控件
        }
        else
        {
            lbClrStat1.BackColor = ColorTranslator.FromHtml(lbClrStat1.Text);
            lbClrStat2.BackColor = ColorTranslator.FromHtml(lbClrStat2.Text);
            lbClrPos.BackColor = ColorTranslator.FromHtml(lbClrPos.Text);
            lbClrNeg.BackColor = ColorTranslator.FromHtml(lbClrNeg.Text);
            lbClrHdr1.BackColor = ColorTranslator.FromHtml(lbClrHdr1.Text);
            lbClrHdr2.BackColor = ColorTranslator.FromHtml(lbClrHdr2.Text);
        }

        //  LoadDataGridViews();

    }
    private void SetBackColorFromHex(TextBox textBox, string hexColor)
    {
        try
        {
            Color color = ColorTranslator.FromHtml(hexColor);
            textBox.BackColor = color;
            textBox.Text = hexColor.ToUpper();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Invalid color format: {hexColor}. Error: {ex.Message}");
        }
    }
    private void AttendForm_FormClosing(object sender, FormClosingEventArgs e)
    {
        var controlState = new ControlState
        {
            RbWeek = rbWeek.Checked,
            RbHalfYear = rbHalfYear.Checked,
            RbMonth = rbMonth.Checked,
            RbSelfDef = rbSelfDef.Checked,
            TxtbIgnoreLevel = txtbIgnoreLevel.Text,
            TxtBoxStable = txtBoxStable.Text,
            TxtBoxStartColumn = txtBoxStartColumn.Text,
            TbSheet1WeekCat2 = tbSheet1WeekCat2.Text,
            TbSheet1WeekCat1 = tbSheet1WeekCat1.Text,
            TbSheet4Cat3 = tbSheet4Cat3.Text,
            TbSheet3Cat3 = tbSheet3Cat3.Text,
            TbSheet2Cat3 = tbSheet2Cat3.Text,
            TbSheet1Cat3 = tbSheet1Cat3.Text,
            TbSheet4Cat2 = tbSheet4Cat2.Text,
            TbSheet4Cat1 = tbSheet4Cat1.Text,
            TbSheet3Cat2 = tbSheet3Cat2.Text,
            TbSheet3Cat1 = tbSheet3Cat1.Text,
            TbSheet2Cat2 = tbSheet2Cat2.Text,
            TbSheet2Cat1 = tbSheet2Cat1.Text,
            TbSheet1Cat2 = tbSheet1Cat2.Text,
            TbSheet1Cat1 = tbSheet1Cat1.Text,
            TbSheet4WeekCat2 = tbSheet4WeekCat2.Text,
            TbSheet4WeekCat1 = tbSheet4WeekCat1.Text,
            TbSheet3WeekCat2 = tbSheet3WeekCat2.Text,
            TbSheet3WeekCat1 = tbSheet3WeekCat1.Text,
            TbSheet2WeekCat2 = tbSheet2WeekCat2.Text,
            TbSheet2WeekCat1 = tbSheet2WeekCat1.Text,
            TbSelfDefWeek = tbSelfDefWeek.Text,
            CbIgnoreNoData = cbIgnoreNoData.Checked,
            CbIgnoreElementarySchool = cbIgnoreElementarySchool.Checked,
            CbDisplayLastAttend = cbDisplayLastAttend.Checked,
            CbAllDisplayLastAttend = cbAllDisplayLastAttend.Checked,
            CkbCompare = ckbCompare.Checked,
            CkbFwdBwd = ckbFwdBwd.Checked,
            TbFontSize = tbFontSize.Text,
            LbClrStat1 = lbClrStat1.Text,
            LbClrStat2 = lbClrStat2.Text,
            LbClrHdr1 = lbClrHdr1.Text,
            LbClrHdr2 = lbClrHdr2.Text,
            LbClrNeg = lbClrNeg.Text,
            LbClrPos = lbClrPos.Text,
            // ... 其他控件
        };

        var json = JsonConvert.SerializeObject(controlState);
        File.WriteAllText("controlState.json", json);
        // SaveDataGridViews();
        SaveAttendanceRecords();
    }

    private string ShortenDateRange(string input)
    {
        var parts = input.Split('~');
        if (parts.Length != 2) return input; // Return original input if format is not as expected

        var startParts = parts[0].Split('年', '月', '週');
        var endParts = parts[1].Split('年', '月', '週');

        if (startParts.Length < 3 || endParts.Length < 3) return input; // Return original input if format is not as expected

        string year1 = startParts[0];
        string month1 = startParts[1];
        string week1 = startParts[2];

        string year2 = endParts[0];
        string month2 = endParts[1];
        string week2 = endParts[2];

        if (year1 == year2)
        {
            if (month1 == month2)
            {
                return $"{year1}年{month1}月{week1}週~{week2}週";
            }
            else
            {
                return $"{year1}年{month1}月{week1}週~{month2}月{week2}週";
            }
        }
        else
        {
            return input; // If years are different, return original input
        }
    }


    private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
    {
        // 將所有的tabpage的前景顏色設為淺灰色
        foreach (TabPage tp in tabControl1.TabPages)
        {
            tp.ForeColor = System.Drawing.Color.LightGray;
        }

        // 將被選擇的tabpage的前景顏色設為黑色
        tabControl1.SelectedTab.ForeColor = System.Drawing.Color.Black;
    }

    private void btnAddNewFile_Click(object sender, EventArgs e)
    {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        // 只顯示 .xls 檔案
        openFileDialog.Filter = "Excel 檔案 (*.xls)|*.xls";
        // 允許選擇多個檔案
        openFileDialog.Multiselect = true;

        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            foreach (string fileName in openFileDialog.FileNames)
            {
                string firstCell = ReadFirstCell(fileName);

                // 檢查 firstCell 是否已經存在於 lbFileInfo 中
                int index = lbFileInfo.Items.IndexOf(firstCell);
                if (index != -1)
                {
                    // 如果存在，則移除舊的對應檔案
                    FileNames.RemoveAt(index);
                    lbFileInfo.Items.RemoveAt(index);
                }

                // 將檔案名稱添加到列表中
                FileNames.Add(fileName);

                // 將 firstCell 添加到 lbFileInfo 中
                lbFileInfo.Items.Add(firstCell);
            }
        }
    }

    private void btnRemoveFile_Click(object sender, EventArgs e)
    {
        if (lbFileInfo.SelectedIndex != -1)
        {
            FileNames.RemoveAt(lbFileInfo.SelectedIndex);
            lbFileInfo.Items.RemoveAt(lbFileInfo.SelectedIndex);
        }
    }

    private void ColorPicker(object sender, EventArgs e)
    {
        Label lbClr = sender as Label;

        Color initialColor = ColorTranslator.FromHtml(lbClr.Text);
        colorDialog.Color = initialColor;

        if (colorDialog.ShowDialog() == DialogResult.OK)
        {
            // Get the selected color
            Color selectedColor = colorDialog.Color;

            // Update the TextBox background color and text
            lbClr.BackColor = selectedColor;
            lbClr.Text = ColorToHex(selectedColor);
        }

    }
    private string ColorToHex(Color color)
    {
        return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
    }

    private void lbClrHdr1_Click(object sender, EventArgs e)
    {
        ColorPicker(sender, e);
    }

    private void lbClrHdr2_Click(object sender, EventArgs e)
    {
        ColorPicker(sender, e);
    }

    private void lbClrPos_Click(object sender, EventArgs e)
    {
        ColorPicker(sender, e);
    }

    private void lbClrNeg_Click(object sender, EventArgs e)
    {
        ColorPicker(sender, e);
    }

    private void lbClrHdr1_Click_1(object sender, EventArgs e)
    {
        ColorPicker(sender, e);
    }

    private void lbClrHdr2_Click_1(object sender, EventArgs e)
    {
        ColorPicker(sender, e);
    }
    private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
    {
        // 绘制标签页的标题文本
        TabControl tabControl = (TabControl)sender;
        TabPage tabPage = tabControl.TabPages[e.Index];
        string tabText = tabPage.Text;

        // 根据选中状态设置文本颜色
        Color textColor = (e.State == DrawItemState.Selected) ? SystemColors.ControlText : SystemColors.GrayText;

        // 使用自定义字体绘制标签页文本
        using (Font tabFont = new Font(tabControl.Font.FontFamily, (e.State == DrawItemState.Selected) ? 14 : 10)) // 设置不同的字体大小
        {
            // 绘制文本
            e.Graphics.DrawString(tabText, tabFont, new SolidBrush(textColor), e.Bounds.X + 3, e.Bounds.Y + 3);
        }
    }
    private void tabControl1_SelectedIndexChanged_1(object sender, EventArgs e)
    {
        for (int i = 0; i < tabControl1.TabCount; i++)
        {
            tabControl1.TabPages[i].Font = new Font(tabControl1.Font.FontFamily, (i == tabControl1.SelectedIndex) ? 14 : 10);
        }
    }

    private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
    {
        // 绘制标签页的标题文本
        TabControl tabControl = (TabControl)sender;
        TabPage tabPage = tabControl.TabPages[e.Index];
        string tabText = tabPage.Text;

        // 根据选中状态设置文本颜色
        Color textColor = (e.State == DrawItemState.Selected) ? SystemColors.ControlText : SystemColors.GrayText;

        // 使用自定义字体绘制标签页文本
        using (Font tabFont = new Font(tabControl.Font.FontFamily, (e.State == DrawItemState.Selected) ? 10 : 9)) // 设置不同的字体大小
        {
            // 绘制文本
            e.Graphics.DrawString(tabText, tabFont, new SolidBrush(textColor), e.Bounds.X + 3, e.Bounds.Y + 3);
        }
    }
}