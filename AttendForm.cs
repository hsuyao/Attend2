﻿using System;
using System.IO;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using CellType = NPOI.SS.UserModel.CellType;
using IndexedColors = NPOI.SS.UserModel.IndexedColors;
using DocumentFormat.OpenXml.Office2016.Excel;

namespace Attend;

public partial class AttendForm : Form
{
    HSSFWorkbook workbook;
    public AttendForm()
    {
        InitializeComponent();
    }

    public void ConvertHssfToXssf(string inputFileName, string outputFileName)
    {
        // 讀取 .xls 文件
        HSSFWorkbook hssfWorkbook;
        using (FileStream file = new FileStream(inputFileName, FileMode.Open, FileAccess.Read))
        {
            hssfWorkbook = new HSSFWorkbook(file);
        }

        // 讀取或創建一個 .xlsx 文件
        XSSFWorkbook xssfWorkbook;
        if (File.Exists(outputFileName))
        {
            using (FileStream file = new FileStream(outputFileName, FileMode.Open, FileAccess.Read))
            {
                xssfWorkbook = new XSSFWorkbook(file);
            }
        }
        else
        {
            xssfWorkbook = new XSSFWorkbook();
        }

        // 遍歷所有的工作表
        for (int i = 0; i < hssfWorkbook.NumberOfSheets; i++)
        {
            ISheet hssfSheet = hssfWorkbook.GetSheetAt(i);
            ISheet xssfSheet = xssfWorkbook.GetSheet(hssfSheet.SheetName);

            // 如果工作表不存在，則創建一個新的
            if (xssfSheet == null)
            {
                xssfSheet = xssfWorkbook.CreateSheet(hssfSheet.SheetName);
            }

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
    }

    private void btnSelect_Click(object sender, EventArgs e)
    {
        OpenFileDialog ofd = new OpenFileDialog();
        ofd.ShowDialog();
        txtBoxLord.Text = ofd.FileName;
        btnCalculate.Enabled = true;
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
        int columnIndex = int.Parse(input.Split('-')[0]);

        IRow row0 = sheet.GetRow(0);
        IRow row1 = sheet.GetRow(1);
        ICell cell0 = row0?.GetCell(columnIndex);
        ICell cell1 = row1?.GetCell(columnIndex);

        string cellValue0 = cell0?.StringCellValue ?? string.Empty;
        string cellValue1 = cell1?.StringCellValue ?? string.Empty;

        // 如果Row(0)內容是空的，往前面的column搜尋以取得該column的Row(0)內容
        if (string.IsNullOrEmpty(cellValue0))
        {
            for (int i = columnIndex - 1; i >= 0; i--)
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

        return newCellValue;
    }
    private void WriteToExcel(Dictionary<string, double> result, string filePath, string sheetName, int startRow, int startColumn)
    {
        XSSFWorkbook workbook;
        using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(file);
        }

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
        var lightBlue = new XSSFColor(new byte[] { 145, 203, 232 });
        var navyBlue = new XSSFColor(new byte[] { 157, 189, 220 });

        bool isLightBlue = true;

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
                if (isLightBlue == true)
                {
                    style.SetFillForegroundColor(lightBlue);
                }
                else
                {
                    style.SetFillForegroundColor(navyBlue);
                }

                style.FillPattern = FillPattern.SolidForeground;
                for (int i = startColumn; i <= startColumn + 2; i++)
                {
                    row.GetCell(i).CellStyle = style;
                }
            }

            // 切換顏色
            isLightBlue = !isLightBlue;
        }

        MergeColumnCells(sheet, startColumn);
        sheet.SetColumnWidth(startColumn, 14 * 256);
        sheet.SetColumnWidth(startColumn + 1, 10 * 256);
        sheet.SetColumnWidth(startColumn + 2, 4 * 256);

        using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Write))
        {
            workbook.Write(file);
        }
    }

    private void OpenExcelFile()
    {
        string startColumnLetter = txtBoxStartColumn.Text; // 使用者輸入的開始列名
        int startColumnIndex = startColumnLetter.ToUpper()[0] - 'A'; // 轉換列名為索引
        ConvertHssfToXssf(txtBoxLord.Text, txtBoxOutput.Text);

        using (var fs = new FileStream(txtBoxOutput.Text, FileMode.Open, FileAccess.Read))
        {
            var workbook = new XSSFWorkbook(fs);
            var sheet = workbook.GetSheetAt(0); // 選擇第一個工作表
            var row = sheet.GetRow(1); // 選擇第五行

            int lastColumnWithData = startColumnIndex;
            lastColumnWithData = GetLastColumnWithData(sheet, 1, startColumnIndex); // 分析第二列
                                                                                    // ICell cell = row.GetCell(lastColumnWithData);
                                                                                    // MessageBox.Show("從 '" + startColumnLetter + "' 列開始，共有資料筆數: " + (lastColumnWithData-startColumnIndex+1)+ " "+cell);

            if (rbMonth.Checked == true)
            {
                List<string> result = GroupByMonth(sheet);
                var sheetName = new List<string>();
                foreach (string s in result)
                {
                    // MessageBox.Show(s);
                    var dict = ClassifyAttendancy(s, sheet);
                    var month = GetMonthString(sheet, s);
                    sheetName.Add(month);
                    FillSheetWithDict(dict, month, txtBoxOutput.Text, false);
                    var byIdentity = AttendanceCountByIdentity(s, sheet);
                    var calculateAverageResult = CalculateAverage(byIdentity);
                    WriteToExcel(calculateAverageResult, txtBoxOutput.Text, month, 1, 0);
                }
                for (int i = 1; i < sheetName.Count; i++)
                {
                    // 取得當前表單和前一個表單的名稱
                    string currentSheetName = sheetName[i];
                    string previousSheetName = sheetName[i - 1];

                    CompareSheets(txtBoxOutput.Text, currentSheetName, previousSheetName);
                }
            }
            if (rbWeek.Checked == true)
            {
                lastColumnWithData = GetLastColumnWithData(sheet, 1, startColumnIndex); // 分析第二列
                List<string> result = GroupNumbers(startColumnIndex, lastColumnWithData, 1);
                var sheetName = new List<string>();
                foreach (string s in result)
                {
                    // MessageBox.Show(s);
                    var dict = ClassifyAttendancy(s, sheet);
                    var week = GetWeekString(sheet, s);
                    sheetName.Add(week);
                    FillSheetWithDict(dict, week, txtBoxOutput.Text, false);
                    var byIdentity = AttendanceCountByIdentity(s, sheet);
                    var calculateAverageResult = CalculateAverage(byIdentity);
                }
                for (int i = 1; i < sheetName.Count; i++)
                {
                    // 取得當前表單和前一個表單的名稱
                    string currentSheetName = sheetName[i];
                    string previousSheetName = sheetName[i - 1];

                    CompareSheets(txtBoxOutput.Text, currentSheetName, previousSheetName);
                }
            }
            if (rbHalfYear.Checked == true)
            {
                lastColumnWithData = GetLastColumnWithData(sheet, 1, startColumnIndex); // 分析第二列
                List<string> result = GroupNumbers(startColumnIndex, lastColumnWithData, 26);
                foreach (string s in result)
                {
                    // MessageBox.Show(s);
                    var dict = ClassifyAttendancy(s, sheet);
                    FillSheetWithDict(dict, s, txtBoxOutput.Text, false);
                    var byIdentity = AttendanceCountByIdentity(s, sheet);
                    var calculateAverageResult = CalculateAverage(byIdentity);
                }
            }
            MessageBox.Show("Finished!");
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
                    if (sheet.GetRow(1).GetCell(i + j + 1).ToString() == "第一週" || sheet.GetRow(1).GetCell(i + j + 1).ToString() == "")
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

    private List<string> GroupNumbers(long startNum, long endNum, long groupSize)
    {
        var groups = new List<string>();
        long total = endNum - startNum + 1;
        long numGroups = (long)Math.Ceiling((double)total / groupSize);

        for (long i = numGroups - 1; i >= 0; i--)
        {
            long groupEndNum = startNum + groupSize - 1;
            if (groupEndNum > endNum) groupEndNum = endNum;
            groups.Add(startNum + "-" + groupEndNum);
            startNum = groupEndNum + 1;
        }

        return groups;
    }
    private Dictionary<string, List<string>> ClassifyAttendancy(string columnRange, ISheet sheet)
    {
        var dict = new Dictionary<string, List<string>>();

        // 使用 Split 函數來分割字串
        string[] parts = columnRange.Split('-');
        int startColumn = int.Parse(parts[0]);
        int endColumn = int.Parse(parts[1]);

        for (int i = 2; i <= sheet.LastRowNum; i++)
        {
            var row = sheet.GetRow(i);
            string groupName = row.GetCell(0)?.ToString() + row.GetCell(1)?.ToString();

            if (string.IsNullOrEmpty(groupName))
                continue;

            string[] categories = startColumn == endColumn ? new[] { "本週到會", "未到會  " } : new[] { "正常聚會", "不穩定", "無紀錄" };

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
                if (row.GetCell(j)?.NumericCellValue == 1)
                    attendanceCount++;
                weekCount++;
            }

            double attendanceRate = (double)attendanceCount / weekCount;
            string category;

            if (startColumn == endColumn)
                category = attendanceRate > double.Parse(txtBoxStable.Text) ? "本週到會" : "未到會  ";
            else
                category = attendanceRate > double.Parse(txtBoxStable.Text) ? "正常聚會" : attendanceRate > 0 ? "不穩定" : "無紀錄";

            string key = groupName + "_" + category;

            if (!dict.ContainsKey(key))
                dict[key] = new List<string>();

            dict[key].Add(name + "_" + attendanceRate.ToString("P"));
        }

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
                int weekAttendance = (int)row.GetCell(j)?.NumericCellValue;

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
            double lowerBound = median * double.Parse(txtbIgnoreLevel.Text);
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

    private void FillSheetWithDict(Dictionary<string, List<string>> dict, string sheetName, string filePath, bool replaceUnderscore)
    {
        IWorkbook workbook;
        using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
        {
            workbook = new XSSFWorkbook(stream);
        }

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

        MergeCells(sheet, 0);
        // Save the changes and close the workbook
        using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
        {
            workbook.Write(stream);
        }
        

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


    private void CompareSheets(string fileName, string mainSheetName, string compareSheetName)
    {
        XSSFWorkbook workbook;
        using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(file);
        }

        ISheet mainSheet = workbook.GetSheet(mainSheetName);
        ISheet compareSheet = workbook.GetSheet(compareSheetName);

        // Define the colors
        XSSFColor lightGreen = new XSSFColor(new byte[] { 44, 238, 144 }); // light green
        XSSFColor lightRed = new XSSFColor(new byte[] { 255, 210, 210 }); // light red

        // Loop through columns
        for (int col = 1; col < mainSheet.GetRow(0).LastCellNum; col++)
        {
            // Compare column
            for (int row = 2; row <= mainSheet.LastRowNum; row++)
            {
                ICell mainCell = mainSheet.GetRow(row)?.GetCell(col);
                string mainCellValue = mainCell?.ToString() ?? string.Empty;

                int bracketPos = mainCellValue.IndexOf("(");
                string removeAfterBracket = bracketPos > 0 ? mainCellValue.Substring(0, bracketPos) : mainCellValue;

                if (!string.IsNullOrEmpty(mainCellValue) && !CellExistsInColumn(compareSheet, col, removeAfterBracket))
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

        // 儲存變更回 Excel 檔案
        using (FileStream file = new FileStream(fileName, FileMode.Create, FileAccess.Write))
        {
            workbook.Write(file);
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


    private void textBox1_TextChanged(object sender, EventArgs e)
    {

    }

    private void label4_Click(object sender, EventArgs e)
    {

    }
}
