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

namespace Attend;

public partial class AttendForm : Form
{
    HSSFWorkbook workbook;
    public AttendForm()
    {
        InitializeComponent();
    }

    private void btnSelect_Click(object sender, EventArgs e)
    {
        OpenFileDialog ofd = new OpenFileDialog();
        ofd.ShowDialog();
        lblCurFile.Text = ofd.FileName;
        btnCalculate.Enabled = true;
    }

    private void OpenExcelFile()
    {
        string startColumnLetter = txtBoxStartColumn.Text; // 使用者輸入的開始列名
        int startColumnIndex = startColumnLetter.ToUpper()[0] - 'A'; // 轉換列名為索引

        using (var fs = new FileStream(lblCurFile.Text, FileMode.Open, FileAccess.Read))
        {
            var workbook = new HSSFWorkbook(fs);
            var sheet = workbook.GetSheetAt(0); // 選擇第一個工作表
            var row = sheet.GetRow(1); // 選擇第五行

            int lastColumnWithData = startColumnIndex;
            lastColumnWithData = GetLastColumnWithData(sheet, 1, startColumnIndex); // 分析第二列
            // ICell cell = row.GetCell(lastColumnWithData);
            // MessageBox.Show("從 '" + startColumnLetter + "' 列開始，共有資料筆數: " + (lastColumnWithData-startColumnIndex+1)+ " "+cell);

            if (rbMonth.Checked == true)
            {
                List<string> result = GroupByMonth(sheet);
                foreach (string s in result)
                {
                    // MessageBox.Show(s);
                    var dict = ClassifyAttendancy(s, sheet);
                    FillSheetWithDict(dict, s, lblCurFile.Text, false);
                    var byIdentity = AttendanceCountByIdentity(s, sheet);
                    var calculateAverageResult = CalculateAverage(byIdentity);
                }
            }
            if (rbWeek.Checked == true)
            {
                lastColumnWithData = GetLastColumnWithData(sheet, 1, startColumnIndex); // 分析第二列
                List<string> result = GroupNumbers(startColumnIndex, lastColumnWithData, 1);
                foreach (string s in result)
                {
                    // MessageBox.Show(s);
                    var dict = ClassifyAttendancy(s, sheet);
                    FillSheetWithDict(dict, s, lblCurFile.Text, false);
                    var byIdentity = AttendanceCountByIdentity(s, sheet);
                    var calculateAverageResult = CalculateAverage(byIdentity);
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
                    FillSheetWithDict(dict, s, lblCurFile.Text, false);
                    var byIdentity = AttendanceCountByIdentity(s, sheet);
                    var calculateAverageResult = CalculateAverage(byIdentity);
                }
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
                    if (dict.ContainsKey(key2))
                    {
                        dict[key2][0] += 1;
                    }
                    else
                    {
                        dict[key2] = new List<int> { 1 };
                    }
                }
            }
        }

        return dict;
    }
    private Dictionary<string, double> CalculateAverage(Dictionary<string, List<int>> dict)
    {
        var districtIdentityDict = new Dictionary<string, List<int>>();

        // 分類每個小區和身份的數字
        foreach (var key in dict.Keys)
        {
            var parts = key.Split('|');
            var district = parts[0];
            var identity = parts[1];
            var districtIdentityKey = district + "|" + identity;

            if (!districtIdentityDict.ContainsKey(districtIdentityKey))
                districtIdentityDict[districtIdentityKey] = new List<int>();

            districtIdentityDict[districtIdentityKey].AddRange(dict[key]);
        }

        var result = new Dictionary<string, double>();

        // 計算每個小區和身份的中位數平均數
        foreach (var key in districtIdentityDict.Keys)
        {
            var values = districtIdentityDict[key];

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
            workbook = new HSSFWorkbook(stream);
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

        // Save the changes and close the workbook
        using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
        {
            workbook.Write(stream);
        }
        MergeCells(filePath, sheetName, 0);
    }



    private void MergeCells(string fileName, string sheetName, int rowNumber)
    {
        using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
        {
            HSSFWorkbook workbook = new HSSFWorkbook(file);
            ISheet sheet = workbook.GetSheet(sheetName);

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

            using (FileStream fileOut = new FileStream(fileName, FileMode.Open, FileAccess.Write))
            {
                workbook.Write(fileOut);
            }
        }
    }

}
