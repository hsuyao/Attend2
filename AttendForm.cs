using System;
using System.IO;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;

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

            string[] categories = startColumn == endColumn ? new[] { "本週到會", "未到會  " } : new[] { "正常聚會", "不穩定", "不聚會" };

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
                category = attendanceRate > double.Parse(txtBoxStable.Text) ? "正常聚會" : attendanceRate > 0 ? "不穩定" : "不聚會";

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

            if (string.IsNullOrEmpty(district))
                continue;

            string identity = row.GetCell(5)?.ToString();

            for (int j = startColumn; j <= endColumn; j++)
            {
                int weekAttendance = (int)row.GetCell(j)?.NumericCellValue;

                if (weekAttendance == 1)
                {
                    string key = district + "|" + identity + "|" + (j - 7); // 小區|身分|週次

                    if (dict.ContainsKey(key))
                    {
                        dict[key][0] += 1;
                    }
                    else
                    {
                        dict[key] = new List<int> { 1 };
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
                }else
                {
                    count = count;// debug purpose
                }
            }

            // 計算平均數並更新字典
            result[key] = Math.Round(sum / count, 0);
        }

        return result;
    }

private void lblCurFile_Click(object sender, EventArgs e)
    {

    }

    private void txtBoxStable_TextChanged(object sender, EventArgs e)
    {

    }
}
