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

            List<string> result = GroupByMonth(sheet);
            foreach (string s in result)
            {
                MessageBox.Show(s);
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

    private List<string> GroupNumbers(long total, long groupSize)
    {
        var groups = new List<string>();
        long numGroups = (long)Math.Ceiling((double)total / groupSize);
        long endNum = total;

        for (long i = numGroups - 1; i >= 0; i--)
        {
            long startNum = endNum - groupSize + 1;
            if (startNum < 1) startNum = 1;
            groups.Add(startNum + "-" + endNum);
            endNum = startNum - 1;
        }

        return groups;
    }
}



