using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.Drawing;
using System.Web;


public class AttendanceRecordComparer : IEqualityComparer<AttendanceRecord>
{
    public bool Equals(AttendanceRecord x, AttendanceRecord y)
    {
        if (x == null || y == null)
            return false;

        return x.Name == y.Name && x.Date == y.Date;
    }

    public int GetHashCode(AttendanceRecord record)
    {
        if (record == null)
            return 0;

        // Combine hash codes of Name and Date for uniqueness
        int hash = record.Name.GetHashCode();
        hash = (hash * 31) + record.Date.GetHashCode();
        return hash;
    }
}
public class AttendanceRecord
{
    public string Name { get; set; }
    public string Date { get; set; }
    public int Attendance { get; set; }
}

namespace Attend
{
    internal class ControlState
    {
        public bool RbWeek { get; set; }
        public bool RbHalfYear { get; set; }
        public bool RbMonth { get; set; }
        public bool RbSelfDef { get; set; }
        public string TxtbIgnoreLevel { get; set; }
        public string TxtBoxStable { get; set; }
        public string TxtBoxStartColumn { get; set; }
        public string TbSheet1WeekCat2 { get; set; }
        public string TbSheet1WeekCat1 { get; set; }
        public string TbSheet4Cat3 { get; set; }
        public string TbSheet3Cat3 { get; set; }
        public string TbSheet2Cat3 { get; set; }
        public string TbSheet1Cat3 { get; set; }
        public string TbSheet4Cat2 { get; set; }
        public string TbSheet4Cat1 { get; set; }
        public string TbSheet3Cat2 { get; set; }
        public string TbSheet3Cat1 { get; set; }
        public string TbSheet2Cat2 { get; set; }
        public string TbSheet2Cat1 { get; set; }
        public string TbSheet1Cat2 { get; set; }
        public string TbSheet1Cat1 { get; set; }
        public string TbSheet4WeekCat2 { get; set; }
        public string TbSheet4WeekCat1 { get; set; }
        public string TbSheet3WeekCat2 { get; set; }
        public string TbSheet3WeekCat1 { get; set; }
        public string TbSheet2WeekCat2 { get; set; }
        public string TbSheet2WeekCat1 { get; set; }
        public string TbSelfDefWeek { get; set; }
        public bool CbIgnoreNoData { get; set; }
        public bool CbIgnoreElementarySchool { get; set; }
        public bool CkbCompare { get; set; }
        public bool CkbFwdBwd { get; set; }
        public string TbFontSize { get; set; }
        public string LbClrHdr1 { get; set; }
        public string LbClrHdr2 { get; set; }
        public string LbClrPos { get; set; }
        public string LbClrNeg { get; set; }
        public string LbClrStat1 { get; set; }
        public string LbClrStat2 { get; set; }
        // ... 其他控件
    }
    internal class DataGridViewData
    {
        public string[,] CellValues { get; set; }
        public string[,] CellColors { get; set; }
    }

}
