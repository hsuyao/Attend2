
namespace Attend;

partial class AttendForm
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    ///  Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        groupBox1 = new GroupBox();
        rbHalfYear = new RadioButton();
        rbMonth = new RadioButton();
        rbWeek = new RadioButton();
        tabPage3 = new TabPage();
        dgvResult3 = new DataGridView();
        tabPage2 = new TabPage();
        dgvResult2 = new DataGridView();
        tabPage1 = new TabPage();
        dgvResult1 = new DataGridView();
        tabControl1 = new TabControl();
        tabPage4 = new TabPage();
        dgvResult4 = new DataGridView();
        btnCalculateAllExcel = new Button();
        btnSelect1 = new Button();
        txtBoxSelect1 = new TextBox();
        btnSelect2 = new Button();
        txtBoxSelect2 = new TextBox();
        btnSelect3 = new Button();
        txtBoxSelect3 = new TextBox();
        label4 = new Label();
        label7 = new Label();
        label8 = new Label();
        btnSelect4 = new Button();
        txtBoxSelect4 = new TextBox();
        label9 = new Label();
        tabControl2 = new TabControl();
        tabPage5 = new TabPage();
        cbIgnoreNoData = new CheckBox();
        cbIgnoreElementarySchool = new CheckBox();
        ckbCompare = new CheckBox();
        label6 = new Label();
        label5 = new Label();
        txtbIgnoreLevel = new TextBox();
        label3 = new Label();
        txtBoxStable = new TextBox();
        label2 = new Label();
        label1 = new Label();
        txtBoxStartColumn = new TextBox();
        tabPage6 = new TabPage();
        groupBox2 = new GroupBox();
        tbSheet4Cat3 = new TextBox();
        tbSheet3Cat3 = new TextBox();
        tbSheet2Cat3 = new TextBox();
        tbSheet1Cat3 = new TextBox();
        label14 = new Label();
        tbSheet4Cat2 = new TextBox();
        tbSheet4Cat1 = new TextBox();
        label15 = new Label();
        tbSheet3Cat2 = new TextBox();
        tbSheet3Cat1 = new TextBox();
        label16 = new Label();
        tbSheet2Cat2 = new TextBox();
        tbSheet2Cat1 = new TextBox();
        label17 = new Label();
        tbSheet1Cat2 = new TextBox();
        tbSheet1Cat1 = new TextBox();
        groupBox3 = new GroupBox();
        label13 = new Label();
        tbSheet4WeekCat2 = new TextBox();
        tbSheet4WeekCat1 = new TextBox();
        label12 = new Label();
        tbSheet3WeekCat2 = new TextBox();
        tbSheet3WeekCat1 = new TextBox();
        label11 = new Label();
        tbSheet2WeekCat2 = new TextBox();
        tbSheet2WeekCat1 = new TextBox();
        label10 = new Label();
        tbSheet1WeekCat2 = new TextBox();
        tbSheet1WeekCat1 = new TextBox();
        splitContainer1 = new SplitContainer();
        groupBox1.SuspendLayout();
        tabPage3.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult3).BeginInit();
        tabPage2.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult2).BeginInit();
        tabPage1.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult1).BeginInit();
        tabControl1.SuspendLayout();
        tabPage4.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult4).BeginInit();
        tabControl2.SuspendLayout();
        tabPage5.SuspendLayout();
        tabPage6.SuspendLayout();
        groupBox2.SuspendLayout();
        groupBox3.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)splitContainer1).BeginInit();
        splitContainer1.SuspendLayout();
        SuspendLayout();
        // 
        // groupBox1
        // 
        groupBox1.Controls.Add(rbHalfYear);
        groupBox1.Controls.Add(rbMonth);
        groupBox1.Controls.Add(rbWeek);
        groupBox1.Location = new Point(498, 18);
        groupBox1.Name = "groupBox1";
        groupBox1.Size = new Size(173, 224);
        groupBox1.TabIndex = 7;
        groupBox1.TabStop = false;
        groupBox1.Text = "統計單位";
        // 
        // rbHalfYear
        // 
        rbHalfYear.AutoSize = true;
        rbHalfYear.Location = new Point(25, 118);
        rbHalfYear.Margin = new Padding(5);
        rbHalfYear.Name = "rbHalfYear";
        rbHalfYear.Size = new Size(121, 27);
        rbHalfYear.TabIndex = 11;
        rbHalfYear.Text = "半年(26週)";
        rbHalfYear.UseVisualStyleBackColor = true;
        // 
        // rbMonth
        // 
        rbMonth.AutoSize = true;
        rbMonth.Location = new Point(25, 80);
        rbMonth.Margin = new Padding(5);
        rbMonth.Name = "rbMonth";
        rbMonth.Size = new Size(137, 27);
        rbMonth.TabIndex = 10;
        rbMonth.Text = "月(四到五週)";
        rbMonth.UseVisualStyleBackColor = true;
        // 
        // rbWeek
        // 
        rbWeek.AutoSize = true;
        rbWeek.Checked = true;
        rbWeek.Location = new Point(25, 41);
        rbWeek.Margin = new Padding(5);
        rbWeek.Name = "rbWeek";
        rbWeek.Size = new Size(53, 27);
        rbWeek.TabIndex = 9;
        rbWeek.TabStop = true;
        rbWeek.Text = "週";
        rbWeek.UseVisualStyleBackColor = true;
        // 
        // tabPage3
        // 
        tabPage3.Controls.Add(dgvResult3);
        tabPage3.Location = new Point(4, 32);
        tabPage3.Name = "tabPage3";
        tabPage3.Padding = new Padding(3);
        tabPage3.Size = new Size(1568, 715);
        tabPage3.TabIndex = 2;
        tabPage3.Text = "第三表單最新統計結果";
        tabPage3.UseVisualStyleBackColor = true;
        // 
        // dgvResult3
        // 
        dgvResult3.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResult3.Dock = DockStyle.Fill;
        dgvResult3.Location = new Point(3, 3);
        dgvResult3.Margin = new Padding(5);
        dgvResult3.Name = "dgvResult3";
        dgvResult3.RowHeadersWidth = 62;
        dgvResult3.Size = new Size(1562, 709);
        dgvResult3.TabIndex = 4;
        // 
        // tabPage2
        // 
        tabPage2.Controls.Add(dgvResult2);
        tabPage2.Location = new Point(4, 32);
        tabPage2.Name = "tabPage2";
        tabPage2.Padding = new Padding(3);
        tabPage2.Size = new Size(1568, 715);
        tabPage2.TabIndex = 1;
        tabPage2.Text = "第二表單最新統計結果";
        tabPage2.UseVisualStyleBackColor = true;
        // 
        // dgvResult2
        // 
        dgvResult2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResult2.Dock = DockStyle.Fill;
        dgvResult2.Location = new Point(3, 3);
        dgvResult2.Margin = new Padding(5);
        dgvResult2.Name = "dgvResult2";
        dgvResult2.RowHeadersWidth = 62;
        dgvResult2.Size = new Size(1562, 709);
        dgvResult2.TabIndex = 2;
        // 
        // tabPage1
        // 
        tabPage1.Controls.Add(dgvResult1);
        tabPage1.Location = new Point(4, 32);
        tabPage1.Name = "tabPage1";
        tabPage1.Padding = new Padding(3);
        tabPage1.Size = new Size(1568, 715);
        tabPage1.TabIndex = 0;
        tabPage1.Text = "第一表單最新統計結果";
        tabPage1.UseVisualStyleBackColor = true;
        // 
        // dgvResult1
        // 
        dgvResult1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResult1.Dock = DockStyle.Fill;
        dgvResult1.Location = new Point(3, 3);
        dgvResult1.Margin = new Padding(5);
        dgvResult1.Name = "dgvResult1";
        dgvResult1.RowHeadersWidth = 62;
        dgvResult1.Size = new Size(1562, 709);
        dgvResult1.TabIndex = 1;
        // 
        // tabControl1
        // 
        tabControl1.Controls.Add(tabPage1);
        tabControl1.Controls.Add(tabPage2);
        tabControl1.Controls.Add(tabPage3);
        tabControl1.Controls.Add(tabPage4);
        tabControl1.Location = new Point(0, 267);
        tabControl1.Name = "tabControl1";
        tabControl1.SelectedIndex = 0;
        tabControl1.Size = new Size(1576, 751);
        tabControl1.TabIndex = 28;
        // 
        // tabPage4
        // 
        tabPage4.Controls.Add(dgvResult4);
        tabPage4.Location = new Point(4, 32);
        tabPage4.Margin = new Padding(5);
        tabPage4.Name = "tabPage4";
        tabPage4.Padding = new Padding(5);
        tabPage4.Size = new Size(1568, 715);
        tabPage4.TabIndex = 3;
        tabPage4.Text = "第四表單最新統計結果";
        tabPage4.UseVisualStyleBackColor = true;
        // 
        // dgvResult4
        // 
        dgvResult4.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResult4.Dock = DockStyle.Fill;
        dgvResult4.Location = new Point(5, 5);
        dgvResult4.Margin = new Padding(5);
        dgvResult4.Name = "dgvResult4";
        dgvResult4.RowHeadersWidth = 62;
        dgvResult4.Size = new Size(1558, 705);
        dgvResult4.TabIndex = 2;
        // 
        // btnCalculateAllExcel
        // 
        btnCalculateAllExcel.BackColor = SystemColors.ActiveCaption;
        btnCalculateAllExcel.Font = new Font("Microsoft JhengHei UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 136);
        btnCalculateAllExcel.Location = new Point(418, 25);
        btnCalculateAllExcel.Name = "btnCalculateAllExcel";
        btnCalculateAllExcel.Size = new Size(60, 210);
        btnCalculateAllExcel.TabIndex = 36;
        btnCalculateAllExcel.Text = "統計已選擇檔案";
        btnCalculateAllExcel.UseVisualStyleBackColor = false;
        btnCalculateAllExcel.Click += btnCalculateAllExcel_Click;
        // 
        // btnSelect1
        // 
        btnSelect1.BackColor = SystemColors.ActiveCaption;
        btnSelect1.Location = new Point(19, 31);
        btnSelect1.Margin = new Padding(5);
        btnSelect1.Name = "btnSelect1";
        btnSelect1.Size = new Size(157, 38);
        btnSelect1.TabIndex = 0;
        btnSelect1.Text = "選擇報表1檔案";
        btnSelect1.UseVisualStyleBackColor = false;
        btnSelect1.Click += btnSelect_Click;
        // 
        // txtBoxSelect1
        // 
        txtBoxSelect1.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxSelect1.Location = new Point(295, 31);
        txtBoxSelect1.Name = "txtBoxSelect1";
        txtBoxSelect1.ReadOnly = true;
        txtBoxSelect1.Size = new Size(114, 33);
        txtBoxSelect1.TabIndex = 13;
        txtBoxSelect1.TextAlign = HorizontalAlignment.Center;
        // 
        // btnSelect2
        // 
        btnSelect2.BackColor = SystemColors.ActiveCaption;
        btnSelect2.Location = new Point(19, 84);
        btnSelect2.Margin = new Padding(5);
        btnSelect2.Name = "btnSelect2";
        btnSelect2.Size = new Size(157, 38);
        btnSelect2.TabIndex = 14;
        btnSelect2.Text = "選擇報表2檔案";
        btnSelect2.UseVisualStyleBackColor = false;
        btnSelect2.Click += btnSelect2_Click;
        // 
        // txtBoxSelect2
        // 
        txtBoxSelect2.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxSelect2.Location = new Point(295, 84);
        txtBoxSelect2.Name = "txtBoxSelect2";
        txtBoxSelect2.ReadOnly = true;
        txtBoxSelect2.Size = new Size(114, 33);
        txtBoxSelect2.TabIndex = 17;
        txtBoxSelect2.TextAlign = HorizontalAlignment.Center;
        // 
        // btnSelect3
        // 
        btnSelect3.BackColor = SystemColors.ActiveCaption;
        btnSelect3.Location = new Point(19, 138);
        btnSelect3.Margin = new Padding(5);
        btnSelect3.Name = "btnSelect3";
        btnSelect3.Size = new Size(157, 38);
        btnSelect3.TabIndex = 18;
        btnSelect3.Text = "選擇報表3檔案";
        btnSelect3.UseVisualStyleBackColor = false;
        btnSelect3.Click += btnSelect3_Click;
        // 
        // txtBoxSelect3
        // 
        txtBoxSelect3.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxSelect3.Location = new Point(295, 136);
        txtBoxSelect3.Name = "txtBoxSelect3";
        txtBoxSelect3.ReadOnly = true;
        txtBoxSelect3.Size = new Size(114, 33);
        txtBoxSelect3.TabIndex = 21;
        txtBoxSelect3.TextAlign = HorizontalAlignment.Center;
        // 
        // label4
        // 
        label4.AutoSize = true;
        label4.Location = new Point(185, 38);
        label4.Margin = new Padding(5, 0, 5, 0);
        label4.Name = "label4";
        label4.Size = new Size(96, 23);
        label4.TabIndex = 29;
        label4.Text = "報表1類別:";
        // 
        // label7
        // 
        label7.AutoSize = true;
        label7.Location = new Point(185, 143);
        label7.Margin = new Padding(5, 0, 5, 0);
        label7.Name = "label7";
        label7.Size = new Size(96, 23);
        label7.TabIndex = 30;
        label7.Text = "報表3類別:";
        // 
        // label8
        // 
        label8.AutoSize = true;
        label8.Location = new Point(185, 90);
        label8.Margin = new Padding(5, 0, 5, 0);
        label8.Name = "label8";
        label8.Size = new Size(96, 23);
        label8.TabIndex = 31;
        label8.Text = "報表2類別:";
        // 
        // btnSelect4
        // 
        btnSelect4.BackColor = SystemColors.ActiveCaption;
        btnSelect4.Location = new Point(19, 196);
        btnSelect4.Margin = new Padding(5);
        btnSelect4.Name = "btnSelect4";
        btnSelect4.Size = new Size(157, 38);
        btnSelect4.TabIndex = 32;
        btnSelect4.Text = "選擇報表4檔案";
        btnSelect4.UseVisualStyleBackColor = false;
        btnSelect4.Click += btnSelect4_Click;
        // 
        // txtBoxSelect4
        // 
        txtBoxSelect4.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxSelect4.Location = new Point(295, 195);
        txtBoxSelect4.Name = "txtBoxSelect4";
        txtBoxSelect4.ReadOnly = true;
        txtBoxSelect4.Size = new Size(114, 33);
        txtBoxSelect4.TabIndex = 33;
        txtBoxSelect4.TextAlign = HorizontalAlignment.Center;
        // 
        // label9
        // 
        label9.AutoSize = true;
        label9.Location = new Point(185, 201);
        label9.Margin = new Padding(5, 0, 5, 0);
        label9.Name = "label9";
        label9.Size = new Size(96, 23);
        label9.TabIndex = 35;
        label9.Text = "報表4類別:";
        // 
        // tabControl2
        // 
        tabControl2.Controls.Add(tabPage5);
        tabControl2.Controls.Add(tabPage6);
        tabControl2.Location = new Point(701, 5);
        tabControl2.Margin = new Padding(5);
        tabControl2.Name = "tabControl2";
        tabControl2.SelectedIndex = 0;
        tabControl2.Size = new Size(858, 238);
        tabControl2.TabIndex = 38;
        // 
        // tabPage5
        // 
        tabPage5.Controls.Add(cbIgnoreNoData);
        tabPage5.Controls.Add(cbIgnoreElementarySchool);
        tabPage5.Controls.Add(ckbCompare);
        tabPage5.Controls.Add(label6);
        tabPage5.Controls.Add(label5);
        tabPage5.Controls.Add(txtbIgnoreLevel);
        tabPage5.Controls.Add(label3);
        tabPage5.Controls.Add(txtBoxStable);
        tabPage5.Controls.Add(label2);
        tabPage5.Controls.Add(label1);
        tabPage5.Controls.Add(txtBoxStartColumn);
        tabPage5.Location = new Point(4, 32);
        tabPage5.Margin = new Padding(5);
        tabPage5.Name = "tabPage5";
        tabPage5.Padding = new Padding(5);
        tabPage5.Size = new Size(850, 202);
        tabPage5.TabIndex = 0;
        tabPage5.Text = "參數";
        tabPage5.UseVisualStyleBackColor = true;
        // 
        // cbIgnoreNoData
        // 
        cbIgnoreNoData.AutoSize = true;
        cbIgnoreNoData.Checked = true;
        cbIgnoreNoData.CheckState = CheckState.Checked;
        cbIgnoreNoData.Location = new Point(267, 52);
        cbIgnoreNoData.Margin = new Padding(5);
        cbIgnoreNoData.Name = "cbIgnoreNoData";
        cbIgnoreNoData.Size = new Size(162, 27);
        cbIgnoreNoData.TabIndex = 26;
        cbIgnoreNoData.Text = "忽略無紀錄日期";
        cbIgnoreNoData.UseVisualStyleBackColor = true;
        // 
        // cbIgnoreElementarySchool
        // 
        cbIgnoreElementarySchool.AutoSize = true;
        cbIgnoreElementarySchool.Checked = true;
        cbIgnoreElementarySchool.CheckState = CheckState.Checked;
        cbIgnoreElementarySchool.Location = new Point(267, 15);
        cbIgnoreElementarySchool.Margin = new Padding(5);
        cbIgnoreElementarySchool.Name = "cbIgnoreElementarySchool";
        cbIgnoreElementarySchool.Size = new Size(198, 27);
        cbIgnoreElementarySchool.TabIndex = 25;
        cbIgnoreElementarySchool.Text = "小學未受浸納入總計";
        cbIgnoreElementarySchool.UseVisualStyleBackColor = true;
        // 
        // ckbCompare
        // 
        ckbCompare.AutoSize = true;
        ckbCompare.Checked = true;
        ckbCompare.CheckState = CheckState.Checked;
        ckbCompare.Location = new Point(267, 98);
        ckbCompare.Margin = new Padding(5);
        ckbCompare.Name = "ckbCompare";
        ckbCompare.Size = new Size(180, 27);
        ckbCompare.TabIndex = 24;
        ckbCompare.Text = "比較前後統計週期";
        ckbCompare.UseVisualStyleBackColor = true;
        // 
        // label6
        // 
        label6.AutoSize = true;
        label6.Location = new Point(159, 67);
        label6.Name = "label6";
        label6.Size = new Size(67, 23);
        label6.TabIndex = 23;
        label6.Text = "% 剔除";
        // 
        // label5
        // 
        label5.AutoSize = true;
        label5.Location = new Point(212, 18);
        label5.Name = "label5";
        label5.Size = new Size(26, 23);
        label5.TabIndex = 22;
        label5.Text = "%";
        // 
        // txtbIgnoreLevel
        // 
        txtbIgnoreLevel.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtbIgnoreLevel.Location = new Point(115, 64);
        txtbIgnoreLevel.Name = "txtbIgnoreLevel";
        txtbIgnoreLevel.Size = new Size(46, 30);
        txtbIgnoreLevel.TabIndex = 21;
        txtbIgnoreLevel.Text = "40";
        txtbIgnoreLevel.TextAlign = HorizontalAlignment.Center;
        // 
        // label3
        // 
        label3.AutoSize = true;
        label3.Location = new Point(8, 67);
        label3.Name = "label3";
        label3.Size = new Size(100, 23);
        label3.TabIndex = 20;
        label3.Text = "低於中位數";
        // 
        // txtBoxStable
        // 
        txtBoxStable.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxStable.Location = new Point(162, 15);
        txtBoxStable.Name = "txtBoxStable";
        txtBoxStable.Size = new Size(43, 30);
        txtBoxStable.TabIndex = 19;
        txtBoxStable.Text = "40";
        txtBoxStable.TextAlign = HorizontalAlignment.Center;
        // 
        // label2
        // 
        label2.AutoSize = true;
        label2.Location = new Point(8, 18);
        label2.Name = "label2";
        label2.Size = new Size(140, 23);
        label2.TabIndex = 18;
        label2.Text = "穩定聚會出席率:";
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(8, 107);
        label1.Name = "label1";
        label1.Size = new Size(158, 23);
        label1.TabIndex = 16;
        label1.Text = "出席紀錄開始欄位:";
        // 
        // txtBoxStartColumn
        // 
        txtBoxStartColumn.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxStartColumn.Location = new Point(181, 101);
        txtBoxStartColumn.Name = "txtBoxStartColumn";
        txtBoxStartColumn.Size = new Size(62, 33);
        txtBoxStartColumn.TabIndex = 17;
        txtBoxStartColumn.Text = "I";
        txtBoxStartColumn.TextAlign = HorizontalAlignment.Center;
        // 
        // tabPage6
        // 
        tabPage6.Controls.Add(groupBox2);
        tabPage6.Controls.Add(groupBox3);
        tabPage6.Location = new Point(4, 32);
        tabPage6.Margin = new Padding(5);
        tabPage6.Name = "tabPage6";
        tabPage6.Padding = new Padding(5);
        tabPage6.Size = new Size(850, 202);
        tabPage6.TabIndex = 1;
        tabPage6.Text = "分類文字";
        tabPage6.UseVisualStyleBackColor = true;
        // 
        // groupBox2
        // 
        groupBox2.Controls.Add(tbSheet4Cat3);
        groupBox2.Controls.Add(tbSheet3Cat3);
        groupBox2.Controls.Add(tbSheet2Cat3);
        groupBox2.Controls.Add(tbSheet1Cat3);
        groupBox2.Controls.Add(label14);
        groupBox2.Controls.Add(tbSheet4Cat2);
        groupBox2.Controls.Add(tbSheet4Cat1);
        groupBox2.Controls.Add(label15);
        groupBox2.Controls.Add(tbSheet3Cat2);
        groupBox2.Controls.Add(tbSheet3Cat1);
        groupBox2.Controls.Add(label16);
        groupBox2.Controls.Add(tbSheet2Cat2);
        groupBox2.Controls.Add(tbSheet2Cat1);
        groupBox2.Controls.Add(label17);
        groupBox2.Controls.Add(tbSheet1Cat2);
        groupBox2.Controls.Add(tbSheet1Cat1);
        groupBox2.Location = new Point(347, 9);
        groupBox2.Margin = new Padding(5);
        groupBox2.Name = "groupBox2";
        groupBox2.Padding = new Padding(5);
        groupBox2.Size = new Size(482, 182);
        groupBox2.TabIndex = 40;
        groupBox2.TabStop = false;
        groupBox2.Text = "週以上";
        // 
        // tbSheet4Cat3
        // 
        tbSheet4Cat3.Location = new Point(347, 143);
        tbSheet4Cat3.Margin = new Padding(5);
        tbSheet4Cat3.Name = "tbSheet4Cat3";
        tbSheet4Cat3.Size = new Size(116, 30);
        tbSheet4Cat3.TabIndex = 43;
        tbSheet4Cat3.Text = "無紀錄";
        // 
        // tbSheet3Cat3
        // 
        tbSheet3Cat3.Location = new Point(347, 103);
        tbSheet3Cat3.Margin = new Padding(5);
        tbSheet3Cat3.Name = "tbSheet3Cat3";
        tbSheet3Cat3.Size = new Size(116, 30);
        tbSheet3Cat3.TabIndex = 42;
        tbSheet3Cat3.Text = "無紀錄";
        // 
        // tbSheet2Cat3
        // 
        tbSheet2Cat3.Location = new Point(347, 64);
        tbSheet2Cat3.Margin = new Padding(5);
        tbSheet2Cat3.Name = "tbSheet2Cat3";
        tbSheet2Cat3.Size = new Size(116, 30);
        tbSheet2Cat3.TabIndex = 41;
        tbSheet2Cat3.Text = "無紀錄";
        // 
        // tbSheet1Cat3
        // 
        tbSheet1Cat3.Location = new Point(347, 25);
        tbSheet1Cat3.Margin = new Padding(5);
        tbSheet1Cat3.Name = "tbSheet1Cat3";
        tbSheet1Cat3.Size = new Size(116, 30);
        tbSheet1Cat3.TabIndex = 40;
        tbSheet1Cat3.Text = "無紀錄";
        // 
        // label14
        // 
        label14.AutoSize = true;
        label14.Location = new Point(9, 149);
        label14.Margin = new Padding(5, 0, 5, 0);
        label14.Name = "label14";
        label14.Size = new Size(60, 23);
        label14.TabIndex = 39;
        label14.Text = "報表4:";
        // 
        // tbSheet4Cat2
        // 
        tbSheet4Cat2.Location = new Point(209, 143);
        tbSheet4Cat2.Margin = new Padding(5);
        tbSheet4Cat2.Name = "tbSheet4Cat2";
        tbSheet4Cat2.Size = new Size(127, 30);
        tbSheet4Cat2.TabIndex = 38;
        tbSheet4Cat2.Text = "不穩定";
        // 
        // tbSheet4Cat1
        // 
        tbSheet4Cat1.Location = new Point(83, 144);
        tbSheet4Cat1.Margin = new Padding(5);
        tbSheet4Cat1.Name = "tbSheet4Cat1";
        tbSheet4Cat1.Size = new Size(114, 30);
        tbSheet4Cat1.TabIndex = 37;
        tbSheet4Cat1.Text = "穩定";
        // 
        // label15
        // 
        label15.AutoSize = true;
        label15.Location = new Point(9, 109);
        label15.Margin = new Padding(5, 0, 5, 0);
        label15.Name = "label15";
        label15.Size = new Size(60, 23);
        label15.TabIndex = 36;
        label15.Text = "報表3:";
        // 
        // tbSheet3Cat2
        // 
        tbSheet3Cat2.Location = new Point(209, 103);
        tbSheet3Cat2.Margin = new Padding(5);
        tbSheet3Cat2.Name = "tbSheet3Cat2";
        tbSheet3Cat2.Size = new Size(127, 30);
        tbSheet3Cat2.TabIndex = 35;
        tbSheet3Cat2.Text = "不穩定";
        // 
        // tbSheet3Cat1
        // 
        tbSheet3Cat1.Location = new Point(83, 104);
        tbSheet3Cat1.Margin = new Padding(5);
        tbSheet3Cat1.Name = "tbSheet3Cat1";
        tbSheet3Cat1.Size = new Size(114, 30);
        tbSheet3Cat1.TabIndex = 34;
        tbSheet3Cat1.Text = "穩定聚會";
        // 
        // label16
        // 
        label16.AutoSize = true;
        label16.Location = new Point(9, 71);
        label16.Margin = new Padding(5, 0, 5, 0);
        label16.Name = "label16";
        label16.Size = new Size(60, 23);
        label16.TabIndex = 33;
        label16.Text = "報表2:";
        // 
        // tbSheet2Cat2
        // 
        tbSheet2Cat2.Location = new Point(209, 64);
        tbSheet2Cat2.Margin = new Padding(5);
        tbSheet2Cat2.Name = "tbSheet2Cat2";
        tbSheet2Cat2.Size = new Size(127, 30);
        tbSheet2Cat2.TabIndex = 32;
        tbSheet2Cat2.Text = "不穩定";
        // 
        // tbSheet2Cat1
        // 
        tbSheet2Cat1.Location = new Point(83, 66);
        tbSheet2Cat1.Margin = new Padding(5);
        tbSheet2Cat1.Name = "tbSheet2Cat1";
        tbSheet2Cat1.Size = new Size(114, 30);
        tbSheet2Cat1.TabIndex = 31;
        tbSheet2Cat1.Text = "穩定聚會";
        // 
        // label17
        // 
        label17.AutoSize = true;
        label17.Location = new Point(9, 31);
        label17.Margin = new Padding(5, 0, 5, 0);
        label17.Name = "label17";
        label17.Size = new Size(60, 23);
        label17.TabIndex = 30;
        label17.Text = "報表1:";
        // 
        // tbSheet1Cat2
        // 
        tbSheet1Cat2.Location = new Point(209, 25);
        tbSheet1Cat2.Margin = new Padding(5);
        tbSheet1Cat2.Name = "tbSheet1Cat2";
        tbSheet1Cat2.Size = new Size(127, 30);
        tbSheet1Cat2.TabIndex = 2;
        tbSheet1Cat2.Text = "不穩定";
        // 
        // tbSheet1Cat1
        // 
        tbSheet1Cat1.Location = new Point(83, 26);
        tbSheet1Cat1.Margin = new Padding(5);
        tbSheet1Cat1.Name = "tbSheet1Cat1";
        tbSheet1Cat1.Size = new Size(114, 30);
        tbSheet1Cat1.TabIndex = 0;
        tbSheet1Cat1.Text = "穩定聚會";
        // 
        // groupBox3
        // 
        groupBox3.Controls.Add(label13);
        groupBox3.Controls.Add(tbSheet4WeekCat2);
        groupBox3.Controls.Add(tbSheet4WeekCat1);
        groupBox3.Controls.Add(label12);
        groupBox3.Controls.Add(tbSheet3WeekCat2);
        groupBox3.Controls.Add(tbSheet3WeekCat1);
        groupBox3.Controls.Add(label11);
        groupBox3.Controls.Add(tbSheet2WeekCat2);
        groupBox3.Controls.Add(tbSheet2WeekCat1);
        groupBox3.Controls.Add(label10);
        groupBox3.Controls.Add(tbSheet1WeekCat2);
        groupBox3.Controls.Add(tbSheet1WeekCat1);
        groupBox3.Location = new Point(9, 8);
        groupBox3.Margin = new Padding(5);
        groupBox3.Name = "groupBox3";
        groupBox3.Padding = new Padding(5);
        groupBox3.Size = new Size(328, 182);
        groupBox3.TabIndex = 38;
        groupBox3.TabStop = false;
        groupBox3.Text = "週";
        // 
        // label13
        // 
        label13.AutoSize = true;
        label13.Location = new Point(9, 149);
        label13.Margin = new Padding(5, 0, 5, 0);
        label13.Name = "label13";
        label13.Size = new Size(60, 23);
        label13.TabIndex = 39;
        label13.Text = "報表4:";
        // 
        // tbSheet4WeekCat2
        // 
        tbSheet4WeekCat2.Location = new Point(195, 144);
        tbSheet4WeekCat2.Margin = new Padding(5);
        tbSheet4WeekCat2.Name = "tbSheet4WeekCat2";
        tbSheet4WeekCat2.Size = new Size(92, 30);
        tbSheet4WeekCat2.TabIndex = 38;
        tbSheet4WeekCat2.Text = "無紀錄";
        // 
        // tbSheet4WeekCat1
        // 
        tbSheet4WeekCat1.Location = new Point(83, 144);
        tbSheet4WeekCat1.Margin = new Padding(5);
        tbSheet4WeekCat1.Name = "tbSheet4WeekCat1";
        tbSheet4WeekCat1.Size = new Size(100, 30);
        tbSheet4WeekCat1.TabIndex = 37;
        tbSheet4WeekCat1.Text = "本週有紀錄";
        // 
        // label12
        // 
        label12.AutoSize = true;
        label12.Location = new Point(9, 109);
        label12.Margin = new Padding(5, 0, 5, 0);
        label12.Name = "label12";
        label12.Size = new Size(60, 23);
        label12.TabIndex = 36;
        label12.Text = "報表3:";
        // 
        // tbSheet3WeekCat2
        // 
        tbSheet3WeekCat2.Location = new Point(195, 104);
        tbSheet3WeekCat2.Margin = new Padding(5);
        tbSheet3WeekCat2.Name = "tbSheet3WeekCat2";
        tbSheet3WeekCat2.Size = new Size(92, 30);
        tbSheet3WeekCat2.TabIndex = 35;
        tbSheet3WeekCat2.Text = "未到會";
        // 
        // tbSheet3WeekCat1
        // 
        tbSheet3WeekCat1.Location = new Point(83, 104);
        tbSheet3WeekCat1.Margin = new Padding(5);
        tbSheet3WeekCat1.Name = "tbSheet3WeekCat1";
        tbSheet3WeekCat1.Size = new Size(100, 30);
        tbSheet3WeekCat1.TabIndex = 34;
        tbSheet3WeekCat1.Text = "本週到會";
        // 
        // label11
        // 
        label11.AutoSize = true;
        label11.Location = new Point(9, 71);
        label11.Margin = new Padding(5, 0, 5, 0);
        label11.Name = "label11";
        label11.Size = new Size(60, 23);
        label11.TabIndex = 33;
        label11.Text = "報表2:";
        // 
        // tbSheet2WeekCat2
        // 
        tbSheet2WeekCat2.Location = new Point(195, 66);
        tbSheet2WeekCat2.Margin = new Padding(5);
        tbSheet2WeekCat2.Name = "tbSheet2WeekCat2";
        tbSheet2WeekCat2.Size = new Size(92, 30);
        tbSheet2WeekCat2.TabIndex = 32;
        tbSheet2WeekCat2.Text = "未到會";
        // 
        // tbSheet2WeekCat1
        // 
        tbSheet2WeekCat1.Location = new Point(83, 66);
        tbSheet2WeekCat1.Margin = new Padding(5);
        tbSheet2WeekCat1.Name = "tbSheet2WeekCat1";
        tbSheet2WeekCat1.Size = new Size(100, 30);
        tbSheet2WeekCat1.TabIndex = 31;
        tbSheet2WeekCat1.Text = "本週到會";
        // 
        // label10
        // 
        label10.AutoSize = true;
        label10.Location = new Point(9, 31);
        label10.Margin = new Padding(5, 0, 5, 0);
        label10.Name = "label10";
        label10.Size = new Size(60, 23);
        label10.TabIndex = 30;
        label10.Text = "報表1:";
        // 
        // tbSheet1WeekCat2
        // 
        tbSheet1WeekCat2.Location = new Point(195, 26);
        tbSheet1WeekCat2.Margin = new Padding(5);
        tbSheet1WeekCat2.Name = "tbSheet1WeekCat2";
        tbSheet1WeekCat2.Size = new Size(92, 30);
        tbSheet1WeekCat2.TabIndex = 2;
        tbSheet1WeekCat2.Text = "未到會";
        // 
        // tbSheet1WeekCat1
        // 
        tbSheet1WeekCat1.Location = new Point(83, 26);
        tbSheet1WeekCat1.Margin = new Padding(5);
        tbSheet1WeekCat1.Name = "tbSheet1WeekCat1";
        tbSheet1WeekCat1.Size = new Size(100, 30);
        tbSheet1WeekCat1.TabIndex = 0;
        tbSheet1WeekCat1.Text = "本週到會";
        // 
        // splitContainer1
        // 
        splitContainer1.Location = new Point(252, 80);
        splitContainer1.Name = "splitContainer1";
        splitContainer1.Size = new Size(225, 150);
        splitContainer1.SplitterDistance = 75;
        splitContainer1.TabIndex = 39;
        // 
        // AttendForm
        // 
        AutoScaleDimensions = new SizeF(11F, 23F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(1578, 1017);
        Controls.Add(splitContainer1);
        Controls.Add(tabControl2);
        Controls.Add(btnCalculateAllExcel);
        Controls.Add(label9);
        Controls.Add(txtBoxSelect4);
        Controls.Add(btnSelect4);
        Controls.Add(label8);
        Controls.Add(label7);
        Controls.Add(label4);
        Controls.Add(tabControl1);
        Controls.Add(txtBoxSelect3);
        Controls.Add(btnSelect3);
        Controls.Add(txtBoxSelect2);
        Controls.Add(btnSelect2);
        Controls.Add(txtBoxSelect1);
        Controls.Add(groupBox1);
        Controls.Add(btnSelect1);
        Margin = new Padding(5);
        Name = "AttendForm";
        Text = "點名系統表單整理小幫手";
        FormClosing += AttendForm_FormClosing;
        Load += AttendForm_Load;
        SizeChanged += AttendForm_SizeChanged;
        groupBox1.ResumeLayout(false);
        groupBox1.PerformLayout();
        tabPage3.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResult3).EndInit();
        tabPage2.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResult2).EndInit();
        tabPage1.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResult1).EndInit();
        tabControl1.ResumeLayout(false);
        tabPage4.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResult4).EndInit();
        tabControl2.ResumeLayout(false);
        tabPage5.ResumeLayout(false);
        tabPage5.PerformLayout();
        tabPage6.ResumeLayout(false);
        groupBox2.ResumeLayout(false);
        groupBox2.PerformLayout();
        groupBox3.ResumeLayout(false);
        groupBox3.PerformLayout();
        ((System.ComponentModel.ISupportInitialize)splitContainer1).EndInit();
        splitContainer1.ResumeLayout(false);
        ResumeLayout(false);
        PerformLayout();
    }

    private void btnCalculate_Click(object sender, EventArgs e)
    {
        OpenExcelFile(filenames[0], txtBoxSelect1.Text + ".xlsx", dgvResult1);
        MessageBox.Show("Finished!");
    }

#endregion
    private GroupBox groupBox1;
    private RadioButton rbWeek;
    private RadioButton rbHalfYear;
    private RadioButton rbMonth;
    private TabPage tabPage3;
    private DataGridView dgvResult3;
    private TabPage tabPage2;
    private DataGridView dgvResult2;
    private TabPage tabPage1;
    private DataGridView dgvResult1;
    private TabControl tabControl1;
    private TabPage tabPage4;
    private DataGridView dgvResult4;
    private Button btnCalculateAllExcel;
    private Button btnSelect1;
    private TextBox txtBoxSelect1;
    private Button btnSelect2;
    private TextBox txtBoxSelect2;
    private Button btnSelect3;
    private TextBox txtBoxSelect3;
    private Label label4;
    private Label label7;
    private Label label8;
    private Button btnSelect4;
    private TextBox txtBoxSelect4;
    private Label label9;
    private TabControl tabControl2;
    private TabPage tabPage5;
    private CheckBox cbIgnoreNoData;
    private CheckBox cbIgnoreElementarySchool;
    private CheckBox ckbCompare;
    private Label label6;
    private Label label5;
    private TextBox txtbIgnoreLevel;
    private Label label3;
    private TextBox txtBoxStable;
    private Label label2;
    private Label label1;
    private TextBox txtBoxStartColumn;
    private TabPage tabPage6;
    private GroupBox groupBox3;
    private TextBox tbSheet1WeekCat2;
    private TextBox tbSheet1WeekCat1;
    private GroupBox groupBox2;
    private TextBox tbSheet4Cat3;
    private TextBox tbSheet3Cat3;
    private TextBox tbSheet2Cat3;
    private TextBox tbSheet1Cat3;
    private Label label14;
    private TextBox tbSheet4Cat2;
    private TextBox tbSheet4Cat1;
    private Label label15;
    private TextBox tbSheet3Cat2;
    private TextBox tbSheet3Cat1;
    private Label label16;
    private TextBox tbSheet2Cat2;
    private TextBox tbSheet2Cat1;
    private Label label17;
    private TextBox tbSheet1Cat2;
    private TextBox tbSheet1Cat1;
    private Label label13;
    private TextBox tbSheet4WeekCat2;
    private TextBox tbSheet4WeekCat1;
    private Label label12;
    private TextBox tbSheet3WeekCat2;
    private TextBox tbSheet3WeekCat1;
    private Label label11;
    private TextBox tbSheet2WeekCat2;
    private TextBox tbSheet2WeekCat1;
    private Label label10;
    private SplitContainer splitContainer1;
}
