
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
        folderBrowserDialog1 = new FolderBrowserDialog();
        btnSelect1 = new Button();
        btnCalculateExcel1 = new Button();
        label1 = new Label();
        txtBoxStartColumn = new TextBox();
        groupBox1 = new GroupBox();
        rbHalfYear = new RadioButton();
        rbMonth = new RadioButton();
        rbWeek = new RadioButton();
        groupBox2 = new GroupBox();
        ckbCompare = new CheckBox();
        label6 = new Label();
        label5 = new Label();
        txtbIgnoreLevel = new TextBox();
        label3 = new Label();
        txtBoxStable = new TextBox();
        label2 = new Label();
        txtBoxOutput = new TextBox();
        label4 = new Label();
        txtBoxLord = new TextBox();
        txtBoxOutputPray = new TextBox();
        txtBoxPray = new TextBox();
        label7 = new Label();
        btnSelect2 = new Button();
        txtBoxOutputHome = new TextBox();
        txtBoxHome = new TextBox();
        label8 = new Label();
        btnSelect3 = new Button();
        groupBox3 = new GroupBox();
        dgvResult1 = new DataGridView();
        groupBox4 = new GroupBox();
        dgvResult2 = new DataGridView();
        groupBox5 = new GroupBox();
        dgvResult3 = new DataGridView();
        btnCalculateExcel2 = new Button();
        btnCalculateExcel3 = new Button();
        groupBox1.SuspendLayout();
        groupBox2.SuspendLayout();
        groupBox3.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult1).BeginInit();
        groupBox4.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult2).BeginInit();
        groupBox5.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult3).BeginInit();
        SuspendLayout();
        // 
        // btnSelect1
        // 
        btnSelect1.Location = new Point(12, 11);
        btnSelect1.Name = "btnSelect1";
        btnSelect1.Size = new Size(85, 25);
        btnSelect1.TabIndex = 0;
        btnSelect1.Text = "選擇表單1";
        btnSelect1.UseVisualStyleBackColor = true;
        btnSelect1.Click += btnSelect_Click;
        // 
        // btnCalculateExcel1
        // 
        btnCalculateExcel1.BackColor = SystemColors.ActiveCaption;
        btnCalculateExcel1.Location = new Point(318, 1);
        btnCalculateExcel1.Margin = new Padding(2);
        btnCalculateExcel1.Name = "btnCalculateExcel1";
        btnCalculateExcel1.Size = new Size(92, 44);
        btnCalculateExcel1.TabIndex = 3;
        btnCalculateExcel1.Text = "統計表單1";
        btnCalculateExcel1.UseVisualStyleBackColor = false;
        btnCalculateExcel1.Click += btnCalculate_Click;
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(15, 79);
        label1.Margin = new Padding(2, 0, 2, 0);
        label1.Name = "label1";
        label1.Size = new Size(106, 15);
        label1.TabIndex = 5;
        label1.Text = "出席紀錄開始欄位:";
        // 
        // txtBoxStartColumn
        // 
        txtBoxStartColumn.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxStartColumn.Location = new Point(125, 75);
        txtBoxStartColumn.Margin = new Padding(2);
        txtBoxStartColumn.Name = "txtBoxStartColumn";
        txtBoxStartColumn.Size = new Size(41, 24);
        txtBoxStartColumn.TabIndex = 6;
        txtBoxStartColumn.Text = "I";
        txtBoxStartColumn.TextAlign = HorizontalAlignment.Center;
        // 
        // groupBox1
        // 
        groupBox1.Controls.Add(rbHalfYear);
        groupBox1.Controls.Add(rbMonth);
        groupBox1.Controls.Add(rbWeek);
        groupBox1.Location = new Point(589, 15);
        groupBox1.Margin = new Padding(2);
        groupBox1.Name = "groupBox1";
        groupBox1.Padding = new Padding(2);
        groupBox1.Size = new Size(133, 127);
        groupBox1.TabIndex = 7;
        groupBox1.TabStop = false;
        groupBox1.Text = "統計單位";
        // 
        // rbHalfYear
        // 
        rbHalfYear.AutoSize = true;
        rbHalfYear.Location = new Point(16, 77);
        rbHalfYear.Name = "rbHalfYear";
        rbHalfYear.Size = new Size(83, 19);
        rbHalfYear.TabIndex = 11;
        rbHalfYear.Text = "半年(26週)";
        rbHalfYear.UseVisualStyleBackColor = true;
        // 
        // rbMonth
        // 
        rbMonth.AutoSize = true;
        rbMonth.Location = new Point(16, 52);
        rbMonth.Name = "rbMonth";
        rbMonth.Size = new Size(93, 19);
        rbMonth.TabIndex = 10;
        rbMonth.Text = "月(四到五週)";
        rbMonth.UseVisualStyleBackColor = true;
        // 
        // rbWeek
        // 
        rbWeek.AutoSize = true;
        rbWeek.Checked = true;
        rbWeek.Location = new Point(16, 27);
        rbWeek.Name = "rbWeek";
        rbWeek.Size = new Size(37, 19);
        rbWeek.TabIndex = 9;
        rbWeek.TabStop = true;
        rbWeek.Text = "週";
        rbWeek.UseVisualStyleBackColor = true;
        // 
        // groupBox2
        // 
        groupBox2.Controls.Add(ckbCompare);
        groupBox2.Controls.Add(label6);
        groupBox2.Controls.Add(label5);
        groupBox2.Controls.Add(txtbIgnoreLevel);
        groupBox2.Controls.Add(label3);
        groupBox2.Controls.Add(txtBoxStable);
        groupBox2.Controls.Add(label2);
        groupBox2.Controls.Add(label1);
        groupBox2.Controls.Add(txtBoxStartColumn);
        groupBox2.Location = new Point(727, 13);
        groupBox2.Name = "groupBox2";
        groupBox2.Size = new Size(209, 128);
        groupBox2.TabIndex = 8;
        groupBox2.TabStop = false;
        groupBox2.Text = "統計參數";
        // 
        // ckbCompare
        // 
        ckbCompare.AutoSize = true;
        ckbCompare.Location = new Point(15, 102);
        ckbCompare.Name = "ckbCompare";
        ckbCompare.Size = new Size(122, 19);
        ckbCompare.TabIndex = 13;
        ckbCompare.Text = "比較前後統計週期";
        ckbCompare.UseVisualStyleBackColor = true;
        // 
        // label6
        // 
        label6.AutoSize = true;
        label6.Location = new Point(111, 53);
        label6.Margin = new Padding(2, 0, 2, 0);
        label6.Name = "label6";
        label6.Size = new Size(45, 15);
        label6.TabIndex = 12;
        label6.Text = "% 剔除";
        // 
        // label5
        // 
        label5.AutoSize = true;
        label5.Location = new Point(145, 21);
        label5.Margin = new Padding(2, 0, 2, 0);
        label5.Name = "label5";
        label5.Size = new Size(18, 15);
        label5.TabIndex = 11;
        label5.Text = "%";
        // 
        // txtbIgnoreLevel
        // 
        txtbIgnoreLevel.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtbIgnoreLevel.Location = new Point(83, 51);
        txtbIgnoreLevel.Margin = new Padding(2);
        txtbIgnoreLevel.Name = "txtbIgnoreLevel";
        txtbIgnoreLevel.Size = new Size(31, 23);
        txtbIgnoreLevel.TabIndex = 10;
        txtbIgnoreLevel.Text = "40";
        txtbIgnoreLevel.TextAlign = HorizontalAlignment.Center;
        // 
        // label3
        // 
        label3.AutoSize = true;
        label3.Location = new Point(15, 53);
        label3.Margin = new Padding(2, 0, 2, 0);
        label3.Name = "label3";
        label3.Size = new Size(67, 15);
        label3.TabIndex = 9;
        label3.Text = "低於中位數";
        // 
        // txtBoxStable
        // 
        txtBoxStable.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxStable.Location = new Point(113, 19);
        txtBoxStable.Margin = new Padding(2);
        txtBoxStable.Name = "txtBoxStable";
        txtBoxStable.Size = new Size(29, 23);
        txtBoxStable.TabIndex = 8;
        txtBoxStable.Text = "40";
        txtBoxStable.TextAlign = HorizontalAlignment.Center;
        // 
        // label2
        // 
        label2.AutoSize = true;
        label2.Location = new Point(15, 21);
        label2.Margin = new Padding(2, 0, 2, 0);
        label2.Name = "label2";
        label2.Size = new Size(94, 15);
        label2.TabIndex = 7;
        label2.Text = "穩定聚會出席率:";
        // 
        // txtBoxOutput
        // 
        txtBoxOutput.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxOutput.Location = new Point(464, 20);
        txtBoxOutput.Margin = new Padding(2);
        txtBoxOutput.Name = "txtBoxOutput";
        txtBoxOutput.Size = new Size(115, 23);
        txtBoxOutput.TabIndex = 12;
        txtBoxOutput.Text = "主日.xlsx";
        // 
        // label4
        // 
        label4.AutoSize = true;
        label4.Location = new Point(414, 23);
        label4.Margin = new Padding(2, 0, 2, 0);
        label4.Name = "label4";
        label4.Size = new Size(46, 15);
        label4.TabIndex = 11;
        label4.Text = "輸出檔:";
        // 
        // txtBoxLord
        // 
        txtBoxLord.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxLord.Location = new Point(102, 12);
        txtBoxLord.Margin = new Padding(2);
        txtBoxLord.Name = "txtBoxLord";
        txtBoxLord.Size = new Size(205, 24);
        txtBoxLord.TabIndex = 13;
        txtBoxLord.TextAlign = HorizontalAlignment.Right;
        // 
        // txtBoxOutputPray
        // 
        txtBoxOutputPray.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxOutputPray.Location = new Point(464, 67);
        txtBoxOutputPray.Margin = new Padding(2);
        txtBoxOutputPray.Name = "txtBoxOutputPray";
        txtBoxOutputPray.Size = new Size(115, 23);
        txtBoxOutputPray.TabIndex = 16;
        txtBoxOutputPray.Text = "禱告聚會.xlsx";
        // 
        // txtBoxPray
        // 
        txtBoxPray.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxPray.Location = new Point(102, 58);
        txtBoxPray.Margin = new Padding(2);
        txtBoxPray.Name = "txtBoxPray";
        txtBoxPray.Size = new Size(205, 24);
        txtBoxPray.TabIndex = 17;
        txtBoxPray.TextAlign = HorizontalAlignment.Right;
        // 
        // label7
        // 
        label7.AutoSize = true;
        label7.Location = new Point(414, 70);
        label7.Margin = new Padding(2, 0, 2, 0);
        label7.Name = "label7";
        label7.Size = new Size(46, 15);
        label7.TabIndex = 15;
        label7.Text = "輸出檔:";
        // 
        // btnSelect2
        // 
        btnSelect2.Location = new Point(12, 56);
        btnSelect2.Name = "btnSelect2";
        btnSelect2.Size = new Size(85, 25);
        btnSelect2.TabIndex = 14;
        btnSelect2.Text = "選擇表單2";
        btnSelect2.UseVisualStyleBackColor = true;
        btnSelect2.Click += btnSelect2_Click;
        // 
        // txtBoxOutputHome
        // 
        txtBoxOutputHome.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxOutputHome.Location = new Point(464, 112);
        txtBoxOutputHome.Margin = new Padding(2);
        txtBoxOutputHome.Name = "txtBoxOutputHome";
        txtBoxOutputHome.Size = new Size(115, 23);
        txtBoxOutputHome.TabIndex = 20;
        txtBoxOutputHome.Text = "小排.xlsx";
        // 
        // txtBoxHome
        // 
        txtBoxHome.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxHome.Location = new Point(102, 105);
        txtBoxHome.Margin = new Padding(2);
        txtBoxHome.Name = "txtBoxHome";
        txtBoxHome.Size = new Size(205, 24);
        txtBoxHome.TabIndex = 21;
        txtBoxHome.TextAlign = HorizontalAlignment.Right;
        // 
        // label8
        // 
        label8.AutoSize = true;
        label8.Location = new Point(414, 115);
        label8.Margin = new Padding(2, 0, 2, 0);
        label8.Name = "label8";
        label8.Size = new Size(46, 15);
        label8.TabIndex = 19;
        label8.Text = "輸出檔:";
        // 
        // btnSelect3
        // 
        btnSelect3.Location = new Point(12, 104);
        btnSelect3.Name = "btnSelect3";
        btnSelect3.Size = new Size(85, 25);
        btnSelect3.TabIndex = 18;
        btnSelect3.Text = "選擇表單3";
        btnSelect3.UseVisualStyleBackColor = true;
        btnSelect3.Click += btnSelect3_Click;
        // 
        // groupBox3
        // 
        groupBox3.Controls.Add(dgvResult1);
        groupBox3.Location = new Point(13, 154);
        groupBox3.Name = "groupBox3";
        groupBox3.Size = new Size(294, 423);
        groupBox3.TabIndex = 22;
        groupBox3.TabStop = false;
        groupBox3.Text = "表單一最新週期統計結果";
        // 
        // dgvResult1
        // 
        dgvResult1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResult1.Location = new Point(16, 22);
        dgvResult1.Name = "dgvResult1";
        dgvResult1.Size = new Size(258, 395);
        dgvResult1.TabIndex = 0;
        // 
        // groupBox4
        // 
        groupBox4.Controls.Add(dgvResult2);
        groupBox4.Location = new Point(318, 154);
        groupBox4.Name = "groupBox4";
        groupBox4.Size = new Size(294, 423);
        groupBox4.TabIndex = 23;
        groupBox4.TabStop = false;
        groupBox4.Text = "表單二最新週期統計結果";
        // 
        // dgvResult2
        // 
        dgvResult2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResult2.Location = new Point(17, 21);
        dgvResult2.Name = "dgvResult2";
        dgvResult2.Size = new Size(258, 390);
        dgvResult2.TabIndex = 1;
        // 
        // groupBox5
        // 
        groupBox5.Controls.Add(dgvResult3);
        groupBox5.FlatStyle = FlatStyle.Flat;
        groupBox5.Location = new Point(630, 154);
        groupBox5.Name = "groupBox5";
        groupBox5.Size = new Size(306, 423);
        groupBox5.TabIndex = 23;
        groupBox5.TabStop = false;
        groupBox5.Text = "表單三最新週期統計結果";
        // 
        // dgvResult3
        // 
        dgvResult3.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResult3.Location = new Point(18, 22);
        dgvResult3.Name = "dgvResult3";
        dgvResult3.Size = new Size(258, 389);
        dgvResult3.TabIndex = 2;
        // 
        // btnCalculateExcel2
        // 
        btnCalculateExcel2.BackColor = SystemColors.ActiveCaption;
        btnCalculateExcel2.Location = new Point(318, 49);
        btnCalculateExcel2.Margin = new Padding(2);
        btnCalculateExcel2.Name = "btnCalculateExcel2";
        btnCalculateExcel2.Size = new Size(92, 41);
        btnCalculateExcel2.TabIndex = 24;
        btnCalculateExcel2.Text = "統計表單2";
        btnCalculateExcel2.UseVisualStyleBackColor = false;
        btnCalculateExcel2.Click += btnCalculateSheet2_Click;
        // 
        // btnCalculateExcel3
        // 
        btnCalculateExcel3.BackColor = SystemColors.ActiveCaption;
        btnCalculateExcel3.Location = new Point(318, 94);
        btnCalculateExcel3.Margin = new Padding(2);
        btnCalculateExcel3.Name = "btnCalculateExcel3";
        btnCalculateExcel3.Size = new Size(92, 45);
        btnCalculateExcel3.TabIndex = 25;
        btnCalculateExcel3.Text = "統計表單3";
        btnCalculateExcel3.UseVisualStyleBackColor = false;
        btnCalculateExcel3.Click += btnCalculateSheet3_Click;
        // 
        // AttendForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(954, 589);
        Controls.Add(btnCalculateExcel3);
        Controls.Add(btnCalculateExcel2);
        Controls.Add(groupBox5);
        Controls.Add(groupBox4);
        Controls.Add(groupBox3);
        Controls.Add(txtBoxOutputHome);
        Controls.Add(txtBoxHome);
        Controls.Add(label8);
        Controls.Add(btnSelect3);
        Controls.Add(txtBoxOutputPray);
        Controls.Add(txtBoxPray);
        Controls.Add(label7);
        Controls.Add(btnSelect2);
        Controls.Add(txtBoxOutput);
        Controls.Add(txtBoxLord);
        Controls.Add(label4);
        Controls.Add(groupBox2);
        Controls.Add(groupBox1);
        Controls.Add(btnCalculateExcel1);
        Controls.Add(btnSelect1);
        Name = "AttendForm";
        Text = "Form1";
        groupBox1.ResumeLayout(false);
        groupBox1.PerformLayout();
        groupBox2.ResumeLayout(false);
        groupBox2.PerformLayout();
        groupBox3.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResult1).EndInit();
        groupBox4.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResult2).EndInit();
        groupBox5.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResult3).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    private void btnCalculate_Click(object sender, EventArgs e)
    {
        OpenExcelFile(txtBoxLord.Text, txtBoxOutput.Text, dgvResult1);
    }

    #endregion
    private FolderBrowserDialog folderBrowserDialog1;
    private Button btnSelect1;
    private Button btnCalculateExcel1;
    private Label label1;
    private TextBox txtBoxStartColumn;
    private GroupBox groupBox1;
    private GroupBox groupBox2;
    private TextBox txtBoxStable;
    private Label label2;
    private RadioButton rbWeek;
    private RadioButton rbHalfYear;
    private RadioButton rbMonth;
    private TextBox txtbIgnoreLevel;
    private Label label3;
    private TextBox txtBoxOutput;
    private Label label4;
    private TextBox txtBoxLord;
    private Label label6;
    private Label label5;
    private TextBox txtBoxOutputPray;
    private TextBox txtBoxPray;
    private Label label7;
    private Button btnSelect2;
    private TextBox txtBoxOutputHome;
    private TextBox txtBoxHome;
    private Label label8;
    private Button btnSelect3;
    private CheckBox ckbCompare;
    private GroupBox groupBox3;
    private GroupBox groupBox4;
    private GroupBox groupBox5;
    private Button btnCalculateExcel2;
    private Button btnCalculateExcel3;
    private DataGridView dgvResult1;
    private DataGridView dgvResult2;
    private DataGridView dgvResult3;
}
