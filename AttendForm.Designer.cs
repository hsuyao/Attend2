
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
        btnSelect1 = new Button();
        label1 = new Label();
        txtBoxStartColumn = new TextBox();
        groupBox1 = new GroupBox();
        rbHalfYear = new RadioButton();
        rbMonth = new RadioButton();
        rbWeek = new RadioButton();
        groupBox2 = new GroupBox();
        checkBox1 = new CheckBox();
        ckbCompare = new CheckBox();
        label6 = new Label();
        label5 = new Label();
        txtbIgnoreLevel = new TextBox();
        label3 = new Label();
        txtBoxStable = new TextBox();
        label2 = new Label();
        txtBoxSelect1 = new TextBox();
        txtBoxSelect2 = new TextBox();
        btnSelect2 = new Button();
        txtBoxSelect3 = new TextBox();
        btnSelect3 = new Button();
        tabPage3 = new TabPage();
        dgvResult3 = new DataGridView();
        tabPage2 = new TabPage();
        dgvResult2 = new DataGridView();
        tabPage1 = new TabPage();
        dgvResult1 = new DataGridView();
        tabControl1 = new TabControl();
        tabPage4 = new TabPage();
        dgvResult4 = new DataGridView();
        label4 = new Label();
        label7 = new Label();
        label8 = new Label();
        label9 = new Label();
        txtBoxSelect4 = new TextBox();
        btnSelect4 = new Button();
        btnCalculateAllExcel = new Button();
        groupBox1.SuspendLayout();
        groupBox2.SuspendLayout();
        tabPage3.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult3).BeginInit();
        tabPage2.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult2).BeginInit();
        tabPage1.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult1).BeginInit();
        tabControl1.SuspendLayout();
        tabPage4.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult4).BeginInit();
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
        groupBox1.Location = new Point(489, 8);
        groupBox1.Margin = new Padding(2);
        groupBox1.Name = "groupBox1";
        groupBox1.Padding = new Padding(2);
        groupBox1.Size = new Size(118, 127);
        groupBox1.TabIndex = 7;
        groupBox1.TabStop = false;
        groupBox1.Text = "統計單位";
        // 
        // rbHalfYear
        // 
        rbHalfYear.AutoSize = true;
        rbHalfYear.Location = new Point(16, 89);
        rbHalfYear.Name = "rbHalfYear";
        rbHalfYear.Size = new Size(83, 19);
        rbHalfYear.TabIndex = 11;
        rbHalfYear.Text = "半年(26週)";
        rbHalfYear.UseVisualStyleBackColor = true;
        // 
        // rbMonth
        // 
        rbMonth.AutoSize = true;
        rbMonth.Location = new Point(16, 58);
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
        groupBox2.Controls.Add(checkBox1);
        groupBox2.Controls.Add(ckbCompare);
        groupBox2.Controls.Add(label6);
        groupBox2.Controls.Add(label5);
        groupBox2.Controls.Add(txtbIgnoreLevel);
        groupBox2.Controls.Add(label3);
        groupBox2.Controls.Add(txtBoxStable);
        groupBox2.Controls.Add(label2);
        groupBox2.Controls.Add(label1);
        groupBox2.Controls.Add(txtBoxStartColumn);
        groupBox2.Location = new Point(634, 11);
        groupBox2.Name = "groupBox2";
        groupBox2.Size = new Size(314, 128);
        groupBox2.TabIndex = 8;
        groupBox2.TabStop = false;
        groupBox2.Text = "統計參數";
        // 
        // checkBox1
        // 
        checkBox1.AutoSize = true;
        checkBox1.Checked = true;
        checkBox1.CheckState = CheckState.Checked;
        checkBox1.Location = new Point(180, 19);
        checkBox1.Name = "checkBox1";
        checkBox1.Size = new Size(134, 19);
        checkBox1.TabIndex = 14;
        checkBox1.Text = "小學未受浸納入總計";
        checkBox1.UseVisualStyleBackColor = true;
        // 
        // ckbCompare
        // 
        ckbCompare.AutoSize = true;
        ckbCompare.Checked = true;
        ckbCompare.CheckState = CheckState.Checked;
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
        // txtBoxSelect1
        // 
        txtBoxSelect1.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxSelect1.Location = new Point(168, 11);
        txtBoxSelect1.Margin = new Padding(2);
        txtBoxSelect1.Name = "txtBoxSelect1";
        txtBoxSelect1.Size = new Size(54, 24);
        txtBoxSelect1.TabIndex = 13;
        txtBoxSelect1.TextAlign = HorizontalAlignment.Right;
        // 
        // txtBoxSelect2
        // 
        txtBoxSelect2.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxSelect2.Location = new Point(168, 42);
        txtBoxSelect2.Margin = new Padding(2);
        txtBoxSelect2.Name = "txtBoxSelect2";
        txtBoxSelect2.Size = new Size(54, 24);
        txtBoxSelect2.TabIndex = 17;
        txtBoxSelect2.TextAlign = HorizontalAlignment.Right;
        // 
        // btnSelect2
        // 
        btnSelect2.Location = new Point(12, 41);
        btnSelect2.Name = "btnSelect2";
        btnSelect2.Size = new Size(85, 25);
        btnSelect2.TabIndex = 14;
        btnSelect2.Text = "選擇表單2";
        btnSelect2.UseVisualStyleBackColor = true;
        btnSelect2.Click += btnSelect2_Click;
        // 
        // txtBoxSelect3
        // 
        txtBoxSelect3.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxSelect3.Location = new Point(168, 74);
        txtBoxSelect3.Margin = new Padding(2);
        txtBoxSelect3.Name = "txtBoxSelect3";
        txtBoxSelect3.Size = new Size(54, 24);
        txtBoxSelect3.TabIndex = 21;
        txtBoxSelect3.TextAlign = HorizontalAlignment.Right;
        // 
        // btnSelect3
        // 
        btnSelect3.Location = new Point(12, 74);
        btnSelect3.Name = "btnSelect3";
        btnSelect3.Size = new Size(85, 25);
        btnSelect3.TabIndex = 18;
        btnSelect3.Text = "選擇表單3";
        btnSelect3.UseVisualStyleBackColor = true;
        btnSelect3.Click += btnSelect3_Click;
        // 
        // tabPage3
        // 
        tabPage3.Controls.Add(dgvResult3);
        tabPage3.Location = new Point(4, 24);
        tabPage3.Margin = new Padding(2);
        tabPage3.Name = "tabPage3";
        tabPage3.Padding = new Padding(2);
        tabPage3.Size = new Size(946, 415);
        tabPage3.TabIndex = 2;
        tabPage3.Text = "第三表單最新統計結果";
        tabPage3.UseVisualStyleBackColor = true;
        // 
        // dgvResult3
        // 
        dgvResult3.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResult3.Dock = DockStyle.Fill;
        dgvResult3.Location = new Point(2, 2);
        dgvResult3.Name = "dgvResult3";
        dgvResult3.RowHeadersWidth = 62;
        dgvResult3.Size = new Size(942, 411);
        dgvResult3.TabIndex = 4;
        // 
        // tabPage2
        // 
        tabPage2.Controls.Add(dgvResult2);
        tabPage2.Location = new Point(4, 24);
        tabPage2.Margin = new Padding(2);
        tabPage2.Name = "tabPage2";
        tabPage2.Padding = new Padding(2);
        tabPage2.Size = new Size(946, 415);
        tabPage2.TabIndex = 1;
        tabPage2.Text = "第二表單最新統計結果";
        tabPage2.UseVisualStyleBackColor = true;
        // 
        // dgvResult2
        // 
        dgvResult2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResult2.Dock = DockStyle.Fill;
        dgvResult2.Location = new Point(2, 2);
        dgvResult2.Name = "dgvResult2";
        dgvResult2.RowHeadersWidth = 62;
        dgvResult2.Size = new Size(942, 411);
        dgvResult2.TabIndex = 2;
        // 
        // tabPage1
        // 
        tabPage1.Controls.Add(dgvResult1);
        tabPage1.Location = new Point(4, 24);
        tabPage1.Margin = new Padding(2);
        tabPage1.Name = "tabPage1";
        tabPage1.Padding = new Padding(2);
        tabPage1.Size = new Size(946, 415);
        tabPage1.TabIndex = 0;
        tabPage1.Text = "第一表單最新統計結果";
        tabPage1.UseVisualStyleBackColor = true;
        // 
        // dgvResult1
        // 
        dgvResult1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResult1.Dock = DockStyle.Fill;
        dgvResult1.Location = new Point(2, 2);
        dgvResult1.Name = "dgvResult1";
        dgvResult1.RowHeadersWidth = 62;
        dgvResult1.Size = new Size(942, 411);
        dgvResult1.TabIndex = 1;
        // 
        // tabControl1
        // 
        tabControl1.Controls.Add(tabPage1);
        tabControl1.Controls.Add(tabPage2);
        tabControl1.Controls.Add(tabPage3);
        tabControl1.Controls.Add(tabPage4);
        tabControl1.Location = new Point(0, 146);
        tabControl1.Margin = new Padding(2);
        tabControl1.Name = "tabControl1";
        tabControl1.SelectedIndex = 0;
        tabControl1.Size = new Size(954, 443);
        tabControl1.TabIndex = 28;
        // 
        // tabPage4
        // 
        tabPage4.Controls.Add(dgvResult4);
        tabPage4.Location = new Point(4, 24);
        tabPage4.Name = "tabPage4";
        tabPage4.Padding = new Padding(3);
        tabPage4.Size = new Size(946, 415);
        tabPage4.TabIndex = 3;
        tabPage4.Text = "第四表單最新統計結果";
        tabPage4.UseVisualStyleBackColor = true;
        // 
        // dgvResult4
        // 
        dgvResult4.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dgvResult4.Dock = DockStyle.Fill;
        dgvResult4.Location = new Point(3, 3);
        dgvResult4.Name = "dgvResult4";
        dgvResult4.RowHeadersWidth = 62;
        dgvResult4.Size = new Size(940, 409);
        dgvResult4.TabIndex = 2;
        // 
        // label4
        // 
        label4.AutoSize = true;
        label4.Location = new Point(103, 16);
        label4.Name = "label4";
        label4.Size = new Size(58, 15);
        label4.TabIndex = 29;
        label4.Text = "報表紀錄:";
        // 
        // label7
        // 
        label7.AutoSize = true;
        label7.Location = new Point(105, 78);
        label7.Name = "label7";
        label7.Size = new Size(58, 15);
        label7.TabIndex = 30;
        label7.Text = "報表紀錄:";
        // 
        // label8
        // 
        label8.AutoSize = true;
        label8.Location = new Point(103, 46);
        label8.Name = "label8";
        label8.Size = new Size(58, 15);
        label8.TabIndex = 31;
        label8.Text = "報表紀錄:";
        // 
        // label9
        // 
        label9.AutoSize = true;
        label9.Location = new Point(105, 109);
        label9.Name = "label9";
        label9.Size = new Size(58, 15);
        label9.TabIndex = 35;
        label9.Text = "報表紀錄:";
        // 
        // txtBoxSelect4
        // 
        txtBoxSelect4.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxSelect4.Location = new Point(168, 105);
        txtBoxSelect4.Margin = new Padding(2);
        txtBoxSelect4.Name = "txtBoxSelect4";
        txtBoxSelect4.Size = new Size(54, 24);
        txtBoxSelect4.TabIndex = 33;
        txtBoxSelect4.TextAlign = HorizontalAlignment.Right;
        // 
        // btnSelect4
        // 
        btnSelect4.Location = new Point(12, 105);
        btnSelect4.Name = "btnSelect4";
        btnSelect4.Size = new Size(85, 25);
        btnSelect4.TabIndex = 32;
        btnSelect4.Text = "選擇表單4";
        btnSelect4.UseVisualStyleBackColor = true;
        btnSelect4.Click += btnSelect4_Click;
        // 
        // btnCalculateAllExcel
        // 
        btnCalculateAllExcel.BackColor = SystemColors.ActiveCaption;
        btnCalculateAllExcel.Location = new Point(226, 7);
        btnCalculateAllExcel.Margin = new Padding(2);
        btnCalculateAllExcel.Name = "btnCalculateAllExcel";
        btnCalculateAllExcel.Size = new Size(92, 129);
        btnCalculateAllExcel.TabIndex = 36;
        btnCalculateAllExcel.Text = "統計";
        btnCalculateAllExcel.UseVisualStyleBackColor = false;
        btnCalculateAllExcel.Click += btnCalculateAllExcel_Click;
        // 
        // AttendForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(954, 589);
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
        Controls.Add(groupBox2);
        Controls.Add(groupBox1);
        Controls.Add(btnSelect1);
        Name = "AttendForm";
        Text = "Form1";
        SizeChanged += AttendForm_SizeChanged;
        groupBox1.ResumeLayout(false);
        groupBox1.PerformLayout();
        groupBox2.ResumeLayout(false);
        groupBox2.PerformLayout();
        tabPage3.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResult3).EndInit();
        tabPage2.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResult2).EndInit();
        tabPage1.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResult1).EndInit();
        tabControl1.ResumeLayout(false);
        tabPage4.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)dgvResult4).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    private void btnCalculate_Click(object sender, EventArgs e)
    {
        OpenExcelFile(filenames[0], txtBoxSelect1.Text + ".xlsx", dgvResult1);
        MessageBox.Show("Finished!");
    }

#endregion
    private Button btnSelect1;
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
    private TextBox txtBoxSelect1;
    private Label label6;
    private Label label5;
    private TextBox txtBoxSelect2;
    private Button btnSelect2;
    private TextBox txtBoxSelect3;
    private Button btnSelect3;
    private CheckBox ckbCompare;
    private TabPage tabPage3;
    private DataGridView dgvResult3;
    private TabPage tabPage2;
    private DataGridView dgvResult2;
    private TabPage tabPage1;
    private DataGridView dgvResult1;
    private TabControl tabControl1;
    private Label label4;
    private Label label7;
    private Label label8;
    private Label label9;
    private TextBox txtBoxSelect4;
    private Button btnSelect4;
    private TabPage tabPage4;
    private DataGridView dgvResult4;
    private Button btnCalculateAllExcel;
    private CheckBox checkBox1;
}
