
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
        btnCalculateExcel2 = new Button();
        btnCalculateExcel3 = new Button();
        tabPage3 = new TabPage();
        dgvResult3 = new DataGridView();
        tabPage2 = new TabPage();
        dgvResult2 = new DataGridView();
        tabPage1 = new TabPage();
        dgvResult1 = new DataGridView();
        tabControl1 = new TabControl();
        groupBox1.SuspendLayout();
        groupBox2.SuspendLayout();
        tabPage3.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult3).BeginInit();
        tabPage2.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult2).BeginInit();
        tabPage1.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)dgvResult1).BeginInit();
        tabControl1.SuspendLayout();
        SuspendLayout();
        // 
        // btnSelect1
        // 
        btnSelect1.Location = new Point(19, 17);
        btnSelect1.Margin = new Padding(5);
        btnSelect1.Name = "btnSelect1";
        btnSelect1.Size = new Size(134, 38);
        btnSelect1.TabIndex = 0;
        btnSelect1.Text = "選擇表單1";
        btnSelect1.UseVisualStyleBackColor = true;
        btnSelect1.Click += btnSelect_Click;
        // 
        // btnCalculateExcel1
        // 
        btnCalculateExcel1.BackColor = SystemColors.ActiveCaption;
        btnCalculateExcel1.Location = new Point(500, 2);
        btnCalculateExcel1.Name = "btnCalculateExcel1";
        btnCalculateExcel1.Size = new Size(145, 67);
        btnCalculateExcel1.TabIndex = 3;
        btnCalculateExcel1.Text = "統計表單1";
        btnCalculateExcel1.UseVisualStyleBackColor = false;
        btnCalculateExcel1.Click += btnCalculate_Click;
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(24, 121);
        label1.Name = "label1";
        label1.Size = new Size(158, 23);
        label1.TabIndex = 5;
        label1.Text = "出席紀錄開始欄位:";
        // 
        // txtBoxStartColumn
        // 
        txtBoxStartColumn.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxStartColumn.Location = new Point(196, 115);
        txtBoxStartColumn.Name = "txtBoxStartColumn";
        txtBoxStartColumn.Size = new Size(62, 33);
        txtBoxStartColumn.TabIndex = 6;
        txtBoxStartColumn.Text = "I";
        txtBoxStartColumn.TextAlign = HorizontalAlignment.Center;
        // 
        // groupBox1
        // 
        groupBox1.Controls.Add(rbHalfYear);
        groupBox1.Controls.Add(rbMonth);
        groupBox1.Controls.Add(rbWeek);
        groupBox1.Location = new Point(926, 23);
        groupBox1.Name = "groupBox1";
        groupBox1.Size = new Size(209, 195);
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
        groupBox2.Location = new Point(1142, 20);
        groupBox2.Margin = new Padding(5);
        groupBox2.Name = "groupBox2";
        groupBox2.Padding = new Padding(5);
        groupBox2.Size = new Size(328, 196);
        groupBox2.TabIndex = 8;
        groupBox2.TabStop = false;
        groupBox2.Text = "統計參數";
        // 
        // ckbCompare
        // 
        ckbCompare.AutoSize = true;
        ckbCompare.Location = new Point(24, 156);
        ckbCompare.Margin = new Padding(5);
        ckbCompare.Name = "ckbCompare";
        ckbCompare.Size = new Size(180, 27);
        ckbCompare.TabIndex = 13;
        ckbCompare.Text = "比較前後統計週期";
        ckbCompare.UseVisualStyleBackColor = true;
        // 
        // label6
        // 
        label6.AutoSize = true;
        label6.Location = new Point(174, 81);
        label6.Name = "label6";
        label6.Size = new Size(67, 23);
        label6.TabIndex = 12;
        label6.Text = "% 剔除";
        // 
        // label5
        // 
        label5.AutoSize = true;
        label5.Location = new Point(228, 32);
        label5.Name = "label5";
        label5.Size = new Size(26, 23);
        label5.TabIndex = 11;
        label5.Text = "%";
        // 
        // txtbIgnoreLevel
        // 
        txtbIgnoreLevel.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtbIgnoreLevel.Location = new Point(130, 78);
        txtbIgnoreLevel.Name = "txtbIgnoreLevel";
        txtbIgnoreLevel.Size = new Size(46, 30);
        txtbIgnoreLevel.TabIndex = 10;
        txtbIgnoreLevel.Text = "40";
        txtbIgnoreLevel.TextAlign = HorizontalAlignment.Center;
        // 
        // label3
        // 
        label3.AutoSize = true;
        label3.Location = new Point(24, 81);
        label3.Name = "label3";
        label3.Size = new Size(100, 23);
        label3.TabIndex = 9;
        label3.Text = "低於中位數";
        // 
        // txtBoxStable
        // 
        txtBoxStable.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxStable.Location = new Point(178, 29);
        txtBoxStable.Name = "txtBoxStable";
        txtBoxStable.Size = new Size(43, 30);
        txtBoxStable.TabIndex = 8;
        txtBoxStable.Text = "40";
        txtBoxStable.TextAlign = HorizontalAlignment.Center;
        // 
        // label2
        // 
        label2.AutoSize = true;
        label2.Location = new Point(24, 32);
        label2.Name = "label2";
        label2.Size = new Size(140, 23);
        label2.TabIndex = 7;
        label2.Text = "穩定聚會出席率:";
        // 
        // txtBoxOutput
        // 
        txtBoxOutput.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxOutput.Location = new Point(729, 31);
        txtBoxOutput.Name = "txtBoxOutput";
        txtBoxOutput.Size = new Size(178, 30);
        txtBoxOutput.TabIndex = 12;
        txtBoxOutput.Text = "主日.xlsx";
        // 
        // label4
        // 
        label4.AutoSize = true;
        label4.Location = new Point(651, 35);
        label4.Name = "label4";
        label4.Size = new Size(68, 23);
        label4.TabIndex = 11;
        label4.Text = "輸出檔:";
        // 
        // txtBoxLord
        // 
        txtBoxLord.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxLord.Location = new Point(160, 18);
        txtBoxLord.Name = "txtBoxLord";
        txtBoxLord.Size = new Size(320, 33);
        txtBoxLord.TabIndex = 13;
        txtBoxLord.TextAlign = HorizontalAlignment.Right;
        // 
        // txtBoxOutputPray
        // 
        txtBoxOutputPray.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxOutputPray.Location = new Point(729, 103);
        txtBoxOutputPray.Name = "txtBoxOutputPray";
        txtBoxOutputPray.Size = new Size(151, 30);
        txtBoxOutputPray.TabIndex = 16;
        txtBoxOutputPray.Text = "禱告聚會.xlsx";
        // 
        // txtBoxPray
        // 
        txtBoxPray.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxPray.Location = new Point(160, 89);
        txtBoxPray.Name = "txtBoxPray";
        txtBoxPray.Size = new Size(320, 33);
        txtBoxPray.TabIndex = 17;
        txtBoxPray.TextAlign = HorizontalAlignment.Right;
        // 
        // label7
        // 
        label7.AutoSize = true;
        label7.Location = new Point(651, 107);
        label7.Name = "label7";
        label7.Size = new Size(68, 23);
        label7.TabIndex = 15;
        label7.Text = "輸出檔:";
        // 
        // btnSelect2
        // 
        btnSelect2.Location = new Point(19, 86);
        btnSelect2.Margin = new Padding(5);
        btnSelect2.Name = "btnSelect2";
        btnSelect2.Size = new Size(134, 38);
        btnSelect2.TabIndex = 14;
        btnSelect2.Text = "選擇表單2";
        btnSelect2.UseVisualStyleBackColor = true;
        btnSelect2.Click += btnSelect2_Click;
        // 
        // txtBoxOutputHome
        // 
        txtBoxOutputHome.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxOutputHome.Location = new Point(729, 172);
        txtBoxOutputHome.Name = "txtBoxOutputHome";
        txtBoxOutputHome.Size = new Size(178, 30);
        txtBoxOutputHome.TabIndex = 20;
        txtBoxOutputHome.Text = "小排.xlsx";
        // 
        // txtBoxHome
        // 
        txtBoxHome.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxHome.Location = new Point(160, 161);
        txtBoxHome.Name = "txtBoxHome";
        txtBoxHome.Size = new Size(320, 33);
        txtBoxHome.TabIndex = 21;
        txtBoxHome.TextAlign = HorizontalAlignment.Right;
        // 
        // label8
        // 
        label8.AutoSize = true;
        label8.Location = new Point(651, 176);
        label8.Name = "label8";
        label8.Size = new Size(68, 23);
        label8.TabIndex = 19;
        label8.Text = "輸出檔:";
        // 
        // btnSelect3
        // 
        btnSelect3.Location = new Point(19, 159);
        btnSelect3.Margin = new Padding(5);
        btnSelect3.Name = "btnSelect3";
        btnSelect3.Size = new Size(134, 38);
        btnSelect3.TabIndex = 18;
        btnSelect3.Text = "選擇表單3";
        btnSelect3.UseVisualStyleBackColor = true;
        btnSelect3.Click += btnSelect3_Click;
        // 
        // btnCalculateExcel2
        // 
        btnCalculateExcel2.BackColor = SystemColors.ActiveCaption;
        btnCalculateExcel2.Location = new Point(500, 75);
        btnCalculateExcel2.Name = "btnCalculateExcel2";
        btnCalculateExcel2.Size = new Size(145, 63);
        btnCalculateExcel2.TabIndex = 24;
        btnCalculateExcel2.Text = "統計表單2";
        btnCalculateExcel2.UseVisualStyleBackColor = false;
        btnCalculateExcel2.Click += btnCalculateSheet2_Click;
        // 
        // btnCalculateExcel3
        // 
        btnCalculateExcel3.BackColor = SystemColors.ActiveCaption;
        btnCalculateExcel3.Location = new Point(500, 144);
        btnCalculateExcel3.Name = "btnCalculateExcel3";
        btnCalculateExcel3.Size = new Size(145, 69);
        btnCalculateExcel3.TabIndex = 25;
        btnCalculateExcel3.Text = "統計表單3";
        btnCalculateExcel3.UseVisualStyleBackColor = false;
        btnCalculateExcel3.Click += btnCalculateSheet3_Click;
        // 
        // tabPage3
        // 
        tabPage3.Controls.Add(dgvResult3);
        tabPage3.Location = new Point(4, 32);
        tabPage3.Name = "tabPage3";
        tabPage3.Padding = new Padding(3);
        tabPage3.Size = new Size(1491, 643);
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
        dgvResult3.Size = new Size(1485, 637);
        dgvResult3.TabIndex = 4;
        // 
        // tabPage2
        // 
        tabPage2.Controls.Add(dgvResult2);
        tabPage2.Location = new Point(4, 32);
        tabPage2.Name = "tabPage2";
        tabPage2.Padding = new Padding(3);
        tabPage2.Size = new Size(1491, 643);
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
        dgvResult2.Size = new Size(1485, 637);
        dgvResult2.TabIndex = 2;
        // 
        // tabPage1
        // 
        tabPage1.Controls.Add(dgvResult1);
        tabPage1.Location = new Point(4, 32);
        tabPage1.Name = "tabPage1";
        tabPage1.Padding = new Padding(3);
        tabPage1.Size = new Size(1491, 643);
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
        dgvResult1.Size = new Size(1485, 637);
        dgvResult1.TabIndex = 1;
        // 
        // tabControl1
        // 
        tabControl1.Controls.Add(tabPage1);
        tabControl1.Controls.Add(tabPage2);
        tabControl1.Controls.Add(tabPage3);
        tabControl1.Location = new Point(0, 224);
        tabControl1.Name = "tabControl1";
        tabControl1.SelectedIndex = 0;
        tabControl1.Size = new Size(1499, 679);
        tabControl1.TabIndex = 28;
        tabControl1.SizeChanged += tabControl1_SizeChanged;
        // 
        // AttendForm
        // 
        AutoScaleDimensions = new SizeF(11F, 23F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(1499, 903);
        Controls.Add(tabControl1);
        Controls.Add(btnCalculateExcel3);
        Controls.Add(btnCalculateExcel2);
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
        Margin = new Padding(5);
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
        ResumeLayout(false);
        PerformLayout();
    }

    private void btnCalculate_Click(object sender, EventArgs e)
    {
        OpenExcelFile(txtBoxLord.Text, txtBoxOutput.Text, dgvResult1);
    }

#endregion
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
    private Button btnCalculateExcel2;
    private Button btnCalculateExcel3;
    private TabPage tabPage3;
    private DataGridView dgvResult3;
    private TabPage tabPage2;
    private DataGridView dgvResult2;
    private TabPage tabPage1;
    private DataGridView dgvResult1;
    private TabControl tabControl1;
}
