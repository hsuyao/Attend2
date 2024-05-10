
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
        btnCalculate = new Button();
        label1 = new Label();
        txtBoxStartColumn = new TextBox();
        groupBox1 = new GroupBox();
        rbHalfYear = new RadioButton();
        rbMonth = new RadioButton();
        rbWeek = new RadioButton();
        groupBox2 = new GroupBox();
        label6 = new Label();
        label5 = new Label();
        txtbIgnoreLevel = new TextBox();
        label3 = new Label();
        txtBoxStable = new TextBox();
        label2 = new Label();
        txtBoxOutput = new TextBox();
        label4 = new Label();
        txtBoxLord = new TextBox();
        textBox1 = new TextBox();
        txtBoxPray = new TextBox();
        label7 = new Label();
        btnSelect2 = new Button();
        textBox3 = new TextBox();
        txtBoxHome = new TextBox();
        label8 = new Label();
        btnSelect3 = new Button();
        ckbCompare = new CheckBox();
        groupBox3 = new GroupBox();
        groupBox4 = new GroupBox();
        groupBox5 = new GroupBox();
        textBox2 = new TextBox();
        textBox4 = new TextBox();
        textBox5 = new TextBox();
        groupBox1.SuspendLayout();
        groupBox2.SuspendLayout();
        groupBox3.SuspendLayout();
        groupBox4.SuspendLayout();
        groupBox5.SuspendLayout();
        SuspendLayout();
        // 
        // btnSelect1
        // 
        btnSelect1.Location = new Point(12, 6);
        btnSelect1.Name = "btnSelect1";
        btnSelect1.Size = new Size(68, 25);
        btnSelect1.TabIndex = 0;
        btnSelect1.Text = "選擇表單1";
        btnSelect1.UseVisualStyleBackColor = true;
        btnSelect1.Click += btnSelect_Click;
        // 
        // btnCalculate
        // 
        btnCalculate.Location = new Point(632, 452);
        btnCalculate.Margin = new Padding(2);
        btnCalculate.Name = "btnCalculate";
        btnCalculate.Size = new Size(69, 35);
        btnCalculate.TabIndex = 3;
        btnCalculate.Text = "Calculate";
        btnCalculate.UseVisualStyleBackColor = true;
        btnCalculate.Click += btnCalculate_Click;
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
        txtBoxStartColumn.TextChanged += txtBoxStartColumn_TextChanged_1;
        // 
        // groupBox1
        // 
        groupBox1.Controls.Add(rbHalfYear);
        groupBox1.Controls.Add(rbMonth);
        groupBox1.Controls.Add(rbWeek);
        groupBox1.Location = new Point(13, 329);
        groupBox1.Margin = new Padding(2);
        groupBox1.Name = "groupBox1";
        groupBox1.Padding = new Padding(2);
        groupBox1.Size = new Size(133, 108);
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
        groupBox2.Location = new Point(151, 329);
        groupBox2.Name = "groupBox2";
        groupBox2.Size = new Size(199, 158);
        groupBox2.TabIndex = 8;
        groupBox2.TabStop = false;
        groupBox2.Text = "統計參數";
        // 
        // label6
        // 
        label6.AutoSize = true;
        label6.Location = new Point(147, 56);
        label6.Margin = new Padding(2, 0, 2, 0);
        label6.Name = "label6";
        label6.Size = new Size(18, 15);
        label6.TabIndex = 12;
        label6.Text = "%";
        // 
        // label5
        // 
        label5.AutoSize = true;
        label5.Location = new Point(159, 27);
        label5.Margin = new Padding(2, 0, 2, 0);
        label5.Name = "label5";
        label5.Size = new Size(18, 15);
        label5.TabIndex = 11;
        label5.Text = "%";
        // 
        // txtbIgnoreLevel
        // 
        txtbIgnoreLevel.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtbIgnoreLevel.Location = new Point(101, 48);
        txtbIgnoreLevel.Margin = new Padding(2);
        txtbIgnoreLevel.Name = "txtbIgnoreLevel";
        txtbIgnoreLevel.Size = new Size(42, 23);
        txtbIgnoreLevel.TabIndex = 10;
        txtbIgnoreLevel.Text = "40";
        txtbIgnoreLevel.TextAlign = HorizontalAlignment.Center;
        txtbIgnoreLevel.TextChanged += txtbIgnoreLevel_TextChanged;
        // 
        // label3
        // 
        label3.AutoSize = true;
        label3.Location = new Point(15, 52);
        label3.Margin = new Padding(2, 0, 2, 0);
        label3.Name = "label3";
        label3.Size = new Size(82, 15);
        label3.TabIndex = 9;
        label3.Text = "剔除統計水位:";
        label3.Click += label3_Click;
        // 
        // txtBoxStable
        // 
        txtBoxStable.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxStable.Location = new Point(113, 19);
        txtBoxStable.Margin = new Padding(2);
        txtBoxStable.Name = "txtBoxStable";
        txtBoxStable.Size = new Size(42, 23);
        txtBoxStable.TabIndex = 8;
        txtBoxStable.Text = "40";
        txtBoxStable.TextAlign = HorizontalAlignment.Center;
        // 
        // label2
        // 
        label2.AutoSize = true;
        label2.Location = new Point(15, 22);
        label2.Margin = new Padding(2, 0, 2, 0);
        label2.Name = "label2";
        label2.Size = new Size(94, 15);
        label2.TabIndex = 7;
        label2.Text = "穩定聚會出席率:";
        // 
        // txtBoxOutput
        // 
        txtBoxOutput.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxOutput.Location = new Point(586, 8);
        txtBoxOutput.Margin = new Padding(2);
        txtBoxOutput.Name = "txtBoxOutput";
        txtBoxOutput.Size = new Size(115, 23);
        txtBoxOutput.TabIndex = 12;
        txtBoxOutput.Text = "主日.xlsx";
        txtBoxOutput.TextChanged += textBox1_TextChanged;
        // 
        // label4
        // 
        label4.AutoSize = true;
        label4.Location = new Point(536, 11);
        label4.Margin = new Padding(2, 0, 2, 0);
        label4.Name = "label4";
        label4.Size = new Size(46, 15);
        label4.TabIndex = 11;
        label4.Text = "輸出檔:";
        label4.Click += label4_Click;
        // 
        // txtBoxLord
        // 
        txtBoxLord.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxLord.Location = new Point(85, 7);
        txtBoxLord.Margin = new Padding(2);
        txtBoxLord.Name = "txtBoxLord";
        txtBoxLord.Size = new Size(437, 24);
        txtBoxLord.TabIndex = 13;
        txtBoxLord.TextAlign = HorizontalAlignment.Right;
        // 
        // textBox1
        // 
        textBox1.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        textBox1.Location = new Point(586, 43);
        textBox1.Margin = new Padding(2);
        textBox1.Name = "textBox1";
        textBox1.Size = new Size(115, 23);
        textBox1.TabIndex = 16;
        textBox1.Text = "禱告聚會.xlsx";
        // 
        // txtBoxPray
        // 
        txtBoxPray.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxPray.Location = new Point(85, 42);
        txtBoxPray.Margin = new Padding(2);
        txtBoxPray.Name = "txtBoxPray";
        txtBoxPray.Size = new Size(437, 24);
        txtBoxPray.TabIndex = 17;
        txtBoxPray.TextAlign = HorizontalAlignment.Right;
        // 
        // label7
        // 
        label7.AutoSize = true;
        label7.Location = new Point(536, 46);
        label7.Margin = new Padding(2, 0, 2, 0);
        label7.Name = "label7";
        label7.Size = new Size(46, 15);
        label7.TabIndex = 15;
        label7.Text = "輸出檔:";
        // 
        // btnSelect2
        // 
        btnSelect2.Location = new Point(12, 41);
        btnSelect2.Name = "btnSelect2";
        btnSelect2.Size = new Size(68, 25);
        btnSelect2.TabIndex = 14;
        btnSelect2.Text = "選擇表單2";
        btnSelect2.UseVisualStyleBackColor = true;
        // 
        // textBox3
        // 
        textBox3.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        textBox3.Location = new Point(586, 79);
        textBox3.Margin = new Padding(2);
        textBox3.Name = "textBox3";
        textBox3.Size = new Size(115, 23);
        textBox3.TabIndex = 20;
        textBox3.Text = "小排.xlsx";
        // 
        // txtBoxHome
        // 
        txtBoxHome.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxHome.Location = new Point(85, 78);
        txtBoxHome.Margin = new Padding(2);
        txtBoxHome.Name = "txtBoxHome";
        txtBoxHome.Size = new Size(437, 24);
        txtBoxHome.TabIndex = 21;
        txtBoxHome.TextAlign = HorizontalAlignment.Right;
        // 
        // label8
        // 
        label8.AutoSize = true;
        label8.Location = new Point(536, 82);
        label8.Margin = new Padding(2, 0, 2, 0);
        label8.Name = "label8";
        label8.Size = new Size(46, 15);
        label8.TabIndex = 19;
        label8.Text = "輸出檔:";
        // 
        // btnSelect3
        // 
        btnSelect3.Location = new Point(12, 77);
        btnSelect3.Name = "btnSelect3";
        btnSelect3.Size = new Size(68, 25);
        btnSelect3.TabIndex = 18;
        btnSelect3.Text = "選擇表單3";
        btnSelect3.UseVisualStyleBackColor = true;
        // 
        // ckbCompare
        // 
        ckbCompare.AutoSize = true;
        ckbCompare.Location = new Point(15, 104);
        ckbCompare.Name = "ckbCompare";
        ckbCompare.Size = new Size(122, 19);
        ckbCompare.TabIndex = 13;
        ckbCompare.Text = "比較前後統計週期";
        ckbCompare.UseVisualStyleBackColor = true;
        // 
        // groupBox3
        // 
        groupBox3.Controls.Add(textBox2);
        groupBox3.Location = new Point(13, 118);
        groupBox3.Name = "groupBox3";
        groupBox3.Size = new Size(219, 181);
        groupBox3.TabIndex = 22;
        groupBox3.TabStop = false;
        groupBox3.Text = "表單一最新週期統計結果";
        // 
        // groupBox4
        // 
        groupBox4.Controls.Add(textBox4);
        groupBox4.Location = new Point(252, 118);
        groupBox4.Name = "groupBox4";
        groupBox4.Size = new Size(219, 181);
        groupBox4.TabIndex = 23;
        groupBox4.TabStop = false;
        groupBox4.Text = "表單二最新週期統計結果";
        // 
        // groupBox5
        // 
        groupBox5.Controls.Add(textBox5);
        groupBox5.Location = new Point(481, 118);
        groupBox5.Name = "groupBox5";
        groupBox5.Size = new Size(219, 181);
        groupBox5.TabIndex = 23;
        groupBox5.TabStop = false;
        groupBox5.Text = "表單三最新週期統計結果";
        // 
        // textBox2
        // 
        textBox2.Location = new Point(16, 22);
        textBox2.Multiline = true;
        textBox2.Name = "textBox2";
        textBox2.Size = new Size(197, 153);
        textBox2.TabIndex = 0;
        // 
        // textBox4
        // 
        textBox4.Location = new Point(12, 22);
        textBox4.Multiline = true;
        textBox4.Name = "textBox4";
        textBox4.Size = new Size(197, 153);
        textBox4.TabIndex = 1;
        // 
        // textBox5
        // 
        textBox5.Location = new Point(6, 22);
        textBox5.Multiline = true;
        textBox5.Name = "textBox5";
        textBox5.Size = new Size(197, 153);
        textBox5.TabIndex = 2;
        // 
        // AttendForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(734, 498);
        Controls.Add(groupBox5);
        Controls.Add(groupBox4);
        Controls.Add(groupBox3);
        Controls.Add(textBox3);
        Controls.Add(txtBoxHome);
        Controls.Add(label8);
        Controls.Add(btnSelect3);
        Controls.Add(textBox1);
        Controls.Add(txtBoxPray);
        Controls.Add(label7);
        Controls.Add(btnSelect2);
        Controls.Add(txtBoxOutput);
        Controls.Add(txtBoxLord);
        Controls.Add(label4);
        Controls.Add(groupBox2);
        Controls.Add(groupBox1);
        Controls.Add(btnCalculate);
        Controls.Add(btnSelect1);
        Name = "AttendForm";
        Text = "Form1";
        groupBox1.ResumeLayout(false);
        groupBox1.PerformLayout();
        groupBox2.ResumeLayout(false);
        groupBox2.PerformLayout();
        groupBox3.ResumeLayout(false);
        groupBox3.PerformLayout();
        groupBox4.ResumeLayout(false);
        groupBox4.PerformLayout();
        groupBox5.ResumeLayout(false);
        groupBox5.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    private void btnCalculate_Click(object sender, EventArgs e)
    {
        OpenExcelFile();
    }

    #endregion
    private FolderBrowserDialog folderBrowserDialog1;
    private Button btnSelect1;
    private Button btnCalculate;
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
    private TextBox textBox1;
    private TextBox txtBoxPray;
    private Label label7;
    private Button btnSelect2;
    private TextBox textBox3;
    private TextBox txtBoxHome;
    private Label label8;
    private Button btnSelect3;
    private CheckBox ckbCompare;
    private GroupBox groupBox3;
    private TextBox textBox2;
    private GroupBox groupBox4;
    private TextBox textBox4;
    private GroupBox groupBox5;
    private TextBox textBox5;
}
