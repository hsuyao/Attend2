
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
        btnSelect = new Button();
        lblCurFile = new Label();
        btnCalculate = new Button();
        label1 = new Label();
        txtBoxStartColumn = new TextBox();
        groupBox1 = new GroupBox();
        rbHalfYear = new RadioButton();
        rbMonth = new RadioButton();
        rbWeek = new RadioButton();
        groupBox2 = new GroupBox();
        txtBoxStable = new TextBox();
        label2 = new Label();
        label3 = new Label();
        txtbIgnoreLevel = new TextBox();
        groupBox1.SuspendLayout();
        groupBox2.SuspendLayout();
        SuspendLayout();
        // 
        // btnSelect
        // 
        btnSelect.Location = new Point(12, 7);
        btnSelect.Name = "btnSelect";
        btnSelect.Size = new Size(133, 30);
        btnSelect.TabIndex = 0;
        btnSelect.Text = "Select Lord Day File";
        btnSelect.UseVisualStyleBackColor = true;
        btnSelect.Click += btnSelect_Click;
        // 
        // lblCurFile
        // 
        lblCurFile.AutoSize = true;
        lblCurFile.Location = new Point(151, 15);
        lblCurFile.Name = "lblCurFile";
        lblCurFile.Size = new Size(121, 15);
        lblCurFile.TabIndex = 2;
        lblCurFile.Text = "Select your excel file";
        // 
        // btnCalculate
        // 
        btnCalculate.Location = new Point(12, 234);
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
        label1.Location = new Point(15, 24);
        label1.Margin = new Padding(2, 0, 2, 0);
        label1.Name = "label1";
        label1.Size = new Size(83, 15);
        label1.TabIndex = 5;
        label1.Text = "Start Column:";
        // 
        // txtBoxStartColumn
        // 
        txtBoxStartColumn.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxStartColumn.Location = new Point(102, 20);
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
        groupBox1.Location = new Point(12, 55);
        groupBox1.Margin = new Padding(2);
        groupBox1.Name = "groupBox1";
        groupBox1.Padding = new Padding(2);
        groupBox1.Size = new Size(133, 130);
        groupBox1.TabIndex = 7;
        groupBox1.TabStop = false;
        groupBox1.Text = "Calculate Period";
        // 
        // rbHalfYear
        // 
        rbHalfYear.AutoSize = true;
        rbHalfYear.Location = new Point(16, 89);
        rbHalfYear.Name = "rbHalfYear";
        rbHalfYear.Size = new Size(76, 19);
        rbHalfYear.TabIndex = 11;
        rbHalfYear.Text = "Half Year";
        rbHalfYear.UseVisualStyleBackColor = true;
        // 
        // rbMonth
        // 
        rbMonth.AutoSize = true;
        rbMonth.Location = new Point(16, 57);
        rbMonth.Name = "rbMonth";
        rbMonth.Size = new Size(63, 19);
        rbMonth.TabIndex = 10;
        rbMonth.Text = "Month";
        rbMonth.UseVisualStyleBackColor = true;
        // 
        // rbWeek
        // 
        rbWeek.AutoSize = true;
        rbWeek.Checked = true;
        rbWeek.Location = new Point(16, 27);
        rbWeek.Name = "rbWeek";
        rbWeek.Size = new Size(57, 19);
        rbWeek.TabIndex = 9;
        rbWeek.TabStop = true;
        rbWeek.Text = "Week";
        rbWeek.UseVisualStyleBackColor = true;
        // 
        // groupBox2
        // 
        groupBox2.Controls.Add(txtbIgnoreLevel);
        groupBox2.Controls.Add(label3);
        groupBox2.Controls.Add(txtBoxStable);
        groupBox2.Controls.Add(label2);
        groupBox2.Controls.Add(label1);
        groupBox2.Controls.Add(txtBoxStartColumn);
        groupBox2.Location = new Point(151, 55);
        groupBox2.Name = "groupBox2";
        groupBox2.Size = new Size(193, 130);
        groupBox2.TabIndex = 8;
        groupBox2.TabStop = false;
        groupBox2.Text = "Parameters";
        // 
        // txtBoxStable
        // 
        txtBoxStable.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxStable.Location = new Point(94, 53);
        txtBoxStable.Margin = new Padding(2);
        txtBoxStable.Name = "txtBoxStable";
        txtBoxStable.Size = new Size(42, 23);
        txtBoxStable.TabIndex = 8;
        txtBoxStable.Text = "0.4";
        txtBoxStable.TextAlign = HorizontalAlignment.Center;

        // 
        // label2
        // 
        label2.AutoSize = true;
        label2.Location = new Point(15, 56);
        label2.Margin = new Padding(2, 0, 2, 0);
        label2.Name = "label2";
        label2.Size = new Size(75, 15);
        label2.TabIndex = 7;
        label2.Text = "Stable Rate:";
        // 
        // label3
        // 
        label3.AutoSize = true;
        label3.Location = new Point(15, 89);
        label3.Margin = new Padding(2, 0, 2, 0);
        label3.Name = "label3";
        label3.Size = new Size(106, 15);
        label3.TabIndex = 9;
        label3.Text = "Ignore Threshold:";
        // 
        // txtbIgnoreLevel
        // 
        txtbIgnoreLevel.Font = new Font("Microsoft JhengHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtbIgnoreLevel.Location = new Point(125, 85);
        txtbIgnoreLevel.Margin = new Padding(2);
        txtbIgnoreLevel.Name = "txtbIgnoreLevel";
        txtbIgnoreLevel.Size = new Size(42, 23);
        txtbIgnoreLevel.TabIndex = 10;
        txtbIgnoreLevel.Text = "0.5";
        txtbIgnoreLevel.TextAlign = HorizontalAlignment.Center;
        // 
        // AttendForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(475, 280);
        Controls.Add(groupBox2);
        Controls.Add(groupBox1);
        Controls.Add(btnCalculate);
        Controls.Add(lblCurFile);
        Controls.Add(btnSelect);
        Name = "AttendForm";
        Text = "Form1";
        groupBox1.ResumeLayout(false);
        groupBox1.PerformLayout();
        groupBox2.ResumeLayout(false);
        groupBox2.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    private void btnCalculate_Click(object sender, EventArgs e)
    {
        OpenExcelFile();
    }

    #endregion
    private FolderBrowserDialog folderBrowserDialog1;
    private Button btnSelect;
    private Label lblCurFile;
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
}
