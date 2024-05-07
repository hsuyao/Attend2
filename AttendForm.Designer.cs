
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
        panel1 = new Panel();
        label1 = new Label();
        txtBoxStartColumn = new TextBox();
        SuspendLayout();
        // 
        // btnSelect
        // 
        btnSelect.Location = new Point(19, 11);
        btnSelect.Margin = new Padding(5);
        btnSelect.Name = "btnSelect";
        btnSelect.Size = new Size(108, 46);
        btnSelect.TabIndex = 0;
        btnSelect.Text = "Select";
        btnSelect.UseVisualStyleBackColor = true;
        btnSelect.Click += btnSelect_Click;
        // 
        // lblCurFile
        // 
        lblCurFile.AutoSize = true;
        lblCurFile.Location = new Point(140, 23);
        lblCurFile.Margin = new Padding(5, 0, 5, 0);
        lblCurFile.Name = "lblCurFile";
        lblCurFile.Size = new Size(183, 23);
        lblCurFile.TabIndex = 2;
        lblCurFile.Text = "Select your excel file";
        // 
        // btnCalculate
        // 
        btnCalculate.Enabled = false;
        btnCalculate.Location = new Point(19, 65);
        btnCalculate.Name = "btnCalculate";
        btnCalculate.Size = new Size(108, 53);
        btnCalculate.TabIndex = 3;
        btnCalculate.Text = "Calculate";
        btnCalculate.UseVisualStyleBackColor = true;
        btnCalculate.Click += btnCalculate_Click;
        // 
        // panel1
        // 
        panel1.Location = new Point(140, 155);
        panel1.Name = "panel1";
        panel1.Size = new Size(893, 148);
        panel1.TabIndex = 4;
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(140, 80);
        label1.Name = "label1";
        label1.Size = new Size(127, 23);
        label1.TabIndex = 5;
        label1.Text = "Start Column:";
        // 
        // txtBoxStartColumn
        // 
        txtBoxStartColumn.Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 136);
        txtBoxStartColumn.Location = new Point(273, 75);
        txtBoxStartColumn.Name = "txtBoxStartColumn";
        txtBoxStartColumn.Size = new Size(29, 33);
        txtBoxStartColumn.TabIndex = 6;
        txtBoxStartColumn.Text = "I";
        txtBoxStartColumn.TextAlign = HorizontalAlignment.Center;
        // 
        // AttendForm
        // 
        AutoScaleDimensions = new SizeF(11F, 23F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(1057, 315);
        Controls.Add(txtBoxStartColumn);
        Controls.Add(label1);
        Controls.Add(panel1);
        Controls.Add(btnCalculate);
        Controls.Add(lblCurFile);
        Controls.Add(btnSelect);
        Margin = new Padding(5);
        Name = "AttendForm";
        Text = "Form1";
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
    private Panel panel1;
    private Label label1;
    private TextBox txtBoxStartColumn;
}
