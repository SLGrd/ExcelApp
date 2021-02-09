namespace ExcellApp
{
    partial class FrmExcel
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.diaFindFile = new System.Windows.Forms.OpenFileDialog();
            this.btnFindExcel = new System.Windows.Forms.Button();
            this.txtExcelBook = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.numLinha = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.numColuna = new System.Windows.Forms.NumericUpDown();
            this.txtText2Cell = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btnOpenExcel = new System.Windows.Forms.Button();
            this.btnWrite2Cell = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.numColunaRead = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.numLinhaRead = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.txtCell2Text = new System.Windows.Forms.TextBox();
            this.btnReadFromCell = new System.Windows.Forms.Button();
            this.numColorCol1 = new System.Windows.Forms.NumericUpDown();
            this.label7 = new System.Windows.Forms.Label();
            this.numColorLin1 = new System.Windows.Forms.NumericUpDown();
            this.label8 = new System.Windows.Forms.Label();
            this.numColorCol2 = new System.Windows.Forms.NumericUpDown();
            this.label9 = new System.Windows.Forms.Label();
            this.numColorLin2 = new System.Windows.Forms.NumericUpDown();
            this.label10 = new System.Windows.Forms.Label();
            this.lblRange = new System.Windows.Forms.Label();
            this.btnGetColor = new System.Windows.Forms.Button();
            this.btnApplyColor = new System.Windows.Forms.Button();
            this.numColWidth = new System.Windows.Forms.NumericUpDown();
            this.label12 = new System.Windows.Forms.Label();
            this.grpAlign = new System.Windows.Forms.GroupBox();
            this.rdbAlignRight = new System.Windows.Forms.RadioButton();
            this.rdbAlignCenter = new System.Windows.Forms.RadioButton();
            this.rdbAlignLeft = new System.Windows.Forms.RadioButton();
            this.label11 = new System.Windows.Forms.Label();
            this.txtFormula = new System.Windows.Forms.TextBox();
            this.btnFormula = new System.Windows.Forms.Button();
            this.btnCalculate = new System.Windows.Forms.Button();
            this.numArapuca = new System.Windows.Forms.NumericUpDown();
            this.label13 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numLinha)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColuna)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColunaRead)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLinhaRead)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColorCol1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColorLin1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColorCol2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColorLin2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColWidth)).BeginInit();
            this.grpAlign.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numArapuca)).BeginInit();
            this.SuspendLayout();
            // 
            // diaFindFile
            // 
            this.diaFindFile.FileName = "*.*";
            this.diaFindFile.Filter = "Excel Files | *.xlsx";
            // 
            // btnFindExcel
            // 
            this.btnFindExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFindExcel.Location = new System.Drawing.Point(12, 15);
            this.btnFindExcel.Name = "btnFindExcel";
            this.btnFindExcel.Size = new System.Drawing.Size(93, 29);
            this.btnFindExcel.TabIndex = 0;
            this.btnFindExcel.Text = "Find Excel";
            this.btnFindExcel.UseVisualStyleBackColor = true;
            this.btnFindExcel.Click += new System.EventHandler(this.BtnFindExcel);
            // 
            // txtExcelBook
            // 
            this.txtExcelBook.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtExcelBook.Location = new System.Drawing.Point(111, 15);
            this.txtExcelBook.Name = "txtExcelBook";
            this.txtExcelBook.Size = new System.Drawing.Size(230, 29);
            this.txtExcelBook.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Maroon;
            this.label1.Location = new System.Drawing.Point(12, 93);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "Linha";
            // 
            // numLinha
            // 
            this.numLinha.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numLinha.Location = new System.Drawing.Point(12, 109);
            this.numLinha.Name = "numLinha";
            this.numLinha.Size = new System.Drawing.Size(62, 26);
            this.numLinha.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Maroon;
            this.label2.Location = new System.Drawing.Point(84, 93);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(46, 15);
            this.label2.TabIndex = 2;
            this.label2.Text = "Coluna";
            // 
            // numColuna
            // 
            this.numColuna.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numColuna.Location = new System.Drawing.Point(84, 109);
            this.numColuna.Name = "numColuna";
            this.numColuna.Size = new System.Drawing.Size(62, 26);
            this.numColuna.TabIndex = 3;
            // 
            // txtText2Cell
            // 
            this.txtText2Cell.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtText2Cell.Location = new System.Drawing.Point(157, 109);
            this.txtText2Cell.Name = "txtText2Cell";
            this.txtText2Cell.Size = new System.Drawing.Size(184, 26);
            this.txtText2Cell.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Maroon;
            this.label3.Location = new System.Drawing.Point(154, 93);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(37, 15);
            this.label3.TabIndex = 2;
            this.label3.Text = "Texto";
            // 
            // btnOpenExcel
            // 
            this.btnOpenExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpenExcel.Location = new System.Drawing.Point(12, 50);
            this.btnOpenExcel.Name = "btnOpenExcel";
            this.btnOpenExcel.Size = new System.Drawing.Size(329, 29);
            this.btnOpenExcel.TabIndex = 0;
            this.btnOpenExcel.Text = "Open Excel SpreadSheet";
            this.btnOpenExcel.UseVisualStyleBackColor = true;
            this.btnOpenExcel.Click += new System.EventHandler(this.BtnOpenBook);
            // 
            // btnWrite2Cell
            // 
            this.btnWrite2Cell.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnWrite2Cell.Location = new System.Drawing.Point(12, 141);
            this.btnWrite2Cell.Name = "btnWrite2Cell";
            this.btnWrite2Cell.Size = new System.Drawing.Size(329, 29);
            this.btnWrite2Cell.TabIndex = 0;
            this.btnWrite2Cell.Text = "Write to Excel SpreadSheet";
            this.btnWrite2Cell.UseVisualStyleBackColor = true;
            this.btnWrite2Cell.Click += new System.EventHandler(this.BtnWrite2Cell_Click);
            // 
            // btnClose
            // 
            this.btnClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClose.Location = new System.Drawing.Point(12, 562);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(329, 29);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Close Excel SpreadSheet";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.BtnClose);
            // 
            // numColunaRead
            // 
            this.numColunaRead.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numColunaRead.Location = new System.Drawing.Point(87, 201);
            this.numColunaRead.Name = "numColunaRead";
            this.numColunaRead.Size = new System.Drawing.Size(62, 26);
            this.numColunaRead.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Maroon;
            this.label4.Location = new System.Drawing.Point(157, 185);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(37, 15);
            this.label4.TabIndex = 6;
            this.label4.Text = "Texto";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.Maroon;
            this.label5.Location = new System.Drawing.Point(87, 185);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(46, 15);
            this.label5.TabIndex = 7;
            this.label5.Text = "Coluna";
            // 
            // numLinhaRead
            // 
            this.numLinhaRead.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numLinhaRead.Location = new System.Drawing.Point(15, 201);
            this.numLinhaRead.Name = "numLinhaRead";
            this.numLinhaRead.Size = new System.Drawing.Size(62, 26);
            this.numLinhaRead.TabIndex = 10;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.Maroon;
            this.label6.Location = new System.Drawing.Point(15, 185);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(38, 15);
            this.label6.TabIndex = 8;
            this.label6.Text = "Linha";
            // 
            // txtCell2Text
            // 
            this.txtCell2Text.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCell2Text.Location = new System.Drawing.Point(160, 201);
            this.txtCell2Text.Name = "txtCell2Text";
            this.txtCell2Text.Size = new System.Drawing.Size(184, 26);
            this.txtCell2Text.TabIndex = 5;
            // 
            // btnReadFromCell
            // 
            this.btnReadFromCell.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnReadFromCell.Location = new System.Drawing.Point(15, 233);
            this.btnReadFromCell.Name = "btnReadFromCell";
            this.btnReadFromCell.Size = new System.Drawing.Size(329, 29);
            this.btnReadFromCell.TabIndex = 4;
            this.btnReadFromCell.Text = "Read from Excel SpreadSheet";
            this.btnReadFromCell.UseVisualStyleBackColor = true;
            this.btnReadFromCell.Click += new System.EventHandler(this.BtnReadFromCell_Click);
            // 
            // numColorCol1
            // 
            this.numColorCol1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numColorCol1.Location = new System.Drawing.Point(88, 321);
            this.numColorCol1.Name = "numColorCol1";
            this.numColorCol1.Size = new System.Drawing.Size(62, 26);
            this.numColorCol1.TabIndex = 13;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.Maroon;
            this.label7.Location = new System.Drawing.Point(88, 305);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(46, 15);
            this.label7.TabIndex = 11;
            this.label7.Text = "Coluna";
            // 
            // numColorLin1
            // 
            this.numColorLin1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numColorLin1.Location = new System.Drawing.Point(16, 321);
            this.numColorLin1.Name = "numColorLin1";
            this.numColorLin1.Size = new System.Drawing.Size(62, 26);
            this.numColorLin1.TabIndex = 14;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.Color.Maroon;
            this.label8.Location = new System.Drawing.Point(16, 305);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(38, 15);
            this.label8.TabIndex = 12;
            this.label8.Text = "Linha";
            // 
            // numColorCol2
            // 
            this.numColorCol2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numColorCol2.Location = new System.Drawing.Point(280, 321);
            this.numColorCol2.Name = "numColorCol2";
            this.numColorCol2.Size = new System.Drawing.Size(62, 26);
            this.numColorCol2.TabIndex = 17;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.ForeColor = System.Drawing.Color.Maroon;
            this.label9.Location = new System.Drawing.Point(280, 305);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(46, 15);
            this.label9.TabIndex = 15;
            this.label9.Text = "Coluna";
            // 
            // numColorLin2
            // 
            this.numColorLin2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numColorLin2.Location = new System.Drawing.Point(208, 321);
            this.numColorLin2.Name = "numColorLin2";
            this.numColorLin2.Size = new System.Drawing.Size(62, 26);
            this.numColorLin2.TabIndex = 18;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.ForeColor = System.Drawing.Color.Maroon;
            this.label10.Location = new System.Drawing.Point(208, 305);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(38, 15);
            this.label10.TabIndex = 16;
            this.label10.Text = "Linha";
            // 
            // lblRange
            // 
            this.lblRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblRange.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.lblRange.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRange.Location = new System.Drawing.Point(15, 284);
            this.lblRange.Name = "lblRange";
            this.lblRange.Size = new System.Drawing.Size(329, 21);
            this.lblRange.TabIndex = 19;
            this.lblRange.Text = "Range";
            this.lblRange.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // btnGetColor
            // 
            this.btnGetColor.Location = new System.Drawing.Point(159, 321);
            this.btnGetColor.Name = "btnGetColor";
            this.btnGetColor.Size = new System.Drawing.Size(43, 26);
            this.btnGetColor.TabIndex = 20;
            this.btnGetColor.UseVisualStyleBackColor = true;
            this.btnGetColor.Click += new System.EventHandler(this.BtnGetColor_Click);
            // 
            // btnApplyColor
            // 
            this.btnApplyColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApplyColor.Location = new System.Drawing.Point(12, 433);
            this.btnApplyColor.Name = "btnApplyColor";
            this.btnApplyColor.Size = new System.Drawing.Size(138, 44);
            this.btnApplyColor.TabIndex = 4;
            this.btnApplyColor.Text = "Apply Settings to Range";
            this.btnApplyColor.UseVisualStyleBackColor = true;
            this.btnApplyColor.Click += new System.EventHandler(this.BtnApplyColor_Click);
            // 
            // numColWidth
            // 
            this.numColWidth.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numColWidth.Location = new System.Drawing.Point(12, 398);
            this.numColWidth.Name = "numColWidth";
            this.numColWidth.Size = new System.Drawing.Size(137, 26);
            this.numColWidth.TabIndex = 23;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.ForeColor = System.Drawing.Color.Maroon;
            this.label12.Location = new System.Drawing.Point(12, 382);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(84, 15);
            this.label12.TabIndex = 22;
            this.label12.Text = "Column Width";
            // 
            // grpAlign
            // 
            this.grpAlign.Controls.Add(this.rdbAlignRight);
            this.grpAlign.Controls.Add(this.rdbAlignCenter);
            this.grpAlign.Controls.Add(this.rdbAlignLeft);
            this.grpAlign.Location = new System.Drawing.Point(15, 350);
            this.grpAlign.Name = "grpAlign";
            this.grpAlign.Size = new System.Drawing.Size(329, 26);
            this.grpAlign.TabIndex = 24;
            this.grpAlign.TabStop = false;
            // 
            // rdbAlignRight
            // 
            this.rdbAlignRight.AutoSize = true;
            this.rdbAlignRight.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdbAlignRight.Location = new System.Drawing.Point(243, 8);
            this.rdbAlignRight.Name = "rdbAlignRight";
            this.rdbAlignRight.Size = new System.Drawing.Size(90, 20);
            this.rdbAlignRight.TabIndex = 0;
            this.rdbAlignRight.Text = "Align Right";
            this.rdbAlignRight.UseVisualStyleBackColor = true;
            // 
            // rdbAlignCenter
            // 
            this.rdbAlignCenter.AutoSize = true;
            this.rdbAlignCenter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdbAlignCenter.Location = new System.Drawing.Point(119, 8);
            this.rdbAlignCenter.Name = "rdbAlignCenter";
            this.rdbAlignCenter.Size = new System.Drawing.Size(98, 20);
            this.rdbAlignCenter.TabIndex = 0;
            this.rdbAlignCenter.Text = "Align Center";
            this.rdbAlignCenter.UseVisualStyleBackColor = true;
            // 
            // rdbAlignLeft
            // 
            this.rdbAlignLeft.AutoSize = true;
            this.rdbAlignLeft.Checked = true;
            this.rdbAlignLeft.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdbAlignLeft.Location = new System.Drawing.Point(5, 8);
            this.rdbAlignLeft.Name = "rdbAlignLeft";
            this.rdbAlignLeft.Size = new System.Drawing.Size(80, 20);
            this.rdbAlignLeft.TabIndex = 0;
            this.rdbAlignLeft.TabStop = true;
            this.rdbAlignLeft.Text = "Align Left";
            this.rdbAlignLeft.UseVisualStyleBackColor = true;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.ForeColor = System.Drawing.Color.Maroon;
            this.label11.Location = new System.Drawing.Point(157, 383);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(37, 15);
            this.label11.TabIndex = 26;
            this.label11.Text = "Texto";
            // 
            // txtFormula
            // 
            this.txtFormula.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFormula.Location = new System.Drawing.Point(160, 399);
            this.txtFormula.Name = "txtFormula";
            this.txtFormula.Size = new System.Drawing.Size(184, 26);
            this.txtFormula.TabIndex = 25;
            this.txtFormula.Text = "=B6 * C6";
            // 
            // btnFormula
            // 
            this.btnFormula.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFormula.Location = new System.Drawing.Point(160, 433);
            this.btnFormula.Name = "btnFormula";
            this.btnFormula.Size = new System.Drawing.Size(184, 44);
            this.btnFormula.TabIndex = 4;
            this.btnFormula.Text = "Apply Formula to Range";
            this.btnFormula.UseVisualStyleBackColor = true;
            this.btnFormula.Click += new System.EventHandler(this.BtnFormula_Click);
            // 
            // btnCalculate
            // 
            this.btnCalculate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCalculate.Location = new System.Drawing.Point(160, 493);
            this.btnCalculate.Name = "btnCalculate";
            this.btnCalculate.Size = new System.Drawing.Size(184, 36);
            this.btnCalculate.TabIndex = 4;
            this.btnCalculate.Text = "Arma Arapuca";
            this.btnCalculate.UseVisualStyleBackColor = true;
            this.btnCalculate.Click += new System.EventHandler(this.BtnCalculate_Click);
            // 
            // numArapuca
            // 
            this.numArapuca.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numArapuca.Location = new System.Drawing.Point(15, 503);
            this.numArapuca.Name = "numArapuca";
            this.numArapuca.Size = new System.Drawing.Size(134, 26);
            this.numArapuca.TabIndex = 28;
            this.numArapuca.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.ForeColor = System.Drawing.Color.Maroon;
            this.label13.Location = new System.Drawing.Point(15, 487);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(46, 15);
            this.label13.TabIndex = 27;
            this.label13.Text = "Coluna";
            // 
            // FrmExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(355, 606);
            this.Controls.Add(this.numArapuca);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.txtFormula);
            this.Controls.Add(this.grpAlign);
            this.Controls.Add(this.numColWidth);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.btnGetColor);
            this.Controls.Add(this.lblRange);
            this.Controls.Add(this.numColorCol2);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.numColorLin2);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.numColorCol1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.numColorLin1);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.numColunaRead);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.numLinhaRead);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtCell2Text);
            this.Controls.Add(this.btnCalculate);
            this.Controls.Add(this.btnFormula);
            this.Controls.Add(this.btnApplyColor);
            this.Controls.Add(this.btnReadFromCell);
            this.Controls.Add(this.numColuna);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.numLinha);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtText2Cell);
            this.Controls.Add(this.txtExcelBook);
            this.Controls.Add(this.btnWrite2Cell);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnOpenExcel);
            this.Controls.Add(this.btnFindExcel);
            this.Name = "FrmExcel";
            this.Text = "Excell Automation Example";
            this.Load += new System.EventHandler(this.FrmExcel_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numLinha)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColuna)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColunaRead)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLinhaRead)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColorCol1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColorLin1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColorCol2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColorLin2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColWidth)).EndInit();
            this.grpAlign.ResumeLayout(false);
            this.grpAlign.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numArapuca)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog diaFindFile;
        private System.Windows.Forms.Button btnFindExcel;
        private System.Windows.Forms.TextBox txtExcelBook;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown numLinha;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown numColuna;
        private System.Windows.Forms.TextBox txtText2Cell;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnOpenExcel;
        private System.Windows.Forms.Button btnWrite2Cell;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.NumericUpDown numColunaRead;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.NumericUpDown numLinhaRead;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtCell2Text;
        private System.Windows.Forms.Button btnReadFromCell;
        private System.Windows.Forms.NumericUpDown numColorCol1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.NumericUpDown numColorLin1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.NumericUpDown numColorCol2;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.NumericUpDown numColorLin2;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label lblRange;
        private System.Windows.Forms.Button btnGetColor;
        private System.Windows.Forms.Button btnApplyColor;
        private System.Windows.Forms.NumericUpDown numColWidth;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.GroupBox grpAlign;
        private System.Windows.Forms.RadioButton rdbAlignRight;
        private System.Windows.Forms.RadioButton rdbAlignCenter;
        private System.Windows.Forms.RadioButton rdbAlignLeft;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtFormula;
        private System.Windows.Forms.Button btnFormula;
        private System.Windows.Forms.Button btnCalculate;
        private System.Windows.Forms.NumericUpDown numArapuca;
        private System.Windows.Forms.Label label13;
    }
}

