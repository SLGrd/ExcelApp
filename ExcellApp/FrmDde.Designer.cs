namespace ExcellApp
{
    partial class FrmDde
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
            this.txtAtivo = new System.Windows.Forms.TextBox();
            this.txtDdeServer = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtExcelBook = new System.Windows.Forms.TextBox();
            this.btnFindExcel = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.numColuna = new System.Windows.Forms.NumericUpDown();
            this.numLinha = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.btnWrite2Cell = new System.Windows.Forms.Button();
            this.txtResult = new System.Windows.Forms.TextBox();
            this.cmbDdeTopics = new System.Windows.Forms.ComboBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label6 = new System.Windows.Forms.Label();
            this.numColTo = new System.Windows.Forms.NumericUpDown();
            this.numColFrom = new System.Windows.Forms.NumericUpDown();
            this.numLinhaTo = new System.Windows.Forms.NumericUpDown();
            this.numLinhaFrom = new System.Windows.Forms.NumericUpDown();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.btnCopyWithReplace = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rdbSell = new System.Windows.Forms.RadioButton();
            this.rdbBuy = new System.Windows.Forms.RadioButton();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColuna)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLinha)).BeginInit();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColTo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColFrom)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLinhaTo)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLinhaFrom)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtAtivo
            // 
            this.txtAtivo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtAtivo.Location = new System.Drawing.Point(190, 98);
            this.txtAtivo.Name = "txtAtivo";
            this.txtAtivo.Size = new System.Drawing.Size(184, 26);
            this.txtAtivo.TabIndex = 42;
            // 
            // txtDdeServer
            // 
            this.txtDdeServer.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDdeServer.Location = new System.Drawing.Point(190, 56);
            this.txtDdeServer.Name = "txtDdeServer";
            this.txtDdeServer.Size = new System.Drawing.Size(184, 26);
            this.txtDdeServer.TabIndex = 43;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Maroon;
            this.label3.Location = new System.Drawing.Point(14, 57);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(124, 24);
            this.label3.TabIndex = 40;
            this.label3.Text = "Servidor DDE";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.Maroon;
            this.label5.Location = new System.Drawing.Point(14, 99);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(51, 24);
            this.label5.TabIndex = 41;
            this.label5.Text = "Ativo";
            // 
            // txtExcelBook
            // 
            this.txtExcelBook.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtExcelBook.Location = new System.Drawing.Point(190, 14);
            this.txtExcelBook.Name = "txtExcelBook";
            this.txtExcelBook.Size = new System.Drawing.Size(390, 26);
            this.txtExcelBook.TabIndex = 39;
            // 
            // btnFindExcel
            // 
            this.btnFindExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnFindExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFindExcel.Location = new System.Drawing.Point(14, 14);
            this.btnFindExcel.Name = "btnFindExcel";
            this.btnFindExcel.Size = new System.Drawing.Size(163, 26);
            this.btnFindExcel.TabIndex = 38;
            this.btnFindExcel.Text = "Planilha Excel";
            this.btnFindExcel.UseVisualStyleBackColor = true;
            this.btnFindExcel.Click += new System.EventHandler(this.BtnFindExcel);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.numColuna);
            this.panel1.Controls.Add(this.numLinha);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label12);
            this.panel1.Controls.Add(this.btnWrite2Cell);
            this.panel1.Controls.Add(this.txtResult);
            this.panel1.Controls.Add(this.cmbDdeTopics);
            this.panel1.Location = new System.Drawing.Point(16, 145);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(749, 164);
            this.panel1.TabIndex = 44;
            // 
            // numColuna
            // 
            this.numColuna.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numColuna.Location = new System.Drawing.Point(90, 61);
            this.numColuna.Name = "numColuna";
            this.numColuna.Size = new System.Drawing.Size(62, 26);
            this.numColuna.TabIndex = 66;
            // 
            // numLinha
            // 
            this.numLinha.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numLinha.Location = new System.Drawing.Point(12, 61);
            this.numLinha.Name = "numLinha";
            this.numLinha.Size = new System.Drawing.Size(62, 26);
            this.numLinha.TabIndex = 65;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(494, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(238, 36);
            this.label2.TabIndex = 61;
            this.label2.Text = "Resultado";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(172, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(300, 36);
            this.label1.TabIndex = 62;
            this.label1.Text = "Comando";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(16, 20);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(48, 20);
            this.label4.TabIndex = 63;
            this.label4.Text = "Linha";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(89, 20);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(59, 20);
            this.label12.TabIndex = 63;
            this.label12.Text = "Coluna";
            this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnWrite2Cell
            // 
            this.btnWrite2Cell.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnWrite2Cell.BackColor = System.Drawing.Color.Wheat;
            this.btnWrite2Cell.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnWrite2Cell.Location = new System.Drawing.Point(172, 105);
            this.btnWrite2Cell.Name = "btnWrite2Cell";
            this.btnWrite2Cell.Size = new System.Drawing.Size(560, 41);
            this.btnWrite2Cell.TabIndex = 60;
            this.btnWrite2Cell.Text = "Write to Excel SpreadSheet";
            this.btnWrite2Cell.UseVisualStyleBackColor = false;
            this.btnWrite2Cell.Click += new System.EventHandler(this.BtnWrite2Cell_Click);
            // 
            // txtResult
            // 
            this.txtResult.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtResult.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtResult.Location = new System.Drawing.Point(494, 62);
            this.txtResult.Name = "txtResult";
            this.txtResult.Size = new System.Drawing.Size(238, 26);
            this.txtResult.TabIndex = 57;
            // 
            // cmbDdeTopics
            // 
            this.cmbDdeTopics.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbDdeTopics.FormattingEnabled = true;
            this.cmbDdeTopics.Location = new System.Drawing.Point(172, 60);
            this.cmbDdeTopics.Name = "cmbDdeTopics";
            this.cmbDdeTopics.Size = new System.Drawing.Size(300, 28);
            this.cmbDdeTopics.TabIndex = 46;
            this.cmbDdeTopics.SelectedIndexChanged += new System.EventHandler(this.CmbDdeTopics_SelectedIndexChanged);
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.label6);
            this.panel2.Controls.Add(this.numColTo);
            this.panel2.Controls.Add(this.numColFrom);
            this.panel2.Controls.Add(this.numLinhaTo);
            this.panel2.Controls.Add(this.numLinhaFrom);
            this.panel2.Controls.Add(this.label8);
            this.panel2.Controls.Add(this.label9);
            this.panel2.Controls.Add(this.btnCopyWithReplace);
            this.panel2.Location = new System.Drawing.Point(17, 382);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(749, 117);
            this.panel2.TabIndex = 45;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(171, 20);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(192, 20);
            this.label6.TabIndex = 69;
            this.label6.Text = "Da coluna m até Coluna n";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // numColTo
            // 
            this.numColTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numColTo.Location = new System.Drawing.Point(269, 60);
            this.numColTo.Name = "numColTo";
            this.numColTo.Size = new System.Drawing.Size(62, 26);
            this.numColTo.TabIndex = 68;
            // 
            // numColFrom
            // 
            this.numColFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numColFrom.Location = new System.Drawing.Point(197, 60);
            this.numColFrom.Name = "numColFrom";
            this.numColFrom.Size = new System.Drawing.Size(62, 26);
            this.numColFrom.TabIndex = 67;
            // 
            // numLinhaTo
            // 
            this.numLinhaTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numLinhaTo.Location = new System.Drawing.Point(90, 61);
            this.numLinhaTo.Name = "numLinhaTo";
            this.numLinhaTo.Size = new System.Drawing.Size(62, 26);
            this.numLinhaTo.TabIndex = 66;
            // 
            // numLinhaFrom
            // 
            this.numLinhaFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numLinhaFrom.Location = new System.Drawing.Point(12, 61);
            this.numLinhaFrom.Name = "numLinhaFrom";
            this.numLinhaFrom.Size = new System.Drawing.Size(62, 26);
            this.numLinhaFrom.TabIndex = 65;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(9, 20);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(63, 20);
            this.label8.TabIndex = 63;
            this.label8.Text = "Linha X";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(66, 20);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(89, 20);
            this.label9.TabIndex = 63;
            this.label9.Text = "= > Linha Y";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnCopyWithReplace
            // 
            this.btnCopyWithReplace.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCopyWithReplace.BackColor = System.Drawing.Color.Wheat;
            this.btnCopyWithReplace.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCopyWithReplace.Location = new System.Drawing.Point(380, 48);
            this.btnCopyWithReplace.Name = "btnCopyWithReplace";
            this.btnCopyWithReplace.Size = new System.Drawing.Size(352, 41);
            this.btnCopyWithReplace.TabIndex = 60;
            this.btnCopyWithReplace.Text = "Execute Copy function";
            this.btnCopyWithReplace.UseVisualStyleBackColor = false;
            this.btnCopyWithReplace.Click += new System.EventHandler(this.BtnCopyWithReplace);
            // 
            // label10
            // 
            this.label10.BackColor = System.Drawing.Color.Transparent;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.ForeColor = System.Drawing.Color.DodgerBlue;
            this.label10.Location = new System.Drawing.Point(17, 334);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(743, 48);
            this.label10.TabIndex = 46;
            this.label10.Text = "Para adaptar o conteudo da linha x na linha Y mudando apenas o ativo use a função" +
    " abaixo";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rdbBuy);
            this.groupBox1.Controls.Add(this.rdbSell);
            this.groupBox1.Location = new System.Drawing.Point(398, 97);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(360, 26);
            this.groupBox1.TabIndex = 47;
            this.groupBox1.TabStop = false;
            // 
            // rdbSell
            // 
            this.rdbSell.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdbSell.Location = new System.Drawing.Point(107, 6);
            this.rdbSell.Name = "rdbSell";
            this.rdbSell.Size = new System.Drawing.Size(85, 24);
            this.rdbSell.TabIndex = 2;
            this.rdbSell.TabStop = true;
            this.rdbSell.Text = "Sell";
            this.rdbSell.UseVisualStyleBackColor = true;
            // 
            // rdbBuy
            // 
            this.rdbBuy.Checked = true;
            this.rdbBuy.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdbBuy.Location = new System.Drawing.Point(6, 6);
            this.rdbBuy.Name = "rdbBuy";
            this.rdbBuy.Size = new System.Drawing.Size(85, 24);
            this.rdbBuy.TabIndex = 3;
            this.rdbBuy.TabStop = true;
            this.rdbBuy.Text = "Buy";
            this.rdbBuy.UseVisualStyleBackColor = true;
            // 
            // FrmDde
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(783, 527);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.txtAtivo);
            this.Controls.Add(this.txtDdeServer);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtExcelBook);
            this.Controls.Add(this.btnFindExcel);
            this.ForeColor = System.Drawing.Color.Black;
            this.Name = "FrmDde";
            this.Text = "FrmDde";
            this.Load += new System.EventHandler(this.FrmDde_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColuna)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLinha)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColTo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numColFrom)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLinhaTo)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLinhaFrom)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtAtivo;
        private System.Windows.Forms.TextBox txtDdeServer;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtExcelBook;
        private System.Windows.Forms.Button btnFindExcel;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Button btnWrite2Cell;
        private System.Windows.Forms.TextBox txtResult;
        private System.Windows.Forms.ComboBox cmbDdeTopics;
        private System.Windows.Forms.NumericUpDown numColuna;
        private System.Windows.Forms.NumericUpDown numLinha;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.NumericUpDown numLinhaTo;
        private System.Windows.Forms.NumericUpDown numLinhaFrom;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button btnCopyWithReplace;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.NumericUpDown numColTo;
        private System.Windows.Forms.NumericUpDown numColFrom;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rdbBuy;
        private System.Windows.Forms.RadioButton rdbSell;
    }
}