using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellApp
{
    public partial class FrmExcel : Form
    {
        // ************************ Consulte a Documentação ********************************
        // 
        //  https://docs.microsoft.com/en-us/visualstudio/vsto/excel-solutions?view=vs-2019
        //
        // *********************************************************************************

        readonly _Application exc = new Excel.Application();
        Workbook     excWb;
        Worksheet    excWs;        

        //  Arapuca
        readonly int numColArapuca = 6;

        public FrmExcel() { InitializeComponent(); }

        private void FrmExcel_Load( object sender, EventArgs e)  {}

        private void BtnFindExcel( object sender, EventArgs e)
        {
            // Show the FolderBrowserDialog.
            DialogResult result = diaFindFile.ShowDialog();
            if (result == DialogResult.OK)
            {               
                txtExcelBook.Text = diaFindFile.FileName;
            }
        }

        private void BtnOpenBook( object sender, EventArgs e)
        {            
            excWb = exc.Workbooks.Open( txtExcelBook.Text);
            excWs = excWb.Worksheets[1];
            //  Define o handler do eevnto Change

            //  Torna visivel a aplication e ativa a worksheet
            exc.Visible = true;
            excWs.Activate();           
        }

        private void BtnWrite2Cell_Click( object sender, EventArgs e)
        {            
            excWs.Cells[numLinha.Value, numColuna.Value] = txtText2Cell.Text + "";
        }

        private void BtnReadFromCell_Click( object sender, EventArgs e)
        {
            txtCell2Text.Text = (string)(excWs.Cells[numLinhaRead.Value, numColunaRead.Value].Value);
        }

        private void BtnClose(object sender, EventArgs e)
        {
            excWb.Close( false);
            exc.Quit();
        }

        private void CellsChange( Excel.Range Target)
        {
            if ( ( Target.Row >= numColorLin1.Value && Target.Row <= numColorLin2.Value) && (Target.Column == numColArapuca))
            {               
                try {
                    if (excWs.Cells[Target.Row, Target.Column + 1].Value == null)
                    {                        
                        excWs.Cells[Target.Row, Target.Column + 1].Value = excWs.Cells[Target.Row, Target.Column].Value;
                        excWs.Cells[Target.Row, Target.Column + 2].Value = excWs.Cells[Target.Row, Target.Column].Value;                        
                        return; // <<======
                    }

                    //  Menor que o minimo ?
                    if (excWs.Cells[Target.Row, Target.Column].Value <= excWs.Cells[Target.Row, Target.Column + 1].Value)
                    {
                        excWs.Cells[Target.Row, Target.Column + 1].Value = excWs.Cells[Target.Row, Target.Column].Value;
                    }
                    //  Maior que o maximo ?
                    if (excWs.Cells[Target.Row, Target.Column].Value >= excWs.Cells[Target.Row, Target.Column + 2].Value)
                    {
                        excWs.Cells[Target.Row, Target.Column + 2].Value = excWs.Cells[Target.Row, Target.Column].Value;
                    }
                } catch (Exception ex) { MessageBox.Show ( ex.Message); }
            }
        }

        private void BtnGetColor_Click( object sender, EventArgs e)
        {
            ColorDialog MyDialog = new ColorDialog
            {                
                AllowFullOpen = false,      // Keeps the user from selecting a custom color.                
                ShowHelp = true,            // Allows the user to get help. (The default is false.)                
                Color = lblRange.BackColor  // Sets the initial color select to the current text color.
            };

            // Update the text box color if the user clicks OK 
            if (MyDialog.ShowDialog() == DialogResult.OK)
                { lblRange.BackColor = MyDialog.Color; }
        }

        private void BtnApplyColor_Click( object sender, EventArgs e)
        {
            Range rng;

            try
            {
                rng = excWs.Range[ excWs.Cells[numColorLin1.Value, numColorCol1.Value], excWs.Cells[numColorLin2.Value, numColorCol2.Value] ];

                //  Set Background and Font Colors
                rng.Interior.Color = lblRange.BackColor;
                rng.Font.Color = Color.White;

                //  Set Text Alignent
                if (rdbAlignLeft.Checked) rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                if (rdbAlignCenter.Checked) rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                if (rdbAlignRight.Checked) rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                //  Set Borders - Geral ( default vale para todo o range desde que vc nao mude no detalhe )    
                rng.Borders.Weight = Excel.XlBorderWeight.xlThick;                  // Espessura da linha
                rng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;             // Tipo da linha : tracejada, pontilhada, continua

                // Set Linha por Linha
                rng.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;

                rng.Borders.get_Item(XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
                rng.Borders.get_Item(XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlHairline;

                rng.Borders.get_Item(XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlDouble;
                rng.Borders.get_Item(XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlDot;            
            
                rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone;  

                rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;
                rng.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlThin;

                // set Column Width
                rng.ColumnWidth = numColWidth.Value;

            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void BtnFormula_Click( object sender, EventArgs e)
        {
            Range rng;

            rng = excWs.Range[excWs.Cells[numColorLin1.Value, numColorCol1.Value - 2], excWs.Cells[numColorLin2.Value, numColorCol2.Value + 2]];
            rng.NumberFormat = "0.00";
            
            rng = excWs.Range[excWs.Cells[numColorLin1.Value, numColorCol1.Value], excWs.Cells[numColorLin2.Value, numColorCol2.Value]];
            rng.Formula = txtFormula.Text;
        }

        private void BtnCalculate_Click( object sender, EventArgs e)
        {
            excWs.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(CellsChange);
        }
    }
}