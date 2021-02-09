using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellApp
{
    public partial class FrmDde : Form
    {
        // ************************ Consulte a Documentação ********************************
        // 
        //  https://docs.microsoft.com/en-us/visualstudio/vsto/excel-solutions?view=vs-2019
        //
        // *********************************************************************************

        readonly _Application exc = new Excel.Application();
        private Workbook excWb;
        private Worksheet excWs;
       
        public class DdeTopic
        {
            public string Topic { get; set; }
            public string Descricao { get; set; }

            public DdeTopic(string topic, string descricao)
            {
                Topic = topic;
                Descricao = descricao;
            }
        }

        readonly List<DdeTopic> Topics = new List<DdeTopic>()
        {
            new DdeTopic("text", "Digite o comando ==>>"),
            new DdeTopic("mofc", "Melhor Oferta de Compra"),
            new DdeTopic("mofv", "Melhor Oferta de Venda"),
            new DdeTopic("ult",  "Valor ultimo negocio")
            //
            //  Aqui devem ser colocados todos os topicos oferecidos pelo DDE server
            //
        };

        public FrmDde() { InitializeComponent(); }

        private void FrmDde_Load(object sender, EventArgs e)
        {
            //  Fill DDE Topics combo
            cmbDdeTopics.DisplayMember = "Descricao";
            cmbDdeTopics.ValueMember   = "Topic";
            cmbDdeTopics.DataSource    = Topics;

            txtDdeServer.Text = "BULLDDE";
        }

        private void BtnFindExcel(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                try
                {
                    openFileDialog.InitialDirectory = "c:\\Planilhas";
                    openFileDialog.Filter = "Excel files (*.xls*)|(*.xls*)|All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 2;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        txtExcelBook.Text = openFileDialog.FileName;
                        excWb = exc.Workbooks.Open(txtExcelBook.Text);
                        excWs = excWb.Worksheets[1];

                        //  Torna visivel a aplication e ativa a worksheet
                        exc.DisplayStatusBar = false;
                        exc.Visible = true;
                        excWs.Activate();
                    }
                }   catch (Exception ex) { MessageBox.Show(ex.Message); };
            }
        }

        private void CmbDdeTopics_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DdeTopic ddeItem = (DdeTopic)cmbDdeTopics.SelectedItem;
                if (ddeItem.Topic == "text")
                {
                    //  Move focus to Result text box
                }
                else
                {
                    //  Monta comandp " = DDEServer | Topics ! Ativo "
                    txtResult.Text = "=" + txtDdeServer.Text.Trim() + "|" + ddeItem.Topic.ToString() + "!" + txtAtivo.Text.Trim();
                }
                this.ActiveControl = txtResult; //  Mov focus to Result text box
            }   catch (Exception ex) { MessageBox.Show(ex.Message); }
        }

        private void BtnWrite2Cell_Click(object sender, EventArgs e)
        {
            try
            {
                exc.DisplayStatusBar = false;
                decimal m = numLinha.Value;  //Linha
                string L = "R" + m.ToString("N0");

                string At = "R4C10";
                if (m > 16) At = "R17C10";                    

                excWs = exc.ActiveWorkbook.ActiveSheet;
                if ( numColuna.Value == 0)
                {                    
                    excWs.Cells[ m, 1] =  txtAtivo.Text.Trim();
                    excWs.Cells[ m, 2] = "=BullDde|Strike!" + txtAtivo.Text.Trim();
                    excWs.Cells[ m, 3] = "=BullDde|MOFC!" + txtAtivo.Text.Trim();
                    excWs.Cells[ m, 4] = "=BullDde|MOFV!" + txtAtivo.Text.Trim();

                    if ( rdbBuy.Checked ) excWs.Cells[m, 6] = "Buy";
                    if ( rdbSell.Checked) excWs.Cells[m, 6] = "Sell";                  

                    //  Set Line Column adressing system
                    exc.ReferenceStyle = Excel.XlReferenceStyle.xlR1C1;
                    excWs.Cells[ m, 8] = "=" + (L + "C5") + (" * ") + (L + "C7");

                    if ( string.Compare( txtAtivo.Text.Trim().Substring(4, 1), "L") <= 0 )
                        excWs.Cells[m, 9] = $"=IF( {At} > {L}C2, {At} - {L}C2 , 0)";
                    if (string.Compare(txtAtivo.Text.Trim().Substring(4, 1), "L") > 0)
                        excWs.Cells[m, 9] = $"=IF( {At} < {L}C2, {L}C2 - {At}, 0)";
                    
                    excWs.Cells[m, 10] = $"=({L}C3 +{L}C4) / 2";

                    string Action = (string)(excWs.Cells[ m, 6].Value);
                    if (Action.Equals("Buy"))
                    {
                        excWs.Cells[m, 11] = $"={L}C5 * ( {L}C9  - {L}C7)";
                        excWs.Cells[m, 13] = $"={L}C5 * ( {L}C10 - {L}C7)";
                    }
                    if (Action.Equals("Sell"))
                    {
                        excWs.Cells[m, 11] = $"={L}C5 * ( {L}C7 - {L}C9 )";
                        excWs.Cells[m, 13] = $"={L}C5 * ( {L}C7 - {L}C10)";
                    }

                    excWs.Cells[m, 14] = $"=BullDDE|Neg!{txtAtivo.Text.Trim()}";
                    excWs.Cells[m, 15] = $"=BullDDE|Min!{txtAtivo.Text.Trim()}";
                    excWs.Cells[m, 16] = $"=BullDDE|Max!{txtAtivo.Text.Trim()}";
                    excWs.Cells[m, 18] = $"=100 * ( {L}C2 - {At} ) / {At}";
                }
                else if (numColuna.Value == 6)  
                {
                    if (rdbBuy.Checked)
                    {
                        excWs.Cells[m, 6] = "Buy";
                        excWs.Cells[m, 11] = $"={L}C5 * ( {L}C9 - {L}C7 )";
                    }
                    if (rdbSell.Checked)
                    {
                        excWs.Cells[m, 6] = "Sell";
                        excWs.Cells[m, 11] = $"={L}C5 * ( {L}C7 - {L}C9 ) ";
                    }      
                }
                else
                {
                    excWs.Cells[numLinha.Value, numColuna.Value] = txtResult.Text.Trim() + "";
                }
            }  catch (Exception ex) { MessageBox.Show(ex.Message); }            
        }      
        
        private void BtnCopyWithReplace(object sender, EventArgs e)
        {
            string w, s;
            Range rngFrom, rngTo;

            excWs = exc.ActiveWorkbook.ActiveSheet;
            //  Browse selected line columns
            for ( decimal n = numColFrom.Value; n <= numColTo.Value ; n++ )
            {
                try
                {                    
                    rngFrom = (Range)excWs.Cells[numLinhaFrom.Value, n];     
                    //  Get range contents
                    w = rngFrom.Formula;
                    //  Get formula prefix if any
                    s = w.Substring(0, w.IndexOf("!") + 1);
                    if ( s.Length == 0)
                        { rngTo = (Range)excWs.Cells[numLinhaTo.Value, n];                // Not a formula
                          rngFrom.Copy( rngTo); }
                    else
                        { excWs.Cells[numLinhaTo.Value, n] = s + txtAtivo.Text.Trim(); }  // It is a formula                  
                }   catch ( Exception ex ) { MessageBox.Show(ex.Message); }
            }
        }
    }
}