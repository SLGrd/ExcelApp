using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellApp
{
    public partial class FrmRealTime : Form
    {
        // ************************ Consulte a Documentação ********************************
        // 
        //  https://docs.microsoft.com/en-us/visualstudio/vsto/excel-solutions?view=vs-2019
        //
        // *********************************************************************************

        public class DdeTopic
        {
            public string Topic     { get; set; }
            public string Descricao { get; set; }

            public DdeTopic( string topic, string descricao)
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
            new DdeTopic("ult", "Valor ultimo negocio")
        };

        readonly  _Application exc = new Excel.Application();
        Workbook  excWb;
        Worksheet excWs;

        public FrmRealTime() { InitializeComponent();}

        private void FrmRealTime_Load(object sender, EventArgs e)
        {
            //  Fill Columns combo
            cmbDdeTopics0.DisplayMember = "Descricao";
            cmbDdeTopics0.ValueMember = "Topic";
            cmbDdeTopics0.DataSource = Topics;

            List<DdeTopic> Topics1 = new List<DdeTopic>();
            for ( int i = 0; i < Topics.Count; i++) { Topics1.Add(Topics[i]); }
            cmbDdeTopics1.DisplayMember = "Descricao";
            cmbDdeTopics1.ValueMember = "Topic";
            cmbDdeTopics1.DataSource = Topics1;

            List<DdeTopic> Topics2 = new List<DdeTopic>();
            for (int i = 0; i < Topics.Count; i++) { Topics2.Add(Topics[i]); }
            cmbDdeTopics2.DisplayMember = "Descricao";
            cmbDdeTopics2.ValueMember = "Topic";
            cmbDdeTopics2.DataSource = Topics2;

            List<DdeTopic> Topics3 = new List<DdeTopic>();
            for (int i = 0; i < Topics.Count; i++) { Topics3.Add(Topics[i]); }
            cmbDdeTopics3.DisplayMember = "Descricao";
            cmbDdeTopics3.ValueMember = "Topic";
            cmbDdeTopics3.DataSource = Topics3;

            List<DdeTopic> Topics4 = new List<DdeTopic>();
            for (int i = 0; i < Topics.Count; i++) { Topics4.Add(Topics[i]); }
            cmbDdeTopics4.DisplayMember = "Descricao";
            cmbDdeTopics4.ValueMember = "Topic";
            cmbDdeTopics4.DataSource = Topics4;

            List<DdeTopic> Topics5 = new List<DdeTopic>();
            for (int i = 0; i < Topics.Count; i++) { Topics5.Add(Topics[i]); }
            cmbDdeTopics5.DisplayMember = "Descricao";
            cmbDdeTopics5.ValueMember = "Topic";
            cmbDdeTopics5.DataSource = Topics5;

            List<DdeTopic> Topics6 = new List<DdeTopic>();
            for (int i = 0; i < Topics.Count; i++) { Topics6.Add(Topics[i]); }
            cmbDdeTopics6.DisplayMember = "Descricao";
            cmbDdeTopics6.ValueMember = "Topic";
            cmbDdeTopics6.DataSource = Topics6;
        }

        private void BtnFindExcel(object sender, EventArgs e)
        {

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\Planilhas";
                openFileDialog.Filter = "Exc files (*.xls)|*.xls|Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtExcelBook.Text = openFileDialog.FileName;
                }
            }
        }

        private void BtnOpenBook(object sender, EventArgs e)
        {
            excWb = exc.Workbooks.Open(txtExcelBook.Text);
            excWs = excWb.Worksheets[1];
            //  Define o handler do eevnto Change

            //  Torna visivel a aplication e ativa a worksheet
            exc.Visible = true;
            excWs.Activate();
        }

        private void NumColWidth_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                //Range rng = excWs.Range[excWs.Cells[numColorLin1.Value, numColorCol1.Value], excWs.Cells[numColorLin2.Value, numColorCol2.Value]];
                //// set Column Width
                //rng.ColumnWidth = numColWidth.Value;
            }
            catch (Exception ex)  { MessageBox.Show(ex.Message);}
        }

        private void TxtColTile_TextChanged(object sender, EventArgs e)
        {
            //excWs.Cells[numLinTitle.Value, cmbDdeTopics1.SelectedIndex + 1] = txtColTitle.Text + "";
        }
    }
}
