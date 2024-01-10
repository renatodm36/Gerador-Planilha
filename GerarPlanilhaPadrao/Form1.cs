using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.IO;
using System.Windows.Forms;

namespace GerarPlanilhaPadrao
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public void GerarPlanilhaXLS()
        {
            // Criar um novo workbook do tipo HSSFWorkbook (Excel 97-2003)
            IWorkbook workbook = new HSSFWorkbook();

            // Criar uma planilha
            ISheet sheet = workbook.CreateSheet("Planilha1");

            // Adicionar dados à planilha
            IRow row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue("Nome");
            row.CreateCell(1).SetCellValue("Idade");

            row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue(textBox1.Text); // Supondo que txtNome.Text seja o campo onde você digita o nome
            row.CreateCell(1).SetCellValue(textBox2.Text); // Supondo que txtIdade.Text seja o campo onde você digita a idade

            // Solicitar ao usuário onde salvar o arquivo usando SaveFileDialog
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Arquivos do Excel (*.xls)|*.xls";
            saveFileDialog.FileName = "dados.xls";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;

                // Salvar o arquivo no local escolhido pelo usuário
                using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fileStream);
                }

                MessageBox.Show("Arquivo XLS gerado com sucesso!");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GerarPlanilhaXLS();
        }
    }
}
