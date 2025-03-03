using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.Windows.Forms;

namespace ImpressaoDeCheques
{
    public partial class FormImpressaoCheque : Form
    {
        public FormImpressaoCheque()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.lblNomeCliente = new System.Windows.Forms.Label();
            this.txtNomeCliente = new System.Windows.Forms.TextBox();
            this.lblValorCheque = new System.Windows.Forms.Label();
            this.txtValorCheque = new System.Windows.Forms.TextBox();
            this.lblQuantidadeCheques = new System.Windows.Forms.Label();
            this.txtQuantidadeCheques = new System.Windows.Forms.TextBox();
            this.lblDataVencimento = new System.Windows.Forms.Label();
            this.dtpDataVencimento = new System.Windows.Forms.DateTimePicker();
            this.btnImprimir = new System.Windows.Forms.Button();
            this.btnLimpar = new System.Windows.Forms.Button();
            this.printDocument = new System.Drawing.Printing.PrintDocument();
            this.SuspendLayout();
            // 
            // lblNomeCliente
            // 
            this.lblNomeCliente.AutoSize = true;
            this.lblNomeCliente.Location = new System.Drawing.Point(20, 20);
            this.lblNomeCliente.Name = "lblNomeCliente";
            this.lblNomeCliente.Size = new System.Drawing.Size(100, 15);
            this.lblNomeCliente.TabIndex = 0;
            this.lblNomeCliente.Text = "Nome do Cliente:";
            // 
            // txtNomeCliente
            // 
            this.txtNomeCliente.Location = new System.Drawing.Point(20, 40);
            this.txtNomeCliente.Name = "txtNomeCliente";
            this.txtNomeCliente.Size = new System.Drawing.Size(350, 23);
            this.txtNomeCliente.TabIndex = 1;
            // 
            // lblValorCheque
            // 
            this.lblValorCheque.AutoSize = true;
            this.lblValorCheque.Location = new System.Drawing.Point(20, 80);
            this.lblValorCheque.Name = "lblValorCheque";
            this.lblValorCheque.Size = new System.Drawing.Size(100, 15);
            this.lblValorCheque.TabIndex = 2;
            this.lblValorCheque.Text = "Valor do Cheque:";
            // 
            // txtValorCheque
            // 
            this.txtValorCheque.Location = new System.Drawing.Point(20, 100);
            this.txtValorCheque.Name = "txtValorCheque";
            this.txtValorCheque.Size = new System.Drawing.Size(150, 23);
            this.txtValorCheque.TabIndex = 3;
            this.txtValorCheque.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValorCheque_KeyPress);
            this.txtValorCheque.Leave += new System.EventHandler(this.txtValorCheque_Leave);
            // 
            // lblQuantidadeCheques
            // 
            this.lblQuantidadeCheques.AutoSize = true;
            this.lblQuantidadeCheques.Location = new System.Drawing.Point(200, 80);
            this.lblQuantidadeCheques.Name = "lblQuantidadeCheques";
            this.lblQuantidadeCheques.Size = new System.Drawing.Size(145, 15);
            this.lblQuantidadeCheques.TabIndex = 4;
            this.lblQuantidadeCheques.Text = "Quantidade de Cheques:";
            // 
            // txtQuantidadeCheques
            // 
            this.txtQuantidadeCheques.Location = new System.Drawing.Point(200, 100);
            this.txtQuantidadeCheques.Name = "txtQuantidadeCheques";
            this.txtQuantidadeCheques.Size = new System.Drawing.Size(150, 23);
            this.txtQuantidadeCheques.TabIndex = 5;
            this.txtQuantidadeCheques.Text = "1";
            this.txtQuantidadeCheques.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtQuantidadeCheques_KeyPress);
            // 
            // lblDataVencimento
            // 
            this.lblDataVencimento.AutoSize = true;
            this.lblDataVencimento.Location = new System.Drawing.Point(20, 140);
            this.lblDataVencimento.Name = "lblDataVencimento";
            this.lblDataVencimento.Size = new System.Drawing.Size(130, 15);
            this.lblDataVencimento.TabIndex = 6;
            this.lblDataVencimento.Text = "Data de Vencimento:";
            // 
            // dtpDataVencimento
            // 
            this.dtpDataVencimento.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpDataVencimento.Location = new System.Drawing.Point(20, 160);
            this.dtpDataVencimento.Name = "dtpDataVencimento";
            this.dtpDataVencimento.Size = new System.Drawing.Size(150, 23);
            this.dtpDataVencimento.TabIndex = 7;
            // 
            // btnImprimir
            // 
            this.btnImprimir.Location = new System.Drawing.Point(20, 200);
            this.btnImprimir.Name = "btnImprimir";
            this.btnImprimir.Size = new System.Drawing.Size(150, 40);
            this.btnImprimir.TabIndex = 8;
            this.btnImprimir.Text = "Imprimir Cheque(s)";
            this.btnImprimir.UseVisualStyleBackColor = true;
            this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
            // 
            // btnLimpar
            // 
            this.btnLimpar.Location = new System.Drawing.Point(200, 200);
            this.btnLimpar.Name = "btnLimpar";
            this.btnLimpar.Size = new System.Drawing.Size(150, 40);
            this.btnLimpar.TabIndex = 9;
            this.btnLimpar.Text = "Limpar Campos";
            this.btnLimpar.UseVisualStyleBackColor = true;
            this.btnLimpar.Click += new System.EventHandler(this.btnLimpar_Click);
            // 
            // printDocument
            // 
            this.printDocument.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(this.printDocument_PrintPage);
            // 
            // FormImpressaoCheque
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(399, 261);
            this.Controls.Add(this.btnLimpar);
            this.Controls.Add(this.btnImprimir);
            this.Controls.Add(this.dtpDataVencimento);
            this.Controls.Add(this.lblDataVencimento);
            this.Controls.Add(this.txtQuantidadeCheques);
            this.Controls.Add(this.lblQuantidadeCheques);
            this.Controls.Add(this.txtValorCheque);
            this.Controls.Add(this.lblValorCheque);
            this.Controls.Add(this.txtNomeCliente);
            this.Controls.Add(this.lblNomeCliente);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "FormImpressaoCheque";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Impressão de Cheques";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private Label lblNomeCliente;
        private TextBox txtNomeCliente;
        private Label lblValorCheque;
        private TextBox txtValorCheque;
        private Label lblQuantidadeCheques;
        private TextBox txtQuantidadeCheques;
        private Label lblDataVencimento;
        private DateTimePicker dtpDataVencimento;
        private Button btnImprimir;
        private Button btnLimpar;
        private PrintDocument printDocument;

        private int chequesImpressos = 0;
        private int totalCheques = 0;

        private void txtValorCheque_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permite apenas números, vírgula e ponto para o campo de valor
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != ',' && e.KeyChar != '.')
            {
                e.Handled = true;
            }

            // Substitui ponto por vírgula
            if (e.KeyChar == '.')
            {
                e.KeyChar = ',';
            }

            // Permite apenas uma vírgula
            if (e.KeyChar == ',' && (sender as TextBox).Text.Contains(','))
            {
                e.Handled = true;
            }
        }

        private void txtValorCheque_Leave(object sender, EventArgs e)
        {
            // Formata o valor no padrão monetário brasileiro
            if (decimal.TryParse(txtValorCheque.Text, out decimal valor))
            {
                txtValorCheque.Text = string.Format(CultureInfo.GetCultureInfo("pt-BR"), "{0:N2}", valor);
            }
        }

        private void txtQuantidadeCheques_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permite apenas números inteiros para quantidade de cheques
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void btnLimpar_Click(object sender, EventArgs e)
        {
            // Limpa todos os campos
            txtNomeCliente.Text = string.Empty;
            txtValorCheque.Text = string.Empty;
            txtQuantidadeCheques.Text = "1";
            dtpDataVencimento.Value = DateTime.Today;
        }

        private void btnImprimir_Click(object sender, EventArgs e)
        {
            // Validações
            if (string.IsNullOrWhiteSpace(txtNomeCliente.Text))
            {
                MessageBox.Show("Por favor, informe o nome do cliente.", "Campo obrigatório", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtNomeCliente.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(txtValorCheque.Text) || !decimal.TryParse(txtValorCheque.Text.Replace(".", "").Replace(",", "."), out _))
            {
                MessageBox.Show("Por favor, informe um valor válido para o cheque.", "Campo obrigatório", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtValorCheque.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(txtQuantidadeCheques.Text) || !int.TryParse(txtQuantidadeCheques.Text, out totalCheques) || totalCheques <= 0)
            {
                MessageBox.Show("Por favor, informe uma quantidade válida de cheques.", "Campo obrigatório", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtQuantidadeCheques.Focus();
                return;
            }

            // Configura impressão
            PrintDialog printDialog = new PrintDialog();
            printDialog.Document = printDocument;
            printDialog.UseEXDialog = true;

            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                chequesImpressos = 0;
                printDocument.DocumentName = "Impressão de Cheques";
                printDocument.Print();
            }
        }

        private void printDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            // Fonte para impressão
            Font fonteTitulo = new Font("Courier New", 10, FontStyle.Bold);
            Font fonteNormal = new Font("Courier New", 10, FontStyle.Regular);
            Font fonteValor = new Font("Courier New", 12, FontStyle.Bold);

            // Posição inicial
            int posX = 10;
            int posY = 10;
            int altura = 250; // Altura aproximada de cada cheque

            // Dados do cheque
            string nomeCliente = txtNomeCliente.Text;
            decimal valorCheque = decimal.Parse(txtValorCheque.Text.Replace(".", "").Replace(",", "."), CultureInfo.InvariantCulture);
            string valorPorExtenso = ValorPorExtenso(valorCheque);
            string dataVencimento = dtpDataVencimento.Value.ToString("dd/MM/yyyy");

            // Desenha o cheque
            e.Graphics.DrawString("CHEQUE", fonteTitulo, Brushes.Black, posX, posY);
            posY += 20;

            e.Graphics.DrawString($"Pague por este cheque a: {nomeCliente}", fonteNormal, Brushes.Black, posX, posY);
            posY += 20;

            e.Graphics.DrawString($"A quantia de: R$ {valorCheque.ToString("N2", new CultureInfo("pt-BR"))}", fonteValor, Brushes.Black, posX, posY);
            posY += 20;

            // Dividir valor por extenso em linhas
            string[] linhasExtenso = DividirTexto(valorPorExtenso, 60);
            foreach (string linha in linhasExtenso)
            {
                e.Graphics.DrawString(linha, fonteNormal, Brushes.Black, posX, posY);
                posY += 15;
            }

            e.Graphics.DrawString($"Vencimento: {dataVencimento}", fonteNormal, Brushes.Black, posX, posY);
            posY += 30;

            // Incremente o contador de cheques impressos
            chequesImpressos++;

            // Verifica se há mais cheques para imprimir
            e.HasMorePages = chequesImpressos < totalCheques;
        }

        private string[] DividirTexto(string texto, int tamanhoMaximo)
        {
            // Divide o texto em linhas com tamanho máximo definido
            int numLinhas = (int)Math.Ceiling((double)texto.Length / tamanhoMaximo);
            string[] linhas = new string[numLinhas];

            for (int i = 0; i < numLinhas; i++)
            {
                int inicio = i * tamanhoMaximo;
                int tamanho = Math.Min(tamanhoMaximo, texto.Length - inicio);
                linhas[i] = texto.Substring(inicio, tamanho);
            }

            return linhas;
        }

        private string ValorPorExtenso(decimal valor)
        {
            if (valor == 0)
                return "Zero reais";

            string montagem = string.Empty;
            string valorPorExtenso = string.Empty;

            // Divisão para os centavos
            decimal valorInteiro = Math.Floor(valor);
            decimal centavos = Math.Round((valor - valorInteiro) * 100);

            if (valorInteiro > 0)
            {
                valorPorExtenso = NumerosPorExtenso(valorInteiro) + (valorInteiro == 1 ? " real" : " reais");
            }

            if (centavos > 0)
            {
                if (valorPorExtenso.Length > 0)
                    valorPorExtenso += " e ";

                valorPorExtenso += NumerosPorExtenso(centavos) + (centavos == 1 ? " centavo" : " centavos");
            }

            return valorPorExtenso;
        }

        private string NumerosPorExtenso(decimal numero)
        {
            if (numero == 0)
                return "zero";

            string[] unidades = { "", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove" };
            string[] dezADezenove = { "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove" };
            string[] dezenas = { "", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa" };
            string[] centenas = { "", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos" };
            string[] qualificadores = { "", "mil", "milhão", "bilhão", "trilhão", "quatrilhão" };
            string[] qualificadoresPlural = { "", "mil", "milhões", "bilhões", "trilhões", "quatrilhões" };

            string resultado = "";
            int inteiro = (int)Math.Floor(numero);

            if (inteiro == 0)
                return "zero";

            if (inteiro == 100)
                return "cem";

            int grupo = 0;

            do
            {
                int parte = inteiro % 1000;
                if (parte > 0)
                {
                    string montagem = "";

                    if (parte >= 100)
                    {
                        montagem += centenas[parte / 100];
                        parte %= 100;

                        if (parte > 0)
                            montagem += " e ";
                    }

                    if (parte >= 10 && parte < 20)
                    {
                        montagem += dezADezenove[parte - 10];
                        parte = 0;
                    }
                    else if (parte >= 20)
                    {
                        montagem += dezenas[parte / 10];
                        parte %= 10;

                        if (parte > 0)
                            montagem += " e ";
                    }

                    if (parte > 0)
                        montagem += unidades[parte];

                    if (grupo > 0)
                    {
                        if (inteiro % 1000 == 1)
                            montagem += " " + qualificadores[grupo];
                        else
                            montagem += " " + qualificadoresPlural[grupo];
                    }

                    if (resultado != "")
                        resultado = montagem + (montagem.EndsWith("mil") ? " " : " e ") + resultado;
                    else
                        resultado = montagem;
                }

                inteiro /= 1000;
                grupo++;
            }
            while (inteiro > 0);

            return resultado;
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormImpressaoCheque());
        }
    }
}