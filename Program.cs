using System;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.Windows.Forms;

namespace ChequePrinter
{
    public partial class FormCheque : Form
    {
        public FormCheque()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.lblNome = new Label();
            this.txtNome = new TextBox();
            this.lblValor = new Label();
            this.txtValor = new TextBox();
            this.lblQuantidade = new Label();
            this.txtQuantidade = new TextBox();
            this.lblData = new Label();
            this.dtpData = new DateTimePicker();
            this.btnImprimir = new Button();
            this.btnLimpar = new Button();
            this.SuspendLayout();

            // lblNome
            this.lblNome.Location = new Point(20, 20);
            this.lblNome.Size = new Size(150, 20);
            this.lblNome.Text = "Nome do Cliente:";

            // txtNome
            this.txtNome.Location = new Point(20, 45);
            this.txtNome.Size = new Size(350, 25);

            // lblValor
            this.lblValor.Location = new Point(20, 80);
            this.lblValor.Size = new Size(150, 20);
            this.lblValor.Text = "Valor do Cheque (R$):";

            // txtValor
            this.txtValor.Location = new Point(20, 105);
            this.txtValor.Size = new Size(150, 25);

            // lblQuantidade
            this.lblQuantidade.Location = new Point(200, 80);
            this.lblQuantidade.Size = new Size(150, 20);
            this.lblQuantidade.Text = "Quantidade de Cheques:";

            // txtQuantidade
            this.txtQuantidade.Location = new Point(200, 105);
            this.txtQuantidade.Size = new Size(80, 25);
            this.txtQuantidade.Text = "1";

            // lblData
            this.lblData.Location = new Point(20, 140);
            this.lblData.Size = new Size(150, 20);
            this.lblData.Text = "Data de Vencimento:";

            // dtpData
            this.dtpData.Location = new Point(20, 165);
            this.dtpData.Size = new Size(150, 25);
            this.dtpData.Format = DateTimePickerFormat.Short;

            // btnImprimir
            this.btnImprimir.Location = new Point(20, 210);
            this.btnImprimir.Size = new Size(150, 40);
            this.btnImprimir.Text = "Imprimir Cheque(s)";
            this.btnImprimir.Click += new EventHandler(this.btnImprimir_Click);

            // btnLimpar
            this.btnLimpar.Location = new Point(200, 210);
            this.btnLimpar.Size = new Size(150, 40);
            this.btnLimpar.Text = "Limpar Campos";
            this.btnLimpar.Click += new EventHandler(this.btnLimpar_Click);

            // FormCheque
            this.ClientSize = new Size(400, 270);
            this.Controls.Add(this.lblNome);
            this.Controls.Add(this.txtNome);
            this.Controls.Add(this.lblValor);
            this.Controls.Add(this.txtValor);
            this.Controls.Add(this.lblQuantidade);
            this.Controls.Add(this.txtQuantidade);
            this.Controls.Add(this.lblData);
            this.Controls.Add(this.dtpData);
            this.Controls.Add(this.btnImprimir);
            this.Controls.Add(this.btnLimpar);
            this.Name = "FormCheque";
            this.Text = "Sistema de Impress�o de Cheques";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private Label lblNome;
        private TextBox txtNome;
        private Label lblValor;
        private TextBox txtValor;
        private Label lblQuantidade;
        private TextBox txtQuantidade;
        private Label lblData;
        private DateTimePicker dtpData;
        private Button btnImprimir;
        private Button btnLimpar;

        private void btnLimpar_Click(object sender, EventArgs e)
        {
            txtNome.Text = "";
            txtValor.Text = "";
            txtQuantidade.Text = "1";
            dtpData.Value = DateTime.Today;
        }

        private void btnImprimir_Click(object sender, EventArgs e)
        {
            try
            {
                string nome = txtNome.Text.Trim();
                if (string.IsNullOrEmpty(nome))
                {
                    MessageBox.Show("Por favor, informe o nome do cliente.", "Campo Obrigat�rio", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtNome.Focus();
                    return;
                }

                decimal valor;
                if (!decimal.TryParse(txtValor.Text.Replace(".", ","), out valor) || valor <= 0)
                {
                    MessageBox.Show("Por favor, informe um valor v�lido para o cheque.", "Valor Inv�lido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtValor.Focus();
                    return;
                }

                int quantidade;
                if (!int.TryParse(txtQuantidade.Text, out quantidade) || quantidade <= 0)
                {
                    MessageBox.Show("Por favor, informe uma quantidade v�lida de cheques.", "Quantidade Inv�lida", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtQuantidade.Focus();
                    return;
                }

                DateTime dataVencimento = dtpData.Value;

                // Configurar a impress�o
                PrintDocument pd = new PrintDocument();
                pd.PrintPage += (s, ev) => ImprimirCheque(s, ev, nome, valor, dataVencimento, ref quantidade);
                pd.PrinterSettings.PrinterName = ObterNomeImpressoraMatricial();

                if (!pd.PrinterSettings.IsValid)
                {
                    MessageBox.Show("Impressora n�o encontrada ou inv�lida.", "Erro de Impressora", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                pd.Print();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao processar a impress�o: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string ObterNomeImpressoraMatricial()
        {
            // Idealmente, voc� implementaria uma l�gica para encontrar impressoras matriciais ou permitir sele��o
            // Em ambiente de produ��o, poderia ser armazenado em configura��es ou selecionado pelo usu�rio

            // Verificar impressoras instaladas
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                // Aqui voc� poderia verificar se � uma impressora matricial
                // Por enquanto, retorna a primeira encontrada ou a padr�o
                PrinterSettings settings = new PrinterSettings();
                settings.PrinterName = printer;

                if (settings.IsValid)
                {
                    return printer;
                }
            }

            // Usa a impressora padr�o se nenhuma for encontrada especificamente
            return new PrinterSettings().PrinterName;
        }

        private void ImprimirCheque(object sender, PrintPageEventArgs e, string nome, decimal valor, DateTime dataVencimento, ref int quantidade)
        {
            // Converte o valor para texto por extenso
            string valorPorExtenso = ValorPorExtenso(valor);

            // Formata��o do valor num�rico
            string valorFormatado = valor.ToString("N2", CultureInfo.CreateSpecificCulture("pt-BR"));

            // Fonte para impress�o matricial - usar fonte monoespa�ada
            Font fonte = new Font("Courier New", 10, FontStyle.Regular);
            Font fonteBold = new Font("Courier New", 10, FontStyle.Bold);

            Brush pincel = Brushes.Black;

            // Configura��es para posicionamento no cheque
            // Estas coordenadas devem ser ajustadas conforme o layout espec�fico do cheque
            int yBase = 50;

            // Imprime o nome do benefici�rio
            e.Graphics.DrawString($"c{nome}", fonte, pincel, 140, yBase + 100);

            // Imprime o valor num�rico
            e.Graphics.DrawString($"cR${valorFormatado}", fonteBold, pincel, 140, yBase);

            // Imprime o valor por extenso
            e.Graphics.DrawString($"b{valorPorExtenso.ToUpper()} ", fonte, pincel, 140, yBase + 30);

            // Imprime a data
            string cidade = "Vargem Alta -ES";
            string dia = dataVencimento.Day.ToString().PadLeft(2, ' ');
            string mes = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(dataVencimento.Month);
            string ano = dataVencimento.Year.ToString();

            e.Graphics.DrawString($"{cidade}", fonte, pincel, 500, yBase + 100);
            e.Graphics.DrawString($"{dia}", fonte, pincel, 650, yBase + 100);
            e.Graphics.DrawString($"{mes}", fonte, pincel, 680, yBase + 100);
            e.Graphics.DrawString($"{ano}", fonte, pincel, 800, yBase + 100);

            // Determina se h� mais p�ginas para imprimir
            quantidade--;
            e.HasMorePages = (quantidade > 0);
        }

        private string ValorPorExtenso(decimal valor)
        {
            if (valor == 0)
                return "zero reais";

            string montagem = "";
            if (valor < 0)
            {
                montagem = "menos ";
                valor = Math.Abs(valor);
            }

            if (valor < 1)
            {
                int centavos = (int)(valor * 100);
                if (centavos > 0)
                    montagem += ConverterCentavos(centavos);
            }
            else
            {
                int inteiro = (int)valor;
                int centavos = (int)((valor - inteiro) * 100);

                montagem += ConverterInteiros(inteiro);

                if (centavos > 0)
                {
                    montagem += " e " + ConverterCentavos(centavos);
                }
            }

            return montagem;
        }

        private string ConverterInteiros(int valor)
        {
            string[] unidades = { "", "um", "dois", "tr�s", "quatro", "cinco", "seis", "sete", "oito", "nove", "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove" };
            string[] dezenas = { "", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa" };
            string[] centenas = { "", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos" };

            string resultado = "";

            if (valor == 0)
                return "";

            if (valor == 1)
                return "um real";

            // Bilh�es
            int bilhoes = valor / 1000000000;
            valor %= 1000000000;

            if (bilhoes > 0)
            {
                if (bilhoes == 1)
                    resultado += "um bilh�o";
                else
                    resultado += ConverterCentenas(bilhoes) + " bilh�es";

                if (valor > 0)
                    resultado += " e ";
            }

            // Milh�es
            int milhoes = valor / 1000000;
            valor %= 1000000;

            if (milhoes > 0)
            {
                if (milhoes == 1)
                    resultado += "um milh�o";
                else
                    resultado += ConverterCentenas(milhoes) + " milh�es";

                if (valor > 0)
                    resultado += " e ";
            }

            // Milhares
            int milhares = valor / 1000;
            valor %= 1000;

            if (milhares > 0)
            {
                if (milhares == 1)
                    resultado += "mil";
                else
                    resultado += ConverterCentenas(milhares) + " mil";

                if (valor > 0)
                    resultado += " e ";
            }

            // Centenas
            if (valor > 0)
            {
                resultado += ConverterCentenas(valor);
            }

            return resultado + " reais";
        }

        private string ConverterCentenas(int valor)
        {
            string[] unidades = { "", "um", "dois", "tr�s", "quatro", "cinco", "seis", "sete", "oito", "nove", "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove" };
            string[] dezenas = { "", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa" };
            string[] centenas = { "", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos" };

            string resultado = "";

            if (valor == 0)
                return "";

            // Centena
            int centena = valor / 100;
            valor %= 100;

            if (centena > 0)
            {
                if (centena == 1 && valor == 0)
                    resultado += "cem";
                else
                    resultado += centenas[centena];

                if (valor > 0)
                    resultado += " e ";
            }

            // Dezena e unidade
            if (valor > 0)
            {
                if (valor < 20)
                {
                    resultado += unidades[valor];
                }
                else
                {
                    int dezena = valor / 10;
                    int unidade = valor % 10;

                    resultado += dezenas[dezena];

                    if (unidade > 0)
                        resultado += " e " + unidades[unidade];
                }
            }

            return resultado;
        }

        private string ConverterCentavos(int centavos)
        {
            if (centavos == 1)
                return "um centavo";

            string centavosExtenso = ConverterCentenas(centavos);
            return centavosExtenso.Replace("reais", "centavos");
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormCheque());
        }
    }
}