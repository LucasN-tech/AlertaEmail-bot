using ClosedXML.Excel;

namespace alerta_Email
{
    public class ConsultaBase
    {
        public List<string> Razoes { get; private set; }
        public List<string> Cnpjs { get; private set; }
        public List<string> CaminhoRede { get; private set; }
        public List<string> Emails1 { get; private set; }
        public List<string> Emails2 { get; private set; }
        public List<string> Emails3 { get; private set; }
        public List<string> Emails4 { get; private set; }

        public ConsultaBase(XLWorkbook workbook)
        {
            var planilha = workbook.Worksheets.First(w => w.Name == "Plan1");
            var totalLinhas = planilha.RangeUsed().RowCount();

            Razoes = new List<string>();
            Cnpjs = new List<string>();
            CaminhoRede = new List<string>();
            Emails1 = new List<string>();
            Emails2 = new List<string>();
            Emails3 = new List<string>();
            Emails4 = new List<string>();

            for (int l = 2; l <= totalLinhas; l++)
            {
                var cnpj = planilha.Cell($"B{l}").Value.ToString();
                if (!string.IsNullOrEmpty(cnpj))
                {
                    Razoes.Add(planilha.Cell($"A{l}").Value.ToString());
                    Cnpjs.Add(cnpj);
                    CaminhoRede.Add(planilha.Cell($"C{l}").Value.ToString());
                    Emails1.Add(planilha.Cell($"D{l}").Value.ToString());
                    Emails2.Add(planilha.Cell($"E{l}").Value.ToString());
                    Emails3.Add(planilha.Cell($"F{l}").Value.ToString());
                    Emails4.Add(planilha.Cell($"G{l}").Value.ToString());
                }
            }
        }
    }

        public class FileReader
    {
        public XLWorkbook ReadWorkbook(string filePath)
        {
            try
            {
                return new XLWorkbook(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro ao ler o arquivo: " + ex.Message);
                throw new Exception("Erro ao ler o arquivo.", ex);
            }
        }
    }

    public class AnaliseDocumentos
    {
        public bool? Resultado { get; set; }

        public AnaliseDocumentos(string cnpj, string rede)
        {
            if (string.IsNullOrEmpty(rede))
            {
                Resultado = false;
                Console.WriteLine("Não há arquivo");
                return;
            }

            try
            {
                var arquivos = Directory.GetFiles(rede);

                if (arquivos.Length > 0)
                {
                    Resultado = false;
                    Console.WriteLine("Tem arquivo");
                    return;
                }
                else
                {
                    Resultado = true;
                    Console.WriteLine("Não há arquivo");
                    return;
                }
            }
            catch (Exception ex)
            {
                Resultado = null;
                throw new Exception("Erro ao verificar a rede.", ex);
            }
        }
    }
}

