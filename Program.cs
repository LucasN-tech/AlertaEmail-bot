using DocumentFormat.OpenXml.Wordprocessing;
using alerta_Email;
using DocumentFormat.OpenXml.Bibliography;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Irony.Parsing;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;

namespace alerta_Email
{
    public class Program
    {
        static async Task Main(string[] args)
        {

            string filePath = \\Alerta_Email\\Pasta teste\\controle_clientes.xlsx";

            try
            {
                FileReader fileReader = new FileReader();
                XLWorkbook workbook = fileReader.ReadWorkbook(filePath);
                ConsultaBase consultaBase = new ConsultaBase(workbook);
                List<string> razoes = consultaBase.Razoes;
                
                var ano = DateTime.Now.Year;
                int.Parse(ano.ToString());
                var mounth = DateTime.Now.Month;
                int.Parse(mounth.ToString());
                mounth = (mounth - 1);
                if (mounth == 0) { mounth = 12; ano = ano - 01; }
                for (int i = 0; i < consultaBase.Cnpjs.Count; i++)
                {
                    string cnpj = consultaBase.Cnpjs[i];
                    string rede = consultaBase.CaminhoRede[i];

                    try
                    {
                        AnaliseDocumentos analiseDocumentos = new AnaliseDocumentos(cnpj, rede);                        
                        bool? resultado = analiseDocumentos.Resultado;
                        if (resultado == true) 
                        {
                            string accessToken = await MsGraph.GetAccessToken();
                            if (!string.IsNullOrEmpty(accessToken))
                            {
                                string subject = $"{consultaBase.Razoes[i]} | SOLICITAÇÃO DOCUMENTOS SUPORTE - {(string.Format("{0:00}", mounth))}/{ano} ";
                                string body = ($"Prezado(s)!{Environment.NewLine} {Environment.NewLine} Constatamos que não recebemos os documentos suporte (Extratos bancários, planilha de fluxo de caixa e demais) referente a {(string.Format("{0:00}", mounth))}/{ano} {Environment.NewLine}{Environment.NewLine}Pedimos que assim que possível, nos envie para prosseguirmos com as Apurações. {Environment.NewLine}{Environment.NewLine}Muito Obrigado!{Environment.NewLine}{Environment.NewLine} ******ESTE É UM EMAIL AUTOMÁTICO GERADO POR MEIO DE UM ALGORITIMO. GENTILEZA NÃO RESPONDÊ-LO!!! QUALQUER TENTATIVA DE CONTATO SERÁ SEM SUCESSO!!!");
                                string address = consultaBase.Emails1[i];
                                string address2 = consultaBase.Emails2[i];
                                string address3 = consultaBase.Emails3[i];
                                string address4 = consultaBase.Emails4[i];
                                var sendEmail = await MsGraph.SendEmail(subject, body, address, address2, address3, address4);
                            }
                            else
                            {
                                Console.WriteLine("Falha ao obter o token de acesso.");
                            }                           
                        }
                        else { Console.WriteLine($"Houve erro nao foi enviado{consultaBase.Razoes[i]}"); }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erro para CNPJ: {cnpj}. Mensagem de erro: {ex.Message}");
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro: " + ex.Message);
            }
        }
    }
}
