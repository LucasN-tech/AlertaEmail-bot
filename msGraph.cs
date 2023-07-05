using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;


public class MsGraph
{
    public class AppSettings
    {
        public string ClientId { get; set; }
        public string SecretId { get; set; }
        public string TenantId { get; set; }
        public string PrincipalEmail { get; set; }
        public string Token { get; set; }
    }

    public static (string clientId, string secretId, string tenantId, string principalEmail, string token) GetSettings()
    {
        try
        {
            string json = File.ReadAllText(".\\settings.json");
            //Console.WriteLine("Arquivo lido com sucesso!");
            //Console.WriteLine(json);

            var settingsObject = JsonConvert.DeserializeObject<JObject>(json);
            var appSettings = settingsObject["AppSettings"].ToObject<AppSettings>();

            string clientId = appSettings.ClientId;
            string secretId = appSettings.SecretId;
            string tenantId = appSettings.TenantId;
            string principalEmail = appSettings.PrincipalEmail;
            string token = appSettings.Token;

            return (clientId, secretId, tenantId, principalEmail, token);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Erro ao ler o arquivo:");
            Console.WriteLine(ex.Message);
        }

        return (null, null, null, null, null);
    }

    public static void PrintSettings()
    {
        var settings = GetSettings();

        Console.WriteLine("ClientId: " + settings.clientId);
        Console.WriteLine("SecretId: " + settings.secretId);
        Console.WriteLine("TenantId: " + settings.tenantId);
        Console.WriteLine("PrincipalEmail: " + settings.principalEmail);
        Console.WriteLine("Token: " + settings.token);
    }

    public static async Task<string> GetAccessToken()
    {
        var (clientId, secretId, tenantId, principalEmail, token) = GetSettings();

        using (var httpClient = new HttpClient())
        {
            var requestBody = $"client_id={clientId}&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&client_secret={secretId}&grant_type=client_credentials";
            var content = new StringContent(requestBody, Encoding.UTF8, "application/x-www-form-urlencoded");
            var response = await httpClient.PostAsync($"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token", content);

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();
                var newToken = result.Substring(result.IndexOf("access_token\":\"") + 15);
                newToken = newToken.Substring(0, newToken.IndexOf("\""));

                // Atualizar o token no arquivo JSON
                UpdateTokenInSettingsFile(newToken);

                return newToken;
            }

            //Console.WriteLine("Failed to retrieve access token.");
            return null;
        }
    }

    private static void UpdateTokenInSettingsFile(string newToken)
    {
        try
        {
            string json = File.ReadAllText(".\\settings.json");
            var settingsObject = JsonConvert.DeserializeObject<JObject>(json);
            var appSettings = settingsObject["AppSettings"].ToObject<AppSettings>();
            appSettings.Token = newToken;
            settingsObject["AppSettings"] = JObject.FromObject(appSettings);
            string updatedJson = JsonConvert.SerializeObject(settingsObject, Formatting.Indented);
            File.WriteAllText(".\\settings.json", updatedJson);
            Console.WriteLine("Token atualizado com sucesso no arquivo JSON.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Erro ao atualizar o token no arquivo JSON:");
            Console.WriteLine(ex.Message);
        }
    }

    public static async Task<HttpResponseMessage> GetUsers()
    {

        var settings = GetSettings();

        using (var httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", settings.token);
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            try
            {
                var url = "https://graph.microsoft.com/v1.0/users";
                var response = await httpClient.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    Console.WriteLine(json);
                    return response;
                }
                else
                {
                    Console.WriteLine("Failed to retrieve users. Status code: " + response.StatusCode);
                    return null;
                }
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine("Error in HTTP request: " + ex.Message);
                return null;
            }
        }
    }

    public static async Task<bool> ValidateToken()
    {
        var settings = GetSettings();
        using (var httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", settings.token);
            var response = await httpClient.GetAsync("https://login.microsoftonline.com/common/oauth2/nativeclient"); // Substitua com a URL de validação desejada

            if (response.IsSuccessStatusCode)
            {
                Console.WriteLine("Atividade ValidateToken: Token is valid.");
                return true;
            }

            Console.WriteLine("Atividade ValidateToken: Token is invalid.");
            return false;
        }
    }


    public static async Task<HttpResponseMessage> SendEmail(string subject, string body, string address, string address2, string address3, string address4)
    {
        var settings = GetSettings();
        using (var httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", settings.token);
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var url = $"https://graph.microsoft.com/v1.0/users/{settings.principalEmail}/sendMail";

            var email = new
            {
                message = new
                {
                    subject = subject,
                    body = new
                    {
                        contentType = "Text",
                        content = body
                    },
                    toRecipients = new[]
                    {
                    new
                    {
                        emailAddress = new
                        {
                            address = address
                        }
                    },
                },
                    ccRecipients = new List<object>()
                }
            };

            if (!string.IsNullOrEmpty(address2))
            {
                email.message.ccRecipients.Add(new
                {
                    emailAddress = new
                    {
                        address = address2
                    }
                });
            }

            if (!string.IsNullOrEmpty(address3))
            {
                email.message.ccRecipients.Add(new
                {
                    emailAddress = new
                    {
                        address = address3
                    }
                });
            }

            if (!string.IsNullOrEmpty(address4))
            {
                email.message.ccRecipients.Add(new
                {
                    emailAddress = new
                    {
                        address = address4
                    }
                });
            }

            var jsonEmail = JsonConvert.SerializeObject(email);
            var content = new StringContent(jsonEmail, Encoding.UTF8, "application/json");

            try
            {
                var response = await httpClient.PostAsync(url, content);
                Console.WriteLine("Enviado: ");
                return response;
            }
            catch (HttpRequestException ex)
            {
                Console.WriteLine("Erro na requisição HTTP: " + ex.Message);
                return null;
            }
        }
    }
}
