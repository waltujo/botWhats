using ClosedXML.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace BoWhatsMessage
{
    public partial class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.Title = "Automação WhatsApp";
                Console.ForegroundColor = ConsoleColor.Green;

                var nomeUser = Environment.UserName;
                //LEMBRAR DE CONFIGURAR O CAMINHO DO ARQUIVO EXCEL
                var pathExcel = $@"C:\Users\{nomeUser}\Contatos\Contatos.xlsx";
                var pathMensagem = $@"C:\Users\{nomeUser}\Contatos\mensagem.txt";
                
                Console.WriteLine("Fechando instâncias do navegador Chrome.");
                FecharInstanciasNavegador();

                //DIGITAR O NOME DO ARQUIVO QUE DESEJA ENVIAR, EX: FOTO.MP4
                Console.WriteLine("DIGITAR O NOME DO ARQUIVO QUE DESEJA ENVIAR, EX: FOTO.MP4");
                var arquivo = Console.ReadLine();

                //Lista de contatos
                var contatos = ExtrairNumerosContatos(pathExcel);
                contatos = contatos.Select(c => c.Trim()).Distinct().Where(c => !string.IsNullOrEmpty(c)).ToList();

                var telefone = CorrigirNumerosTelefone(contatos);
                //Mensagem que vai ser enviada.

                var mensagem = File.ReadAllText(pathMensagem).Replace("\r\n", "");

                ChromeOptions options = new();
                options.AddArgument("--start-maximized");
                options.AddArguments("--disable-extensions");
                options.AddArgument(@$"--user-data-dir=C:\Users\{nomeUser}\AppData\Local\Google\Chrome\User Data\Default"); // Altere para o caminho do seu perfil do Chrome
                options.PageLoadStrategy = PageLoadStrategy.Normal;

                var service = ChromeDriverService.CreateDefaultService();
                service.SuppressInitialDiagnosticInformation = false;
                service.DisableBuildCheck = false;
                service.EnableVerboseLogging = false;
                service.HideCommandPromptWindow = false;

                using (var driver = new ChromeDriver(service, options))
                {
                    try
                    {
                        Console.Clear();
                        // Abre o WhatsApp Web
                        driver.Navigate().GoToUrl("https://web.whatsapp.com");

                        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

                        // Aguarda o usuário fazer o login
                        // Verificar se está na tela de escaneamento do QR code
                        if (ScanQrCode(wait))
                        {
                            Console.WriteLine("Por favor, escaneie o QR code.");
                            Console.WriteLine("Após escanear o QR COde pressionar qualquer tecla!");
                            Console.ReadLine();
                            Console.Clear();
                        }


                        foreach (var contato in telefone)
                        {
                            Console.ForegroundColor = ConsoleColor.Blue;
                            Console.WriteLine($"{telefone.IndexOf(contato) + 1} / {telefone.Count} == {contato}");

                            EnviarMensagem(driver, wait, contato, mensagem, nomeUser, arquivo);
                        }
                        Console.WriteLine("MENSAGENS ENVIADAS COM SUCESSO!! -- APERTE QUALQUER TECLA PRA FINALIZAR!");
                        Console.ReadLine();
                    }
                    catch (Exception ex)
                    {
                        Logger($"Erro: {ex.InnerException} + {ex.Message}");
                        throw ex.InnerException;
                    }
                    finally
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                Logger($"Erro: {ex.InnerException} + {ex.Message}");
                throw ex.InnerException;
            }
            finally
            {
                Console.WriteLine("Programa finalizado!");
                Thread.Sleep(3000);
            }
        }
        static void EnviarMensagem(ChromeDriver driver, WebDriverWait wait, string telefone, string mensagem, string nomeUser, string arquivo)
        {
            var link = $"https://web.whatsapp.com/send?phone={telefone}&text={mensagem}";

            try
            {
                driver.Navigate().GoToUrl(link);

                HandlePopup(driver, alert => alert.Accept());

                Thread.Sleep(1500);

                //PEGA O BOTÃO + 
                var attachFile = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//span[@data-icon='plus']")));
                attachFile.Click();

                //var inputFile = driver.FindElement(By.XPath("//input[@type='file']"));
                //inputFile.SendKeys($@"C:\Users\{nomeUser}\Contatos\teste.png");

                Thread.Sleep(1000);

                //ADICIONA O ARQUIVO -- //*[@id="app"]/div/span[6]/div/ul/div/div/div[2]/li/div/input
                var inputFile = driver.FindElement(By.XPath("//*[@id='app']/div/span[6]/div/ul/div/div/div[2]/li/div/input"));

                //var inputFile = driver.FindElement(By.XPath("//*[@id='main']/footer/div[1]/div/span/div/div[1]/div[2]/div/span/div/ul/div/div[2]/li/div/input"));

                inputFile.SendKeys($@"C:\Users\{nomeUser}\Contatos\{arquivo}");

                Thread.Sleep(3000);

                //BOTÃO DE ENVIAR
                IWebElement btnEnviar = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//span[@data-icon='send']")));

                Thread.Sleep(2000);

                btnEnviar.Click();

                Thread.Sleep(1000);

                //IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                //js.ExecuteScript("arguments[0].click();", btnEnviar);

                //var btnSend = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//span[@data-icon='send']")));
                //btnSend.Click();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Número inválido == {telefone}");
                Logger($"Erro: {ex.InnerException} + {ex.Message}");
                // Se ocorrer um erro, vá para o próximo número da lista
                return;
            }

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Mensagem enviada para {telefone}");
        }
        public static void HandlePopup(IWebDriver driver, Action<IAlert> action)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                wait.Until(ExpectedConditions.AlertIsPresent());

                var alert = driver.SwitchTo().Alert();

                try
                {
                    action(alert); // Executa a ação desejada no pop-up
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erro ao lidar com o pop-up: {ex.Message}");
                }

                driver.SwitchTo().DefaultContent(); // Volta para o contexto padrão da página

                // Confirma ou cancela o pop-up após a ação
                // alert.Accept(); // Para confirmar o pop-up
                // alert.Dismiss(); // Para cancelar o pop-up
            }
            catch (NoAlertPresentException) { }
            catch (WebDriverTimeoutException) { }
            catch (Exception ex)
            {
                // Outro tipo de exceção
                Console.WriteLine($"Erro ao lidar com o pop-up: {ex.Message}");
            }
        }
        static List<string> CorrigirNumerosTelefone(List<string> contatos)
        {
            var contatosCorrigidos = new HashSet<string>(); // Evita duplicatas automaticamente

            foreach (var contato in contatos.Select(c => c.Trim()))
            {
                string numeroLimpo = Regex.Replace(contato, @"\D", ""); // Remove caracteres não numéricos

                if (numeroLimpo.Length == 10 || numeroLimpo.Length == 11) // Números sem DDD internacional
                {
                    numeroLimpo = "55" + numeroLimpo; // Adiciona o código do Brasil
                }

                // Verifica se o número já tem o formato correto (10 ou 11 dígitos após 55)
                if (Regex.IsMatch(numeroLimpo, @"^55\d{10,11}$")) 
                {
                    // Se for um celular sem o "9", adiciona o dígito antes do número
                    if (Regex.IsMatch(numeroLimpo, @"^55\d{2}[2-8]\d{7}$")) // Verifica se falta o "9"
                    {
                        numeroLimpo = numeroLimpo.Substring(0, 4) + "9" + numeroLimpo.Substring(4);
                    }

                    contatosCorrigidos.Add($"+{numeroLimpo.Substring(0, 2)} {numeroLimpo.Substring(2, 2)} {numeroLimpo.Substring(4, 5)}-{numeroLimpo.Substring(9)}");
                }
            }

            return contatosCorrigidos.ToList();
        }
        static bool ScanQrCode(WebDriverWait wait)
        {
            try
            {
                // Verifica se o elemento que indica a tela de escaneamento do QR code está visível
                var qrcode = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//canvas[@aria-label='Scan this QR code to link a device!']")));

                return qrcode.Displayed; // Retorna true se o QR code estiver sendo exibido
            }
            catch (WebDriverTimeoutException)
            {
                // Se ocorrer um timeout, significa que já está na tela das conversas
                return false;
            }
            catch (NoSuchElementException)
            {
                // Se o elemento não for encontrado, também significa que já está na tela das conversas
                return false;
            }
        }
        static List<string> ExtrairNumerosContatos(string caminhoArquivo)
        {
            using var workbook = new XLWorkbook(caminhoArquivo);
            var sheet = workbook.Worksheet(1);
            return sheet.RowsUsed().Select(row => row.Cell(1).GetString()).ToList();
        }
        public static void Logger(string mensagem)
        {
            string userProfilepath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
            string logDirectory = Path.Combine(userProfilepath, "Contatos\\log");
            string logFilePath = Path.Combine(logDirectory, $"LOG-{DateTime.Now:dd-MM-yy}.txt");

            try
            {
                if (!Directory.Exists(logDirectory)) 
                {
                    Directory.CreateDirectory(logDirectory);
                }

                using (StreamWriter writer = new StreamWriter(logFilePath, false))
                {
                    writer.WriteLine($"{DateTime.Now:dd-MM-yyyy HH:mm:ss} - {mensagem}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao registrar Log: {ex.Message}");
                throw;
            }
        }
        static void FecharInstanciasNavegador()
        {
            Console.WriteLine("Fechando instâncias do Google Chrome...");
            try
            {
                var processos = System.Diagnostics.Process.GetProcessesByName("chrome")
                    .Where(p => p.MainWindowTitle.Contains("WhatsApp"));

                foreach (var process in processos)
                {
                    try
                    {
                        process.Kill();
                        process.WaitForExit();
                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}