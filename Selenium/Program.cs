using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Net;

namespace AutomacaoCarga
{
    public class Program
    {
        static void Main(string[] args)
        {
            string contrato = null;
            int opcao;

            //https://sites.google.com/a/chromium.org/chromedriver/capabilities
            
            //ChromeOptions options = new ChromeOptions();
            //options.AddArguments("--start-fullscreen");

            string dirSelenium = @"D:\Projeto\Audtool\Selenium\Softwares\chromedriver_win32_75.0";

            IWebDriver drive = new ChromeDriver(dirSelenium);

            //Console.WriteLine("Inforweb o site:");
            //String site = Console.ReadLine();

            AbrirSite(drive);

            //Console.WriteLine("Informe o usuário:");
            //String usuario = Console.ReadLine();

            //Console.WriteLine("Informe a senha:");
            //String senha = Console.ReadLine();

            EfetuarLogin(drive);

            TimeSpan tempoespera = TimeSpan.FromSeconds(30);
            drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);


            opcao = EscolherRotina();

            switch (opcao)
            {
                case 1:
                    Usuario.CadastrarUsuario(drive);
                    break;

                case 2:
                    contrato = EscolherContrato(drive);
                    SelecionarContrato(drive, contrato);
                    TempoEspera(drive, tempoespera);
                    Pergunta.CadastrarPergunta(drive);
                    break;

                case 3:

                    Unidade.CadastrarUnidade(drive);
                    break;

                case 4:
                    contrato = EscolherContrato(drive);
                    SelecionarContrato(drive, contrato);
                    TempoEspera(drive, tempoespera);
                    Console.WriteLine("===================== Configurações do arquivo:=======================" +
                        "\n A data deve estar como texto;" +
                        "\n O ano e o numero do Ciclo devem esta como numero");
                    Ciclo.CadastraCiclo(drive);
                    break;
            }

        }

        private static void TempoEspera(IWebDriver drive, TimeSpan tempoespera)
        {
            IWait<IWebDriver> wait = new WebDriverWait(drive, tempoespera);
            wait.Until(drv => !Program.Exibir(drive, By.Id("bloqueio-tela")));
        }

        private static String EscolherContrato(IWebDriver drive)
        {
            int idContrato;
            var ambiente = drive.FindElements(By.ClassName("selecao_ambiente"));

            do
            {
                Console.WriteLine("Informe o número do contrato desejado:");

                for (int i = 0; i < ambiente.Count; i++)
                {
                    Console.WriteLine(string.Format("{0} - {1}", i, ambiente[i].Text));
                }
                idContrato = Int32.Parse(Console.ReadLine());

            } while (!(idContrato >= 0 && idContrato < ambiente.Count));

            return ambiente[idContrato].Text;
        }

        private static int EscolherRotina()
        {
            Console.WriteLine(" Informe a Rotina:\n" +
                   "1 - CADASTRAR USUÁRIOS \n" +
                   "2 - CADASTRAR PERGUNTAS \n" +
                   "3 - CADASTRAR UNIDADE \n" +
                   "4 - CADASTRAR CICLO");
            return Int32.Parse(Console.ReadLine());
        }

        private static void AbrirSite(IWebDriver drive)
        {
            //string site = "http://audtool.com";
            string site = "http://localhost:12042/Default";
            drive.Navigate().GoToUrl(site);
            //drive.Manage().Window.Maximize();
         
        }

        private static void AbrirSite(string site, IWebDriver drive)
        {
          
                //Cria requisição
                WebRequest request = WebRequest.Create(site);
                //Envia a requisição e recebe uma resposta
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            do {

                drive.Navigate().GoToUrl(site);
                drive.Manage().Window.Maximize();
                

            } while (!(response.StatusCode == HttpStatusCode.OK));
        }

        public static void ZoomIn(IWebDriver drive)
        {
            new Actions(drive)
                .SendKeys(Keys.Control + "+")
                .Perform();

        }

        public static void ZoomOut(IWebDriver drive)
        {
            new Actions(drive)
                .SendKeys(Keys.Control + "-")
                .Perform();
        }


        private static void EfetuarLogin(IWebDriver drive)
        {
            string usuario = "adrielle.araujo@inforweb.com.br";
            string senha = "Infor@123";
            //string senha = "audtool.12";
            //string senha = "abc.123";

            IWebElement inputLogin = drive.FindElement(By.Id("Email"));
            inputLogin.SendKeys(usuario);
            inputLogin.SendKeys(Keys.Tab);

            IWebElement inputPassword = drive.FindElement(By.Id("Senha"));
            inputPassword.SendKeys(senha);
            inputPassword.SendKeys(Keys.Enter);

            //TimeSpan tempoespera = new TimeSpan(0, 0, 2);

            //WebDriverWait wait = new WebDriverWait(drive, tempoespera);
            //wait.Until(ExpectedConditions.AlertIsPresent());
            //IAlert alert = drive.SwitchTo().Alert();
            //alert.Accept();

        }

        private static void EfetuarLogin(IWebDriver drive, string usuario, string senha)
        {
            IWebElement inputLogin = drive.FindElement(By.Id("Email"));
            inputLogin.SendKeys(usuario);
            inputLogin.SendKeys(Keys.Tab);

            IWebElement inputPassword = drive.FindElement(By.Id("Senha"));
            inputPassword.SendKeys(senha);
            inputPassword.SendKeys(Keys.Enter);

        }
        private static void SelecionarContrato(IWebDriver drive)
        {
            string contrato = "DPO/SPO - ABInbev";
            IWebElement lnkContrato = drive.FindElement(By.LinkText(contrato));
            lnkContrato.Click();
        }

        private static void SelecionarContrato(IWebDriver drive, String contrato)
        {
            IWebElement lnkContrato = drive.FindElement(By.LinkText(contrato));
            lnkContrato.Click();
        }

        public static bool VerificarExiste(IWebDriver driver, By locator)
        {
            try
            {
                driver.FindElement(locator);
                return true;
            }
            catch (NoSuchElementException) { return false; }
        }

        public static bool Exibir(IWebDriver driver, By locator)
        {
            try
            {
                IWebElement element1 = driver.FindElement(locator);
                return element1.Displayed;
            }
            catch (NoSuchElementException) { return false; }
        }

        public override string ToString()
        {
            return base.ToString();
        }

        public override bool Equals(object obj)
        {
            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
