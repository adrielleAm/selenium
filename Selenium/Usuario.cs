using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;

namespace AutomacaoCarga
{
    public class Usuario
    {
        public static void CadastrarUsuario(IWebDriver drive)
        {
            AbrirTelaUsuario(drive);

            DataTable dt = Arquivo.LerArquivoUsuarios(drive);

            if (dt == null)
                throw new Exception("Planilha vazia ou inválida.");

            foreach (DataRow item in dt.Rows)
            {
                CadastrarNovoUsuario(drive);

                TimeSpan tempoespera = TimeSpan.FromSeconds(10);
                IWait<IWebDriver> wait = new WebDriverWait(drive, tempoespera);
                //wait.Until(drv => drv.FindElement(By.Id("Celular")));
                wait.Until(drv => !Program.Exibir(drive, By.Id("bloqueio-tela")));

                IWebElement myDynamicElement = wait.Until(d => d.FindElement(By.Id("Celular")));

                PreencherDadosUsuario(drive, (string)item["Nome"], (string)item["Email"], (string)item["Perfil"],
                    (string)item["Telefone"], (string)item["Celular"], (string)item["Senha"], (DateTime)item["DataNascimento"]);

                PreencherDadosEmpresa(drive, (string)item["EmpresaLotacao"], (string)item["UnidadeLotacao"]);

                //Selecionar Contrato
                //var element = drive.FindElement(By.CssSelector("css=td.col-md-1.text-center > input[type='checkbox']"));

                var btn1 = drive.FindElement(By.Id("589eaa0a-2201-42ae-91c6-84e988763cfd"));
                btn1.Click();

                Thread.Sleep(TimeSpan.FromSeconds(2));

                var btnSalvar = drive.FindElement(By.Id("btn_save"));
                btnSalvar.Click();
                //Selecionar Unidade
                
                drive.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(10));

                Arquivo.LogExecucao((string)item["Nome"], (string)item["Email"], (string)item["Linha"]);

            }

        }

        private static void PreencherDadosEmpresa(IWebDriver drive, string empresa, string unidade)
        {
            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownEmpresa = new SelectElement(drive.FindElement(By.Id("ddlEmpresaLotacao")));

            var emp = dropdownEmpresa.Options.Where(x => x.Text.ToLower().IndexOf(empresa.ToLower()) > -1);
            dropdownEmpresa.SelectByText(emp.FirstOrDefault().Text);


            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownUnidade = new SelectElement(drive.FindElement(By.Id("ddlUnidadeLotacao")));

            var und = dropdownUnidade.Options.Where(x => x.Text.ToLower().IndexOf(unidade.ToLower()) > -1);
            dropdownUnidade.SelectByText(und.FirstOrDefault().Text);
        }
        
        private static void PreencherDadosUsuario(IWebDriver drive, string nome, string email, string perfil,
            string telefone, string celular, string senha, DateTime datanascimento)
        {
                        
            IWebElement inputNome = drive.FindElement(By.Id("Nome"));
            inputNome.SendKeys(nome);

            IWebElement inputEmail = drive.FindElement(By.Id("Email"));
            inputEmail.SendKeys(email.ToLower());

            IWebElement inputTelefone = drive.FindElement(By.Id("Telefone"));
            inputTelefone.SendKeys(telefone);
            inputTelefone.SendKeys(Keys.Tab);

            IWebElement inputCelular = drive.FindElement(By.Id("Celular"));
            inputCelular.SendKeys(telefone);

            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownPerfil = new SelectElement(drive.FindElement(By.Id("PerfilId")));
            dropdownPerfil.SelectByText(perfil);

            IWebElement inputSenha = drive.FindElement(By.Id("Senha"));
            inputSenha.SendKeys(senha);

            IWebElement inputDataNascimento = drive.FindElement(By.Id("DtNascimento"));
            inputDataNascimento.Clear();
            inputDataNascimento.SendKeys("10-10-2016");
            inputDataNascimento.SendKeys(Keys.Tab);
        }

        private static void CadastrarNovoUsuario(IWebDriver drive)
        {
            TimeSpan tempoespera = TimeSpan.FromSeconds(15);

            IWait<IWebDriver> wait = new OpenQA.Selenium.Support.UI.WebDriverWait(drive, tempoespera);

            wait.Until(drv => drv.FindElement(By.XPath("//input[@value='Incluir']")));
            wait.Until(drv => !Program.Exibir(drive, By.Id("bloqueio-tela")));

            var btn = drive.FindElement(By.XPath("//input[@value='Incluir']"));
            //drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
            btn.Click();
        }

        private static void AbrirTelaUsuario(IWebDriver drive)
        {
            TimeSpan tempoespera = TimeSpan.FromSeconds(25);

            IWebElement lnkAdministracao = drive.FindElement(By.LinkText("Administração"));
            lnkAdministracao.Click();

            IWebElement lnkUsuario = drive.FindElement(By.LinkText("Usuários"));
            lnkUsuario.Click();

            drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
        }

    }
}
