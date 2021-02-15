using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Data;

namespace AutomacaoCarga
{
    public class Ciclo
    {
        public static void CadastraCiclo(IWebDriver drive)
        {
            AbrirTelaCilco(drive);

            DataTable dt = Arquivo.LerArquivoCiclo(drive);

            if (dt == null)
                throw new Exception("Planilha vazia ou inválida.");

            foreach (DataRow item in dt.Rows)
            {
                CadastrarNovCiclo(drive);

                TimeSpan tempoespera = TimeSpan.FromSeconds(20);
                IWait<IWebDriver> wait = new WebDriverWait(drive, tempoespera);
                //wait.Until(drv => drv.FindElement(By.Id("Celular")));
                wait.Until(drv => !Program.Exibir(drive, By.Id("bloqueio-tela")));

                              
                PreencherDadosCiclo(drive, (string)item["DescricaoPt"], (string)item["DescricaoEn"], (string)item["DescricaoEs"],
                    (string)item["Pilar"], (double)item["NCiclo"], (double)item["Ano"], (string)item["GrupoMeta"], Convert.ToDateTime((string)item["DataInicio"]), Convert.ToDateTime((string)item["DataFim"]), (string)item["Questionario"]);


                var btnSalvar = drive.FindElement(By.ClassName("btnSalvar"));
                btnSalvar.Click();
                //Selecionar Unidade

                drive.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(20));

                //Arquivo.LogExecucao((string)item["DescricaoPt"], (string)item["Pilar"], (string)item["Questionario"]);

            }

        }

        private static void PreencherDadosCiclo(IWebDriver drive, string DescricaoPt, string DescricaoEn, string DescricaoEs, string Pilar, double NCiclo, double Ano, string GrupoMeta, DateTime DataInicio, DateTime DataFim, string Questionario)
        {
            IWebElement inputDescricaoPt = drive.FindElement(By.Name("1"));
            inputDescricaoPt.SendKeys(DescricaoPt);

            IWebElement inputDescricaoEn = drive.FindElement(By.Name("2"));
            inputDescricaoEn.SendKeys(DescricaoEn);

            IWebElement inputDescricaoEs = drive.FindElement(By.Name("3"));
            inputDescricaoEs.SendKeys(DescricaoEs);

            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownPilar = new SelectElement(drive.FindElement(By.Id("ddlCategoriaQuestionario")));
            dropdownPilar.SelectByText(Pilar);


            IWebElement inputNCiclo = drive.FindElement(By.Name("txtnumero"));
            inputNCiclo.Clear();
            inputNCiclo.SendKeys("0" + Convert.ToString(NCiclo));

            IWebElement inputAno = drive.FindElement(By.Name("txtano"));
            inputAno.Clear();
            inputAno.SendKeys(Convert.ToString(Ano));

            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownGrupoMeta = new SelectElement(drive.FindElement(By.Id("ddlGrupoMeta")));
            dropdownGrupoMeta.SelectByText(GrupoMeta);

            IWebElement inputDataInicio = drive.FindElement(By.Id("txtdataInicio"));
            inputDataInicio.Clear();
            inputDataInicio.SendKeys(Convert.ToString(DataInicio));
            inputDataInicio.SendKeys(Keys.Tab);

            IWebElement inputDataFim = drive.FindElement(By.Id("txtDataFim"));
            inputDataFim.Clear();
            inputDataFim.SendKeys(Convert.ToString(DataFim));
            inputDataFim.SendKeys(Keys.Tab);


            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownQuestionario = new SelectElement(drive.FindElement(By.Name("ddlQuestionario")));
            dropdownQuestionario.SelectByText(Questionario);



        }

        private static void CadastrarNovCiclo(IWebDriver drive)
        {
            TimeSpan tempoespera = TimeSpan.FromSeconds(20);

            IWait<IWebDriver> wait = new OpenQA.Selenium.Support.UI.WebDriverWait(drive, tempoespera);

            wait.Until(drv => drv.FindElement(By.Id("novoCiclo")));
            wait.Until(drv => !Program.Exibir(drive, By.Id("bloqueio-tela")));

            var btn = drive.FindElement(By.Id("novoCiclo"));
            //drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
            btn.Click();
        }

        private static void AbrirTelaCilco(IWebDriver drive)
        {
            TimeSpan tempoespera = TimeSpan.FromSeconds(25);


            IWebElement menuPlanejamento = drive.FindElement(By.LinkText("Planejamento"));
            drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
            menuPlanejamento.Click();

            IWebElement menuCiclo = drive.FindElement(By.LinkText("Ano/Ciclo"));
            drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
            menuCiclo.Click();

        }

        //private static void NewMethod(IWebDriver drive, TimeSpan tempoespera)
        //{
        //    drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
        //}
    }
}
