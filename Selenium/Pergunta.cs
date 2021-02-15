using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Data;
using System.Threading;

namespace AutomacaoCarga
{
    public class Pergunta
    {
        public static void CadastrarPergunta(IWebDriver drive)
        {
            try {
                AbrirTelaPergunta(drive);

                drive.Manage().Timeouts();

                //drive.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(40));

                DataTable dtp = Arquivo.LerArquivoPerguntas(drive);
                if (dtp == null)
                    throw new Exception("Falha na Leitura do arquivo de perguntas");

                DataTable dta = Arquivo.LerArquivoAlternativas(drive, 1);
                if (dta == null)
                    throw new Exception("Falha na Leitura do arquivo de perguntas");




                foreach (DataRow item in dtp.Rows)
                {

                    string expression;
                    expression = String.Format("NumeroPergunta = '{0}'",item["NumeroPergunta"].ToString());
                    DataRow[] dtea = dta.Select(expression);
                    DataTable dtAlternativas = new DataTable();

                    dtAlternativas = dta.Clone();

                    foreach (DataRow row in dtea)
                    {
                        dtAlternativas.ImportRow(row);
                    }

                    CadastrarNovaPergunta(drive);

                    PreencherPergunta(drive, (int)item["Quant"], (string)item["Categoria"], (string)item["Modalidade"], item["FormaPreenchimento"].ToString(), item["NumeroPergunta"].ToString(),
                        (string)item["TituloPTBR"], (string)item["DescricaoPTBR"], (string)item["ItemVerificacaoPTBR"], (string)item["AplicacaoPadraoPTBR"], dtAlternativas);

                    //IWebElement msg = drive.FindElement(By.CssSelector("css=div.content"));
                    //msg.Click();
                    Thread.Sleep(TimeSpan.FromSeconds(5));
                }

            }
            catch (WebDriverException ex)
            {
                throw ex;
            }
        }

        private static void AbrirTelaPergunta(IWebDriver drive)
        {
            TimeSpan tempoespera = TimeSpan.FromSeconds(30);

            //IWebElement menuPlanejamento = drive.FindElement(By.XPath("//a[@id='rptMenu_hlkMenu_2']"));
            //drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
            //menuPlanejamento.Click();

            IWebElement menuPlanejamento = drive.FindElement(By.LinkText("Planejamento"));
            NewMethod(drive, tempoespera);
            menuPlanejamento.Click();

            //IWebElement menuPlanejamento = drive.FindElement(By.XPath("//a[@id='rptMenu_hlkMenu_0']"));
            //drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
            //menuPlanejamento.Click();

            IWebElement menuPergunta = drive.FindElement(By.LinkText("Perguntas"));
            drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
            menuPergunta.Click();

            /*IWebElement menuPergunta = drive.FindElement(By.XPath("//a[@id='rptMenu_rptSubMenu_2_hlkMenu_4']"));
            drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
            menuPergunta.Click();*/

            //IWebElement menuPergunta = drive.FindElement(By.XPath("//a[@id='rptMenu_rptSubMenu_0_hlkMenu_4']"));
            //menuPergunta.Click();
        }

        private static void NewMethod(IWebDriver drive, TimeSpan tempoespera)
        {
            drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
        }

        private static void CadastrarNovaPergunta(IWebDriver drive)
        {
            TimeSpan tempoespera = TimeSpan.FromSeconds(40);
            IWait<IWebDriver> wait = new OpenQA.Selenium.Support.UI.WebDriverWait(drive, tempoespera);

            wait.Until(drv => drv.FindElement(By.XPath("//input[@value='Nova Pergunta']")));
            wait.Until(drv => !Program.Exibir(drive, By.Id("bloqueio-tela")));

            var btn = drive.FindElement(By.XPath("//input[@value='Nova Pergunta']"));
            drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
            btn.Click();
        }

        private static void PreencherPergunta(IWebDriver drive, int quant, string categoria, string modalidade, string formaPreenchimento, string numeroPergunta,
            string titulo_PTBR, string descricao_PTBR, string itemVerificacao_PTBR, string aplicacaoPadrao_PTBR, DataTable dtAlternativa)
        //string titulo_EN, string descricao_EN, string itemVerificacao_EN, string aplicacaoPadrao_EN,
        //string titulo_ES, string descricao_ES, string itemVerificacao_ES, string aplicacaoPadrao_ES)
        {
            string separador = " - ";

            bool IncluirNA = true;

            TimeSpan tempoespera = TimeSpan.FromSeconds(40);
            IWait<IWebDriver> wait = new WebDriverWait(drive, tempoespera);
            wait.Until(drv => !Program.Exibir(drive, By.Id("bloqueio-tela")));


            if (string.IsNullOrEmpty(numeroPergunta))
                separador = string.Empty;

            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownCategoria = new SelectElement(drive.FindElement(By.Id("Call01_Conteudo_ddlCategoriaId")));
            dropdownCategoria.SelectByText(categoria);

            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownModalidade = new SelectElement(drive.FindElement(By.Id("Call01_Conteudo_ddlModalidadeId")));
            dropdownModalidade.SelectByText(modalidade);

            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownFormaPreenchimento = new SelectElement(drive.FindElement(By.Id("Call01_Conteudo_ddlFormaPreenchimentoId")));
            dropdownFormaPreenchimento.SelectByText(formaPreenchimento);

            //Tratamento para Caracteres Especiais
            titulo_PTBR = titulo_PTBR.Replace('\n', ' ');
            descricao_PTBR = descricao_PTBR.Replace('\n', ' ');//substituir por <br>
            itemVerificacao_PTBR = itemVerificacao_PTBR.Replace('\n', ' ');

            //string titulopergunta = String.Format("{0}{1}{2}", numeroPergunta, separador, titulo_PTBR);
            string titulopergunta = String.Format("{0}", titulo_PTBR);

            //drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromMilliseconds(10));

            //Preenchimento Português
            IWebElement tituloPTBR = drive.FindElement(By.Id("titulolang_1"));
            tituloPTBR.SendKeys(titulopergunta);

            IWebElement descricaoPTBR = drive.FindElement(By.Id("descricaolang_1"));
            descricaoPTBR.SendKeys(descricao_PTBR);

            IWebElement itemVerificacaoPTBR = drive.FindElement(By.Id("verificacaolang_1"));
            itemVerificacaoPTBR.SendKeys(titulopergunta);
            //itemVerificacaoPTBR.SendKeys(itemVerificacao_PTBR);

            IWebElement aplicabilidadePadraoPTBR = drive.FindElement(By.Id("aplicablang_1"));
            aplicabilidadePadraoPTBR.SendKeys(aplicacaoPadrao_PTBR);

            //Preenchimento Inglês
            IWebElement tituloEN = drive.FindElement(By.Id("titulolang_2"));
            tituloEN.SendKeys(titulopergunta);

            IWebElement descricaoEN = drive.FindElement(By.Id("descricaolang_2"));
            descricaoEN.SendKeys(descricao_PTBR);

            IWebElement itemVerificacaoEN = drive.FindElement(By.Id("verificacaolang_2"));
            itemVerificacaoEN.SendKeys(itemVerificacao_PTBR);

            IWebElement aplicabilidadeEN = drive.FindElement(By.Id("aplicablang_2"));
            aplicabilidadeEN.SendKeys(itemVerificacao_PTBR);


            //Preenchimento Espanhol
            IWebElement tituloES = drive.FindElement(By.Id("titulolang_3"));
            tituloES.SendKeys(titulopergunta);

            IWebElement descricaoES = drive.FindElement(By.Id("descricaolang_3"));
            descricaoES.SendKeys(descricao_PTBR);

            IWebElement itemVerificacaoES = drive.FindElement(By.Id("verificacaolang_3"));
            itemVerificacaoES.SendKeys(itemVerificacao_PTBR);

            IWebElement aplicabilidadeES = drive.FindElement(By.Id("aplicablang_3"));
            aplicabilidadeES.SendKeys(itemVerificacao_PTBR);

            string descricaoAlternativaPTBR = "Descrição Alternativa PTBR";
            string descricaoAlternativaEN = "Descrição Alternativa EN";
            string descricaoAlternativaES = "Descrição Alternativa ES";
            string peso = "1";
            string cor = string.Empty;

            Thread.Sleep(1000);

            DataTable dt = dtAlternativa;
                
                //Arquivo.LerArquivoAlternativas(drive, quant);

            //Assert Logar o sucesso da operação.

            bool despresarCalculo = false;

            int quantreg = dt.Rows.Count + 1;
            //DataRow dr;


            foreach (DataRow dr in dt.Rows) {
                var btn_novaalternativa = drive.FindElement(By.Id("btnIncluirAlternativa"));
                btn_novaalternativa.Click();

                descricaoAlternativaPTBR = dr["DescricaoPTBR"].ToString();
                descricaoAlternativaEN = dr["DescricaoEN"].ToString();
                descricaoAlternativaES = dr["DescricaoES"].ToString();
                peso = dr["peso"].ToString();
                cor = dr["cor"].ToString();


                drive.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(45));

                CadastrarAlternativas(drive, despresarCalculo, peso, cor, descricaoAlternativaPTBR,
                       descricaoAlternativaEN,
                      descricaoAlternativaES);
            }

            if (IncluirNA) {

                var btn_novaalternativa = drive.FindElement(By.Id("btnIncluirAlternativa"));
                btn_novaalternativa.Click();

                descricaoAlternativaPTBR = "N/A - não aplicável";
                descricaoAlternativaEN = "N/A - não aplicável";
                descricaoAlternativaES = "N/A - não aplicável";
                despresarCalculo = true;
                cor = "Branco";

                drive.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(45));

                CadastrarAlternativas(drive, despresarCalculo, peso, cor, descricaoAlternativaPTBR,
                       descricaoAlternativaEN,
                      descricaoAlternativaES);
            }

            //Cadastrar Alternativas
            //for (int i = 0; i < quantreg; i++)
            //{
            //    var btn_novaalternativa = drive.FindElement(By.Id("btnIncluirAlternativa"));
            //    btn_novaalternativa.Click();


            //    if (i == 2)
            //    {
            //        descricaoAlternativaPTBR = "N/A - não aplicável";
            //        descricaoAlternativaEN = "N/A - não aplicável";
            //        descricaoAlternativaES = "N/A - não aplicável";
            //        despresarCalculo = true;
            //    }
            //    else
            //    {
            //        dr = dt.Rows[i];
            //        descricaoAlternativaPTBR = dr["DescricaoPTBR"].ToString();
            //        descricaoAlternativaEN = dr["DescricaoEN"].ToString();
            //        descricaoAlternativaES = dr["DescricaoES"].ToString();
            //        peso = dr["peso"].ToString();
            //        cor = dr["cor"].ToString();
            //    }

            //    drive.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(45));

            //    CadastrarAlternativas(drive, despresarCalculo, peso, cor, descricaoAlternativaPTBR,
            //           descricaoAlternativaEN,
            //          descricaoAlternativaES);
            //}

            var btn_SalvarPergunta = drive.FindElement(By.Id("btnSalvarPergunta"));
            btn_SalvarPergunta.Click();

            Arquivo.LogExecucao(titulo_PTBR, categoria, numeroPergunta);
        }


        public static void AguardarPaginaCarregar(IWebDriver driver, TimeSpan tempoespera)
        {
            WebDriverWait wait = new WebDriverWait(driver, tempoespera);

            Func<IWebDriver, bool> waitForElement = new Func<IWebDriver, bool>((IWebDriver Web) =>
            {
                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                Boolean result = js.ExecuteScript("return document.readyState").Equals("complete");

                return result;
            });

            wait.Until(waitForElement);
        }


        private static void CadastrarAlternativas(IWebDriver drive, bool desprezarCalculo, string peso, string cor,
            string alternativaPTBR, string alternativaEN, string alternativaES)
        {
            TimeSpan tempoespera = TimeSpan.FromSeconds(40);
            IWait<IWebDriver> wait = new WebDriverWait(drive, tempoespera);
            wait.Until(drv => !Program.Exibir(drive, By.Id("bloqueio-tela")));

            IWait<IWebDriver> wait2 = new WebDriverWait(drive, tempoespera);
            //wait2.Until(drv => drv.FindElement(By.Name("Desprezar")));



            IWebElement chkdesprezarCalculo = wait2.Until(drv => drive.FindElement(By.Name("Desprezar")));

            //var chkdesprezarCalculo = drive.FindElement(By.Name("Desprezar"));

            if (desprezarCalculo)
            {
                chkdesprezarCalculo.Click();
            }
            else
            {
                IWebElement txtPeso = drive.FindElement(By.Name("Peso"));
                txtPeso.SendKeys(peso);
            }

            //drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            //SelectElement dropdownFormaPreenchimento = new SelectElement(drive.FindElement(By.Id("Call01_Conteudo_ddlFormaPreenchimentoId")));
            //dropdownFormaPreenchimento.SelectByText(cor);

            IWebElement txtalternativaPTBR = drive.FindElement(By.Name("1"));
            txtalternativaPTBR.SendKeys(alternativaPTBR);

            IWebElement txtalternativaEN = drive.FindElement(By.Name("2"));
            txtalternativaEN.SendKeys(alternativaPTBR);

            IWebElement txtalternativaES = drive.FindElement(By.Name("3"));
            txtalternativaES.SendKeys(alternativaPTBR);

            var btn_SalvarAlternativa = drive.FindElement(By.Id("btnSalvarOpcao"));
            btn_SalvarAlternativa.Click();
        }
    }
}
