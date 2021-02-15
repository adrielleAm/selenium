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
    public class Unidade
    {
        public static DataTable dt = new DataTable();
        public static void  CadastrarUnidade(IWebDriver drive)
        {
            AbrirTelaUnidade(drive);

            dt = Arquivo.LerArquivoUnidade(drive);
           
            //if (dt == null)
            //    throw new Exception("Planilha vazia ou inválida.");

            foreach (DataRow item in dt.Rows)
            {
                CadastrarNovaUnidade(drive);

                TimeSpan tempoespera = TimeSpan.FromSeconds(10);
                IWait<IWebDriver> wait = new WebDriverWait(drive, tempoespera);
                //wait.Until(drv => drv.FindElement(By.Id("Celular")));
                wait.Until(drv => !Program.Exibir(drive, By.Id("bloqueio-tela")));

                PreencherDadosUnidade(drive, (string)item["NomeUnidade"], (string)item["CategoriaUnidade"], (string)item["TipoUnidade"],
                    (string)item["Situacao"], (string)item["Fornecedor"]);

                PreencherDadosUnidadeSuperios(drive, (string)item["SUPERIOR"], (string)item["TIPOSUP"]);

                //Selecionar Contrato
                IWebElement element = drive.FindElement(By.Id("7b85cd9f-db07-4eee-96e9-a52f353f08f7"));

                // IWebElement btn1 = drive.FindElement(By.Id)

                element.Click();

                Thread.Sleep(TimeSpan.FromSeconds(2));

                var btnSalvar = drive.FindElement(By.Id("btn_save"));
                btnSalvar.Click();
                //Selecionar Unidade

                drive.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(5));

                Arquivo.LogExecucao((string)item["NomeUnidade"], (string)item["CategoriaUnidade"], (string)item["Fornecedor"]);

            }

        }

        private static void PreencherDadosUnidadeSuperios(IWebDriver drive, string unidadeSup, string tipoUnidadeSup)
        {
            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownCategoriaSuperiorId = new SelectElement(drive.FindElement(By.Id("CategoriaSuperiorId")));
            dropdownCategoriaSuperiorId.SelectByText("OPERAÇÃO");

            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownTipoUnidadeSuperiorId = new SelectElement(drive.FindElement(By.Id("TipoUnidadeSuperiorId")));
            dropdownTipoUnidadeSuperiorId.SelectByText(tipoUnidadeSup);

            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownUnidadeSuperiorId = new SelectElement(drive.FindElement(By.Id("UnidadeSuperiorId")));
            dropdownUnidadeSuperiorId.SelectByText(unidadeSup);

            var btn1 = drive.FindElement(By.Id("btnAdicionar"));
            Thread.Sleep(TimeSpan.FromSeconds(2));
            btn1.Click();
        }

        private static void PreencherDadosUnidade(IWebDriver drive, string nome, string categoriaUnidade, string tipoUnidade, string situacao, string fornecedor)
        {
            IWebElement inputNome = drive.FindElement(By.Id("UnidadeNome"));
            inputNome.SendKeys(nome);

            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownCategoria = new SelectElement(drive.FindElement(By.Id("CategoriaId")));
            dropdownCategoria.SelectByText("PONTO DE AVALIAÇÃO");

            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownTipoUnidade = new SelectElement(drive.FindElement(By.Id("TipoUnidadeId")));
            dropdownTipoUnidade.SelectByText("COMERCIAL");

            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownSituacao = new SelectElement(drive.FindElement(By.Id("SituacaoId")));
            dropdownSituacao.SelectByText(situacao);

            //drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            //SelectElement dropdownPais = new SelectElement(drive.FindElement(By.Id("PaisId")));
            //dropdownPais.SelectByText(pais);

            //drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            //SelectElement dropdownUF = new SelectElement(drive.FindElement(By.Id("UFId")));
            //dropdownUF.SelectByText(uf);

            //if (municipio == null)
            //{
            //    drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            //    SelectElement dropdownMunicipio = new SelectElement(drive.FindElement(By.Id("MunicipioId")));
            //    dropdownMunicipio.SelectByText(fornecedor);

            //}

            drive.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
            SelectElement dropdownFornecedor = new SelectElement(drive.FindElement(By.Id("TerceirosId")));
            dropdownFornecedor.SelectByText(fornecedor);
        }
        
        private static void CadastrarNovaUnidade(IWebDriver drive)
        {
                TimeSpan tempoespera = TimeSpan.FromSeconds(25);

                IWait<IWebDriver> wait = new OpenQA.Selenium.Support.UI.WebDriverWait(drive, tempoespera);

                wait.Until(drv => drv.FindElement(By.XPath("//input[@value='Incluir']")));
                wait.Until(drv => !Program.Exibir(drive, By.Id("bloqueio-tela")));

                var btn = drive.FindElement(By.XPath("//input[@value='Incluir']"));
            //drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
            btn.Click();
        }
    
        public static void AbrirTelaUnidade(IWebDriver drive)
        {
            TimeSpan tempoespera = TimeSpan.FromSeconds(25);

            IWebElement lnkAdministracao = drive.FindElement(By.LinkText("Administração"));
            lnkAdministracao.Click();

            IWebElement lnkUsuario = drive.FindElement(By.LinkText("Unidades"));
            lnkUsuario.Click();

            drive.Manage().Timeouts().SetPageLoadTimeout(tempoespera);
        }
    }
}
