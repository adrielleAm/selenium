using NPOI.HSSF.Model;
using NPOI.HSSF.UserModel;
using OpenQA.Selenium;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace AutomacaoCarga
{
    public class Arquivo
    {
        public static String EscolherArquivo()
        {
            string diretorioOrigem = @"D:\DEMARCO\Carga\SPO_H2_REVENDA";

            string[] arquivos = Directory.GetFiles(diretorioOrigem, "*.xls", SearchOption.TopDirectoryOnly);

            Console.WriteLine("Informe o arquivo para carga:" +
                "\n=========== Arquivos Disponíveis ===========");

            for (int i = 0; i < arquivos.Length; i++)
            {
                Console.WriteLine("{0} - {1}", i, arquivos[i]);
            }

            int arquivo = Int32.Parse(Console.ReadLine());

            String diretorio = arquivos[arquivo];

            return diretorio;

        }

        internal static DataTable LerArquivoUnidade(IWebDriver drive)
        {

            string caminhoArquivo = EscolherArquivo();

            //string caminhoArquivo = @"D:\CARGA\CicloDash.xls";

            DataSet planilha = CarregarArquivo(drive, caminhoArquivo);
            return planilha.Tables[0];
        }


        public static DataTable LerArquivoCiclo(IWebDriver drive)
        {
            string caminhoArquivo = EscolherArquivo();

            DataSet planilha = CarregarArquivo(drive, caminhoArquivo);
            return planilha.Tables[0];
        }

        public static DataTable LerArquivoUsuarios(IWebDriver drive)
        {
            try
            {
                int inicio = 654;
                int termino = 700;

                //  string enderecoFisico = AppDomain.CurrentDomain.BaseDirectory;
                //  enderecoFisico = enderecoFisico.Replace("\\bin\\Debug", "");
                // string diretorio = string.Format("{0}{1}", enderecoFisico, "Arquivos\\CargaUsuarioParceirosAudtool _20161103.xls");

                String diretorio = EscolherArquivo();

                if (!File.Exists(diretorio))
                {
                    //CriarArquivoExcel(diretorio);
                    new Exception("Arquivo Cadastro de Usuários não localizado");
                }

                FileStream fis = new FileStream(diretorio, FileMode.Open);

                HSSFWorkbook wb = new HSSFWorkbook(fis);

                string nomeGuia = "Usuario";


                ArrayList lstGuias = new ArrayList();
                int numeroguias = wb.NumberOfSheets;
                for (int i = 0; i < numeroguias; i++)
                {
                    lstGuias.Add(wb.GetSheetName(i));
                }

                HSSFSheet sheet = (HSSFSheet)wb.GetSheet(nomeGuia);

                if (sheet == null)
                    new Exception(string.Format("{0}:{1}{2}", "Nome da Guia", nomeGuia, "não localizada na planilha"));

                DataTable dt = new DataTable();
                dt.Columns.Add("Nome", Type.GetType("System.String"));
                dt.Columns.Add("Email", Type.GetType("System.String"));
                dt.Columns.Add("Perfil", Type.GetType("System.String"));
                dt.Columns.Add("Telefone", Type.GetType("System.String"));
                dt.Columns.Add("Celular", Type.GetType("System.String"));
                dt.Columns.Add("Senha", Type.GetType("System.String"));
                dt.Columns.Add("DataNascimento", Type.GetType("System.DateTime"));
                dt.Columns.Add("Ativo", Type.GetType("System.Boolean"));
                dt.Columns.Add("EmpresaLotacao", Type.GetType("System.String"));
                dt.Columns.Add("UnidadeLotacao", Type.GetType("System.String"));
                dt.Columns.Add("Erro", Type.GetType("System.String"));
                dt.Columns.Add("Linha", Type.GetType("System.String"));

                for (; inicio < termino; inicio++)
                {
                    DataRow dr = dt.NewRow();
                    HSSFRow row = (HSSFRow)sheet.GetRow(inicio);

                    //System.Diagnostics.Debug.WriteLine("Running test case " + row.GetCell(0).ToString() +  row.GetCell(1).ToString() +
                    //    row.GetCell(2).ToString() + row.GetCell(3).ToString() + row.GetCell(4).ToString() + row.GetCell(5).ToString() +
                    //    row.GetCell(6).ToString());

                    dr["Nome"] = row.GetCell(4).StringCellValue;
                    dr["Email"] = row.GetCell(9).StringCellValue;
                    dr["Perfil"] = row.GetCell(8).StringCellValue;
                    dr["Telefone"] = "(11)1111 1111";//row.GetCell(4).StringCellValue;
                    dr["Celular"] = "(11)1111 1111"; //row.GetCell(5).StringCellValue;
                    dr["Senha"] = "abc.123";
                    dr["DataNascimento"] = new DateTime(2016, 10, 1);// DateTime.Today.ToString("dd/MM/yyyy");
                    dr["Ativo"] = true;
                    dr["EmpresaLotacao"] = row.GetCell(0) == null ? string.Empty : row.GetCell(0).ToString();
                    dr["UnidadeLotacao"] = row.GetCell(2) == null ? string.Empty : row.GetCell(2).ToString();
                    dr["Linha"] = inicio;
                    // Run the test for the current test data row

                    dt.Rows.Add(dr);

                }

                fis.Close();

                return dt;
            }
            catch (Exception e)
            {
                new Exception("Erro na Leitura dos parametros");
                return null;
            }
        }

        public static void LogExecucao(string nome, string email, string linha)
        {

            //string enderecoFisico = AppDomain.CurrentDomain.BaseDirectory;
            //enderecoFisico = enderecoFisico.Replace("\\bin\\Debug", "");
            //string diretorio = string.Format("{0}{1}", enderecoFisico, "Arquivos\\log.txt");

          //  String diretorio = EscolherArquivo();

            String diretorio = @"D:\DEMARCO\txt.txt";

            File.AppendAllText(diretorio, string.Format("{0}-{1}-{2}-{3}", linha, DateTime.Now, nome, email) + Environment.NewLine);

            //using (StreamWriter writter = new StreamWriter(diretorio)) { 


            //    writter.WriteLine(string.Format("{0}-{1}-{2}", DateTime.Now, nome, email));
            //}   
        }

        public static DataTable LerArquivoPerguntas(IWebDriver drive)
        {
            try
            {
                Console.WriteLine("Informe o Arquivo das PERGUNTAS\n");

                string caminhoArquivo = EscolherArquivo();

                if (!File.Exists(caminhoArquivo))
                {
                    CriarArquivoExcel(caminhoArquivo);
                }

                // Criar 
                var dt = CarregarArquivo(drive, caminhoArquivo);

               /* int inicio = 14;
                int termino = 31;

                FileStream fis = new FileStream(caminhoArquivo, FileMode.Open);

                HSSFWorkbook wb = new HSSFWorkbook(fis);
                string nomeGuia = SelecionarGuia(wb);
                HSSFSheet sheet = (HSSFSheet)wb.GetSheet(nomeGuia);


                DataTable dt = new DataTable();
                dt.Columns.Add("Quant", Type.GetType("System.Int32"));
                dt.Columns.Add("Categoria", Type.GetType("System.String"));
                dt.Columns.Add("Modalidade", Type.GetType("System.String"));
                dt.Columns.Add("FormaPreenchimento", Type.GetType("System.String"));
                dt.Columns.Add("NumeroPergunta", Type.GetType("System.String"));
                dt.Columns.Add("TituloPTBR", Type.GetType("System.String"));
                dt.Columns.Add("DescricaoPTBR", Type.GetType("System.String"));
                dt.Columns.Add("ItemVerificacaoPTBR", Type.GetType("System.String"));
                dt.Columns.Add("AplicacaoPadraoPTBR", Type.GetType("System.String"));


                for (; inicio < termino; inicio++)
                {
                    DataRow dr = dt.NewRow();
                    HSSFRow row = (HSSFRow)sheet.GetRow(inicio);

                    //System.Diagnostics.Debug.WriteLine("Running test case " + row.GetCell(0).ToString() +  row.GetCell(1).ToString() +
                    //    row.GetCell(2).ToString() + row.GetCell(3).ToString() + row.GetCell(4).ToString() + row.GetCell(5).ToString() +
                    //    row.GetCell(6).ToString());

                    dr["Quant"] = row.GetCell(0).NumericCellValue;
                    dr["Categoria"] = row.GetCell(1).ToString();
                    dr["Modalidade"] = row.GetCell(2).ToString();
                    dr["FormaPreenchimento"] = row.GetCell(3).StringCellValue;
                    //dr["Peso"]
                    dr["NumeroPergunta"] = row.GetCell(5) == null ? string.Empty : row.GetCell(5).ToString();
                    dr["TituloPTBR"] = row.GetCell(6).ToString();
                    dr["DescricaoPTBR"] = row.GetCell(7).ToString();
                    dr["ItemVerificacaoPTBR"] = row.GetCell(8) == null ? string.Empty : row.GetCell(8).ToString();
                    dr["AplicacaoPadraoPTBR"] = row.GetCell(9) == null ? string.Empty : row.GetCell(9).ToString();
                    // Run the test for the current test data row

                    dt.Rows.Add(dr);

                }

                fis.Close();*/

                if (dt == null)
                    throw new Exception("Planilha vazia ou inválida.");

                return dt.Tables[0];
            }
            catch (Exception e)
            {
                new Exception("Erro na Leitura de parametros");
                return null;
            }
        }

        private static string SelecionarGuia(HSSFWorkbook wb)
        {
            ArrayList lstGuias = new ArrayList();
            int numeroguias = wb.NumberOfSheets;

            Console.WriteLine("Informe a guia:");
            for (int i = 0; i < numeroguias; i++)
            {
                Console.WriteLine(String.Format("{0} - {1}", i.ToString(), wb.GetSheetName(i)));

            }

            int itemescolhido = Int32.Parse(Console.ReadLine());
            string nomeGuia = wb.GetSheetName(itemescolhido);
            return nomeGuia;
        }

        public static DataTable LerArquivoAlternativas(IWebDriver drive, int start)
        {
            try
            {

                Console.WriteLine("Informe o Arquivo das ALTERNATIVAS\n");
                String caminhoArquivo = EscolherArquivo();
                
                //Console.WriteLine("Informe a quantidade de alternativas:\n");
                //int qtd_alternativas = Int32.Parse(Console.ReadLine());


                if (!File.Exists(caminhoArquivo))
                {
                    CriarArquivoExcel(caminhoArquivo);
                }

                DataSet plan = CarregarArquivo(drive, caminhoArquivo);

                FileStream fis = new FileStream(caminhoArquivo, FileMode.Open);

                HSSFWorkbook wb = new HSSFWorkbook(fis);

                String nomeGuia = SelecionarGuia(wb);


                HSSFSheet sheet = (HSSFSheet)wb.GetSheet(nomeGuia);

                DataTable dt = new DataTable();
                dt.Columns.Add("Desprezar", Type.GetType("System.String"));
                dt.Columns.Add("ZerarCategoria", Type.GetType("System.String"));
                dt.Columns.Add("NumeroPergunta", Type.GetType("System.String"));
                dt.Columns.Add("NumeroAlternativa", Type.GetType("System.String"));
                dt.Columns.Add("Peso", Type.GetType("System.String"));
                dt.Columns.Add("DescricaoPTBR", Type.GetType("System.String"));
                dt.Columns.Add("DescricaoEN", Type.GetType("System.String"));
                dt.Columns.Add("DescricaoES", Type.GetType("System.String"));
                dt.Columns.Add("Cor", Type.GetType("System.String"));


                int linhaInicial = 1;
                int linhaFinal = plan.Tables[nomeGuia].Rows.Count + 1;


                for (; linhaInicial < linhaFinal; linhaInicial++)
                {
                    DataRow dr = dt.NewRow();
                    HSSFRow row = (HSSFRow)sheet.GetRow(linhaInicial);

                    // dr[0] = row.Cells;

                    dr["Desprezar"] = string.Empty; //row.GetCell(0).ToString();
                    dr["ZerarCategoria"] = row.GetCell(0) == null ? string.Empty : row.GetCell(0).ToString();
                    dr["NumeroPergunta"] = row.GetCell(3) == null ? string.Empty : row.GetCell(3).ToString();
                    dr["NumeroAlternativa"] = row.GetCell(4) == null ? string.Empty : row.GetCell(4).ToString();

                    dr["Peso"] = row.GetCell(4).ToString();
                    dr["DescricaoPTBR"] = row.GetCell(5).ToString();
                    dr["DescricaoEN"] = row.GetCell(5).ToString();//row.GetCell(7) == null ? string.Empty : row.GetCell(7).ToString();
                    dr["DescricaoES"] = row.GetCell(5).ToString();//row.GetCell(8) == null ? string.Empty : row.GetCell(8).ToString();
                    dr["Cor"] = row.GetCell(6) == null ? string.Empty : row.GetCell(6).ToString();

                    dt.Rows.Add(dr);
                }

                fis.Close();

                if (dt == null)
                    throw new Exception("Planilha Alternativas vazia ou inválida.");

                return dt;
            }
            catch (Exception e)
            {
                new Exception("Erro na Leitura das Alternativas");
                return null;
            }
        }

        public static void CriarArquivoExcel(string diretorio)
        {

            HSSFWorkbook wb = HSSFWorkbook.Create(InternalWorkbook.CreateWorkbook());

            HSSFSheet sh = (HSSFSheet)wb.CreateSheet("Plan1");

            for (int i = 0; i < 3; i++)
            {
                var r = sh.CreateRow(i);
                for (int j = 0; j < 2; j++)
                {
                    r.CreateCell(j);
                }
            }

            using (var fs = new FileStream(diretorio, FileMode.Create, FileAccess.Write))
            {
                wb.Write(fs);
            }
        }

        public static DataSet CarregarArquivo(IWebDriver drive, String caminhoArquivo)
        {

            //Abrir arquivo
            FileStream fis = new FileStream(caminhoArquivo, FileMode.Open);

            //Recebe o Arquivo
            DataSet ds = new DataSet();
            ds.DataSetName = "Nome Arquivo";

            //Apos o arquivo aberto identifica as guias
            HSSFWorkbook wb = new HSSFWorkbook(fis);

            int numeroguias = wb.NumberOfSheets;
            int numeroColunas;
            Type tipoCelula = null;
            

            for (int i = 0; i < numeroguias; i++)
            {
                DataTable dt = new DataTable();
                //Atribuindo ao nome do dt com o nome da guia
                dt.TableName = wb.GetSheetName(i).ToString();
                // Adcionando as tabelas no datatable
                ds.Tables.Add(dt);

                //Arquivo
                HSSFSheet sheet = (HSSFSheet)wb.GetSheetAt(i);
                //Utilizar o cabeçalho.
                HSSFRow row = (HSSFRow)sheet.GetRow(0);
                //Ultilizar a primeira linha de dados
               
                // Quantidade de linhas o arquivo possui
               // countRow = sheet.LastRowNum;
                int numRow = sheet.PhysicalNumberOfRows;

                //Quantidade de colunas  no cabeçalho
                numeroColunas = sheet.GetRow(0).PhysicalNumberOfCells;
                //Adcionando linhas 
                for (int r = 1; r < numRow; r++)
                {
                    HSSFRow rowData = (HSSFRow)sheet.GetRow(r);
                    // Adcionando Linhas ao datatable
                    DataRow dr = dt.NewRow();
                    //Informações das celulas para cada Linha 
                    for (int c = 0; c < numeroColunas; c++)
                    {
                        HSSFCell cell = (HSSFCell)rowData.GetCell(c);
                        #region Definir Cabecalho do Data Table
                        //garantir que execute somente para linha 1 uma unica vez por guia.
                        if (r == 1)
                        {
                            DataColumn dc = new DataColumn();

                            if (cell != null)
                                {
                                switch (cell.CellType)
                                {
                                    case NPOI.SS.UserModel.CellType.String:
                                        tipoCelula = cell.StringCellValue.GetType();
                                        break;

                                    case NPOI.SS.UserModel.CellType.Numeric:
                                        tipoCelula = cell.NumericCellValue.GetType();
                                        break;
                                }
                                // Tipo da dado da coluna (rowData - Ultilizar a primeira linha de dados)
                                dc.DataType = Type.GetType(tipoCelula.ToString());
                                // Nome da coluna (row - Utilizar o cabeçalho)
                                dc.ColumnName = row.GetCell(c).ToString();
                            }
                            //Adcionando as colunas 
                            dt.Columns.Add(dc);
                            #endregion
                        }
                        if (cell != null)
                        {
                            //Atribuir valor para linha
                            switch (cell.CellType)
                            {
                                case NPOI.SS.UserModel.CellType.String:
                                    dr[c] = rowData.GetCell(c).StringCellValue;
                                    break;

                                case NPOI.SS.UserModel.CellType.Numeric:
 
                                   dr[c] = rowData.GetCell(c).NumericCellValue;
                                    break;
                            }
                        }
                    }
                    //Adicionar a Linha no Data Table
                    dt.Rows.Add(dr);
                }

                fis.Close();
            }
            return ds;
        }

    }
}

