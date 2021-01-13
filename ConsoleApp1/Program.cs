using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading;


namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://veiculos.fipe.org.br/");

            ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementById('selectMarcacarro').removeAttribute('style');");
            ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementById('selectAnoModelocarro').removeAttribute('style');");
            ((IJavaScriptExecutor)driver).ExecuteScript("document.getElementById('selectAnocarro').removeAttribute('style');");
            driver.FindElement(By.CssSelector("li:nth-child(1).ilustra > a[href^='javascript:void(0)']")).Click();

            IWebElement elementMarca = driver.FindElement(By.Id("selectMarcacarro"));
            SelectElement selectMarca = new SelectElement(elementMarca);
            List<Fipe> listaFipe = new List<Fipe>();

            var Marcas = selectMarca.Options.Where(p => p.Text != "").Select(p => p.Text).ToList();
            foreach (string dadoMarca in Marcas)
            {
                selectMarca.SelectByText(dadoMarca);
                IWebElement elementModelo = driver.FindElement(By.Id("selectAnoModelocarro"));
                SelectElement selectModelo = new SelectElement(elementModelo);

                var Modelos = selectModelo.Options.Where(p => p.Text != "").Select(p => p.Text).ToList();
                foreach (string dadoModelo in Modelos)
                {
                    elementModelo = driver.FindElement(By.Id("selectAnoModelocarro"));
                    selectModelo = new SelectElement(elementModelo);
                    selectModelo.SelectByText(dadoModelo);
                    IWebElement elementAno = driver.FindElement(By.Id("selectAnocarro"));
                    SelectElement selectAno = new SelectElement(elementAno);

                    var Anos = selectAno.Options.Where(p => p.Text != "").Select(p => p.Text).ToList();
                    foreach (string dadoAno in Anos)
                    {
                        selectAno.SelectByText(dadoAno);
                        Thread.Sleep(1000);
                        driver.FindElement(By.Id("buttonPesquisarcarro")).Click();

                        IList<IWebElement> allElement = driver.FindElements(By.TagName("td"));
                        IList<IWebElement> allElementTitulo = driver.FindElements(By.ClassName("noborder"));
                        IList<IWebElement> listaDados = allElement.Except(allElementTitulo).Where(p => p.Text != "").ToList();

                        Fipe f = new Fipe();
                        f.Codigo = listaDados[1].Text;
                        f.Marca = listaDados[2].Text;
                        f.Modelo = listaDados[3].Text;
                        f.Ano = listaDados[4].Text;
                        f.Preco = Convert.ToDecimal(listaDados[7].Text.Replace("R$",""));
                        listaFipe.Add(f);

                        InserirFipe(f);

                        Thread.Sleep(1000);
                        driver.FindElement(By.Id("buttonLimparPesquisarcarro")).Click();

                        elementMarca = driver.FindElement(By.Id("selectMarcacarro"));
                        selectMarca = new SelectElement(elementMarca);
                        selectMarca.SelectByText(dadoMarca);

                        elementModelo = driver.FindElement(By.Id("selectAnoModelocarro"));
                        selectModelo = new SelectElement(elementModelo);
                        selectModelo.SelectByText(dadoModelo);

                    }
                }
            }
        }
        
        public static void InserirFipe(Fipe fipe)
        {
          
            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString))
            {
                connection.Open();
                string sql = "INSERT INTO Fipe(Codigo,Marca,Modelo,Ano,Preco) VALUES(@param1,@param2,@param3,@param4,@param5)";
                using (SqlCommand cmd = new SqlCommand(sql, connection))
                {
                    cmd.Parameters.Add("@param1", SqlDbType.NVarChar,100).Value = fipe.Codigo;
                    cmd.Parameters.Add("@param2", SqlDbType.NVarChar, 100).Value = fipe.Marca;
                    cmd.Parameters.Add("@param3", SqlDbType.NVarChar, 100).Value = fipe.Modelo;
                    cmd.Parameters.Add("@param4", SqlDbType.NVarChar, 100).Value = fipe.Ano;
                    cmd.Parameters.Add("@param5", SqlDbType.Decimal).Value =  fipe.Preco;
                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();
                }
            }

        }
    }

    public class Fipe
    {
        public string Codigo { get; set; }
        public string Marca { get; set; }
        public string Modelo { get; set; }
        public string Ano { get; set; }
        public decimal Preco { get; set; }
    }


}
