using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using ClosedXML;
using ClosedXML.Excel;
using sun.reflect.generics.tree;
using Aspose.Cells;
using Newtonsoft.Json.Linq;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;

namespace roboContaMargem
{
    public class AutomationWeb
    {
        public IWebDriver driver;// Criando uma instacia do selenium


        public AutomationWeb() // Construtor
        {
            driver = new EdgeDriver();
        }



        public void LinkAcesso()
        {
            driver.Navigate().GoToUrl("https://consignataria.apconsig.ap.gov.br/login"); // Url de acesso;

            driver.FindElement(By.XPath("/html/body/div/div[2]/div/form/div[1]/div/input")).SendKeys("04633019600"); // Usuario;
            driver.FindElement(By.XPath("/html/body/div/div[2]/div/form/div[2]/div/input")).SendKeys("Famcred25*"); // Senha;
            driver.FindElement(By.XPath("/html/body/div/div[2]/div/form/div[3]/button")).Click(); // Click botao



        }

        public void Menu()
        {

            driver.FindElement(By.XPath("/html/body/div/aside/div/nav/ul/li[3]/a")).Click(); //Click para abrir campo de consulta de cpf
            ColetarDados();

        }


        public void ColetarDados()
        {
            using (var workbook = new XLWorkbook(@"C:\Users\I\Desktop\Consulta\Consultaap.xlsx"))
            {
                var worksheet = workbook.Worksheet(1);
                int linha = 2; //obs utilizada 2 vezes, melhorar esse ponto;

                while (!string.IsNullOrEmpty(worksheet.Cell(linha, 1).GetString()))
                {
                    string cpf = worksheet.Cell(linha, 1).GetString().PadLeft(11, '0'); // PadLeft(11, '0') Para garantir que o cpf possua 11 digitos


                    var campoCpf = driver.FindElement(By.XPath("//*[@id=\"app\"]/div[1]/div/form/div/input"));

                    campoCpf.Clear();
                    campoCpf.SendKeys(cpf);
                    campoCpf.SendKeys(Keys.Enter);
                    Thread.Sleep(200);

                    var botao = driver.FindElements(By.XPath("//*[@id=\"app\"]/div[2]/div[1]/table/tbody/tr/td[5]/a"));


                    string nome = null;
                    string matricula = null;
                    string status = null;
                    string entidade = null;
                    string margemFacultatva = null;
                    string margemConsignavelCartao = null;
                    string reservaCartao = null;
                    string margemCartaoBeneficio = null;
                    string reservaCartaoBeneficio = null;
                    string vinculo = null;


                    if (botao.Count>0)
                    {
                        //Coletar dados
                        driver.FindElement(By.XPath("//*[@id=\"app\"]/div[2]/div[1]/table/tbody/tr/td[5]/a")).Click();
                         nome = driver.FindElement(By.XPath("//*[@id=\"tab_servidor\"]/div[1]/div/div/div[2]/div[2]/div[1]/div/input")).GetAttribute("value");
                         matricula = driver.FindElement(By.XPath("//*[@id=\"tab_servidor\"]/div[1]/div/div/div[2]/div[2]/div[2]/div/input")).GetAttribute("value");
                         status = driver.FindElement(By.XPath("//*[@id=\"tab_servidor\"]/div[1]/div/div/div[2]/div[1]/div[2]/div/input")).GetAttribute("value");
                         entidade = driver.FindElement(By.XPath("//*[@id=\"tab_servidor\"]/div[1]/div/div/div[2]/div[1]/div[3]/div/input")).GetAttribute("value");
                         vinculo = driver.FindElement(By.XPath("//*[@id=\"tab_servidor\"]/div[1]/div/div/div[2]/div[2]/div[3]/div/input")).GetAttribute("value");

                        driver.FindElement(By.XPath("//*[@id=\"app\"]/div[1]/div/ul/li[2]/a")).Click();
                        Thread.Sleep(1000);

                        driver.FindElement(By.XPath("//*[@id=\"tab_margem\"]/div/div/div/div[4]/button")).Click();
                        driver.FindElement(By.XPath("//*[@id=\"modal-autenticacao\"]/div/div/div[2]/div/div/input")).SendKeys("Famcred25*");
                        driver.FindElement(By.XPath("//*[@id=\"modal-autenticacao\"]/div/div/div[2]/div/button")).Click();

                        var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                        wait.Until(d =>
                            {
                                margemFacultatva = driver.FindElement(By.XPath("//*[@id=\"tab_margem\"]/div/div/div/div[1]/div[2]/div/div[6]/div/input")).GetAttribute("value");
                                margemConsignavelCartao = driver.FindElement(By.XPath("//*[@id=\"tab_margem\"]/div/div/div/div[2]/div[2]/div/div[3]/input")).GetAttribute("value");
                                margemCartaoBeneficio = driver.FindElement(By.XPath("//*[@id=\"tab_margem\"]/div/div/div/div[3]/div[2]/div/div[3]/input")).GetAttribute("value");

                                var margemElement = d.FindElement(By.XPath("//*[@id=\"tab_margem\"]/div/div/div/div[1]/div[2]/div/div[2]/input")); // Ajuste o XPath se necessário
                                return !margemElement.GetAttribute("value").Trim().Equals("R$ 0,00");
                            });


                        botao = driver.FindElements(By.XPath("//*[@id=\"app\"]/div[1]/div/ul/li[5]/a"));

                        if (botao.Count > 0)
                        {
                            driver.FindElement(By.XPath("//*[@id=\"app\"]/div[1]/div/ul/li[5]/a")).Click();//Consultar reserva de margem Cartao
                            reservaCartao = driver.FindElement(By.XPath("//table[contains(@class,'table-hover')]//tr[td/span[text()='ATIVA']]//td[3]")).Text;//pegando valor ativo OBS : CONFERIR COMO FAZ
                        }

                        driver.FindElement(By.XPath("//*[@id=\"app\"]/div[1]/div/ul/li[6]/a")).Click();//Consultar reserva de margem Cartao
                        botao = driver.FindElements(By.XPath("//table[contains(@class,'table-hover')]//tr[td/span[text()='ATIVA']]//td[3]"));

                        if (botao.Count > 0)
                        {
                            Thread.Sleep(2000);
                            reservaCartaoBeneficio = driver.FindElement(By.XPath("//table[contains(@class,'table-hover')]//tr[td/span[text()='ATIVA']]//td[3]")).Text; //pegando valor ativo OBS : CONFERIR COMO FAZ

                        }


                        driver.FindElement(By.XPath("/html/body/div/nav/ul[1]/li/a")).Click(); // Abrir opcoes do menu
                        driver.FindElement(By.XPath("/html/body/div/aside/div/nav/ul/li[3]/a")).Click();//Retorno para consultas de cpfs
                        
                    }

                        //Salvar dados

                    worksheet.Cell(linha, 2).Value = nome;
                    worksheet.Cell(linha, 3).Value = matricula;
                    worksheet.Cell(linha, 4).Value = status;
                    worksheet.Cell(linha, 5).Value = entidade;
                    worksheet.Cell(linha, 6).Value = margemFacultatva;
                    worksheet.Cell(linha, 7).Value = margemConsignavelCartao;
                    worksheet.Cell(linha, 8).Value = reservaCartao;
                    worksheet.Cell(linha, 9).Value = margemCartaoBeneficio;
                    worksheet.Cell(linha, 10).Value = reservaCartaoBeneficio;
                    worksheet.Cell(linha, 11).Value = vinculo;


                    linha++;
                }

                driver.FindElement(By.XPath("//*[@id=\"navbarDropdown\"]")).Click();
                driver.FindElement(By.XPath("/html/body/div/nav/ul[3]/li/div/a")).Click();//Sair
                workbook.Save();
                driver.Close();
            }




        }


    }
}
