using HtmlAgilityPack;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ExchangeRateParser
{
    class Program
    {
        /// <param name="args"></param>

        private static WebBrowser wb;
        private static bool xmlOrJSON = true;
        private static int currency;
        [STAThread]
        static void Main(string[] args)
        {
            //This is horrible horrible code, I'm so sorry! Need to fix this
            // I really can't think of a better way to pass console args right now...
            if (args.Length == 0)
            {
                Console.WriteLine("Please enter args in the following format: arg1 = XML/JSON, arg2 = USD/EUR/RUR");
                return;
            }
            else
            {
                if (args[0].Equals("XML"))
                    xmlOrJSON = false;
                switch (args[1])
                {
                    case "USD":
                        currency = 0;
                        break;
                    case "EUR":
                        currency = 1;
                        break;
                    case "RUR":
                        currency = 2;
                        break;
                }
            }
            
            Uri url = new Uri("https://kurs.kz/");
            
            var th = new Thread(() => {
                wb = new WebBrowser();
                wb.ScriptErrorsSuppressed = true;
                wb.DocumentCompleted += browser_DocumentCompleted;
                wb.Navigate(url);
                Application.Run();
            });
            th.SetApartmentState(ApartmentState.STA);
            th.Start();
          
        }

        //All the work is done in this method, since we have to wait till all of dynamic content is loaded 
        static void browser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            var br = sender as WebBrowser;
            if (br.Url == e.Url)
            {
                Console.WriteLine("Navigated to {0}", e.Url);
                Application.ExitThread();   // Stops the thread
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                var documentAsIHtmlDocument3 = (mshtml.IHTMLDocument3)wb.Document.DomDocument;
                StringReader sr = new StringReader(documentAsIHtmlDocument3.documentElement.outerHTML);
                doc.Load(sr);
                
                //parsing company names and times
                var titleList = doc.DocumentNode.SelectNodes("//div[@id='table']//tr/td[1]/a");
                var timeList = doc.DocumentNode.SelectNodes("//div[@id='table']//tr/td[2]/span/time");
                
                //this is gonna be awful just terrible, but oh well, choose your currency
                //major parsing done in here, we're talking buying rates, selling rates and averages for the day
                HtmlNodeCollection buyRateList;
                HtmlNodeCollection sellRateList;
                HtmlNode avgBuyRate;
                HtmlNode avgSellRate;
                string selectedCurrency;
                switch (currency)
                {
                    case 0:
                        buyRateList = doc.DocumentNode.SelectNodes("//div[@id='table']//tr/td[3]/span[1]");
                        sellRateList = doc.DocumentNode.SelectNodes("//div[@id='table']//tr/td[3]/span[3]");
                        avgBuyRate = doc.DocumentNode.SelectSingleNode("//div[@id='table']//tfoot//th[2]/span[1]");
                        avgSellRate = doc.DocumentNode.SelectSingleNode("//div[@id='table']//tfoot//th[2]/span[2]");
                        selectedCurrency = "USD";
                        break;
                    case 1:
                        buyRateList = doc.DocumentNode.SelectNodes("//div[@id='table']//tr/td[4]/span[1]");
                        sellRateList = doc.DocumentNode.SelectNodes("//div[@id='table']//tr/td[4]/span[3]");
                        avgBuyRate = doc.DocumentNode.SelectSingleNode("//div[@id='table']//tfoot//th[3]/span[1]");
                        avgSellRate = doc.DocumentNode.SelectSingleNode("//div[@id='table']//tfoot//th[3]/span[2]");
                        selectedCurrency = "EUR";
                        break;
                    case 2:
                        buyRateList = doc.DocumentNode.SelectNodes("//div[@id='table']//tr/td[5]/span[1]");
                        sellRateList = doc.DocumentNode.SelectNodes("//div[@id='table']//tr/td[5]/span[3]");
                        avgBuyRate = doc.DocumentNode.SelectSingleNode("//div[@id='table']//tfoot//th[4]/span[1]");
                        avgSellRate = doc.DocumentNode.SelectSingleNode("//div[@id='table']//tfoot//th[4]/span[2]");
                        selectedCurrency = "RUR";
                        break;
                    default:
                        buyRateList = doc.DocumentNode.SelectNodes("//div[@id='table']//tr/td[3]/span[1]");
                        sellRateList = doc.DocumentNode.SelectNodes("//div[@id='table']//tr/td[3]/span[3]");
                        avgBuyRate = doc.DocumentNode.SelectSingleNode("//div[@id='table']//tfoot//th[2]/span[1]");
                        avgSellRate = doc.DocumentNode.SelectSingleNode("//div[@id='table']//tfoot//th[2]/span[2]");
                        selectedCurrency = "USD";
                        break;
                }

                //Creating objects for conversion into the appropriate format
                List<rates> ex_rates = new List<rates>();
                for(int i = 0; i < titleList.Count; i++)
                {
                    ex_rates.Add(new rates
                    {
                        CompanyName = titleList[i].InnerText,
                        time = timeList[i].InnerText,
                        chosenCurrency = selectedCurrency,
                        buyRate = buyRateList[i].InnerText,
                        sellRate = sellRateList[i].InnerText
                    });
                }
                
                //Choosing the format to write to file: either XML or JSON
                if (xmlOrJSON)
                {
                    using (StreamWriter file = File.CreateText(@"rates.txt"))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        //serialize object directly into file stream
                        serializer.Serialize(file, ex_rates);
                    }
                }
                else
                {
                    using (XmlWriter writer = XmlWriter.Create("rates.xml"))
                    {
                        writer.WriteStartDocument();
                        writer.WriteStartElement("ExchangeRates");

                        foreach (rates rate in ex_rates)
                        {
                            writer.WriteStartElement("ExchangeRate");

                            writer.WriteElementString("CompanyName", rate.CompanyName);
                            writer.WriteElementString("time", rate.time);
                            writer.WriteElementString("chosenCurrency", rate.chosenCurrency);
                            writer.WriteElementString("buyRate", rate.buyRate);
                            writer.WriteElementString("sellRate", rate.sellRate);

                            writer.WriteEndElement();
                        }

                        writer.WriteEndElement();
                        writer.WriteEndDocument();
                    }
                }

                Console.WriteLine("Average Buy Rate for today: " + avgBuyRate.InnerText + " KZT for 1 " + selectedCurrency);
                Console.WriteLine("Average Sell Rate for today: " + avgSellRate.InnerText + " KZT for 1 " + selectedCurrency);
                Console.Read();
            }
            
        }

        // think of this as a data class...
        public class rates
        {
            public string CompanyName { get; set; }
            public string time { get; set; }
            public string chosenCurrency { get; set; }
            public string buyRate { get; set; }
            public string sellRate { get; set; }
        }
    }
}
