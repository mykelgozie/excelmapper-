using ClosedXML.Excel;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Timers;

namespace ErrorService
{
    class ServiceClass
    {
        
        public Timer _timer;

        public ServiceClass()
        {
            
            _timer = new Timer(90000) { AutoReset = true };
            _timer.Elapsed += _timer_Elapsed;



        }

        public void _timer_Elapsed(object sender , ElapsedEventArgs e)
        {
          

            var dataResults = new List<string>();
            var dataObject = new List<Data>();
            var count = 0;
            var totalData = "";
            var pages = 200;
            var min = 1;
            var max = 10;
            var realTotal = "";




            for (int i = 1; i <= pages; i++)
            {



                HtmlWeb web = new HtmlWeb();
                string url = $"https://www.cbn.gov.ng/rates/GovtSecuritiesDrillDown.asp?beginrec={min}&endrec={max}&market=";
                HtmlDocument document = web.Load(url);
                // var nodes = document.DocumentNode.SelectNodes("//tbody");


                foreach (HtmlNode col in document.DocumentNode.SelectNodes("//table[@id='mytables']//tr//td"))
                {

                    if (col.InnerText == "" && count == 0)
                    {

                        totalData += "*";
                        count++;
                    }
                    else if (col.InnerText != "")
                    {

                        totalData += col.InnerText + "+";
                        count = 0;
                    }

                }


                min += 10;
                max += 10;


            }



            var realData = totalData.Remove(totalData.Length - 1, 1);
            var arrValue = realData.Split('*');



            foreach (var item in arrValue)
            {

                var newItem = item.Remove(item.Length - 1, 1);

                var arrObject = newItem.Split('+');

                var data = new Data();
                data.AuctionDate = arrObject[0];
                data.SecurityType = arrObject[1];
                data.Tenor = arrObject[2];
                data.AuctionNumber = arrObject[3];
                data.Auction = arrObject[4];
                data.MaturityDate = arrObject[5];
                data.TotalSubscription = arrObject[6];
                data.TotalSuccesfull = arrObject[7];
                data.RangeBid = arrObject[8];
                data.SucessfullBidRates = arrObject[9];
                data.Description = arrObject[10];
                data.Rate = arrObject[11];
                data.TrueYield = arrObject[12];
                data.AmountOffered = arrObject[13];

                dataObject.Add(data);



            }



            var workbook = new XLWorkbook();
            workbook.AddWorksheet("sheetName");
            var ws = workbook.Worksheet("sheetName");
            int row1 = 1;
            int row = 2;
            ws.Cell("A" + row1.ToString()).Value = "Auction Date";
            ws.Cell("B" + row1.ToString()).Value = "Security Type";
            ws.Cell("C" + row1.ToString()).Value = "Tenor";
            ws.Cell("D" + row1.ToString()).Value = "Auction No";
            ws.Cell("E" + row1.ToString()).Value = "Auction";
            ws.Cell("F" + row1.ToString()).Value = "Maturity Date";
            ws.Cell("G" + row1.ToString()).Value = "Total Subscription";
            ws.Cell("H" + row1.ToString()).Value = "Total Successful";
            ws.Cell("I" + row1.ToString()).Value = "Range Bid";
            ws.Cell("J" + row1.ToString()).Value = "Successful Bid Rates";
            ws.Cell("K" + row1.ToString()).Value = "Description";
            ws.Cell("L" + row1.ToString()).Value = "Rate";
            ws.Cell("F" + row1.ToString()).Value = "Amount Offered (mn)";

            foreach (var c in dataObject)
            {
                //mapp to excel

                ws.Cell("A" + row.ToString()).Value = c.AuctionDate;
                ws.Cell("B" + row.ToString()).Value = c.SecurityType;
                ws.Cell("C" + row.ToString()).Value = c.Tenor;
                ws.Cell("D" + row.ToString()).Value = c.AuctionNumber;
                ws.Cell("E" + row.ToString()).Value = c.Auction;
                ws.Cell("F" + row.ToString()).Value = c.TotalSubscription;
                ws.Cell("G" + row.ToString()).Value = c.TotalSuccesfull;
                ws.Cell("H" + row.ToString()).Value = c.RangeBid;
                ws.Cell("I" + row.ToString()).Value = c.SucessfullBidRates;
                ws.Cell("J" + row.ToString()).Value = c.Description;
                ws.Cell("K" + row.ToString()).Value = c.Rate;
                ws.Cell("L" + row.ToString()).Value = c.TrueYield;
                ws.Cell("F" + row.ToString()).Value = c.AmountOffered;


                row++;

            }

            workbook.SaveAs(@"D:\testfile4.xlsx");


        }

        public void Start()
        {
            _timer.Start();
        }


        public void Stop()
        {
            _timer.Stop();
        }
    }
}
