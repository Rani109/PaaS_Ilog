using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Web;
using System.Xml;
using Dapper;

namespace AutoInvoicesUS
{
    public static class FuelSurcharge
    {
        public static string GetDieselFuel(int? siteCode = null, int? numerator = null)
        {
            return (SqlHelper.ExecuteScalarSP<string>("sp_get_diesel_fuel", new { SITE_CODE = siteCode, NUMERATOR = numerator }) ?? DefaultDieselFuel);
        }

        public static double GetFuelSurcharge(double fuelPricePerGallon)
        {
            return SqlHelper.ExecuteScalarSP<double>("sp_get_fuel_surcharge", new { FUEL_PRICE_PER_GALLON = fuelPricePerGallon });
        }

        public const string UrlFuelPrices = "https://www.eia.gov/petroleum/gasdiesel/includes/gas_diesel_rss.xml";

        public static List<Tuple<string, double>> GetRSSDieselFuel()
        {
            XmlDocument xmlDoc = new XmlDocument();
            HttpWebRequest request = WebRequest.Create(UrlFuelPrices) as HttpWebRequest;
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            using (Stream responseStream = response.GetResponseStream())
            {
                using (StreamReader streamReader = new StreamReader(responseStream))
                {
                    xmlDoc.LoadXml(streamReader.ReadToEnd());
                }
            }

            var node = xmlDoc.SelectSingleNode("/rss/channel/item/description");
            if (node != null)
            {
                var html = HttpUtility.HtmlDecode(node.InnerText);
                int index = html.IndexOf("On-Highway Diesel Fuel Retail Price");
                if (index != -1)
                {
                    string[] rssLines = html.Substring(index).Split(new string[] { "<br>" }, StringSplitOptions.None);
                    if (rssLines.Length == 1)
                        rssLines = html.Substring(index).Split(new string[] { "<br/>" }, StringSplitOptions.None);

                    // On-Highway Diesel Fuel Retail Price
                    // (Dollars per Gallon)
                    // 2.996 .. U.S.
                    // 3.025 ... East Coast
                    // 3.077 .... New England
                    // 3.209 .... Central Atlantic
                    // 2.887 .... Lower Atlantic
                    // 2.947 ... Midwest
                    // 2.785 ... Gulf Coast
                    // 2.974 ... Rocky Mountain
                    // 3.394 ... West Coast
                    // 3.087 ... West Coast less California
                    // 3.638 .... California

                    List<Tuple<string, double>> dieselLines = new List<Tuple<string, double>>();

                    foreach (string rssLine in rssLines)
                    {
                        string[] segments = rssLine.Replace("....", "|").Replace("...", "|").Replace("..", "|").Split('|');
                        if (segments.Length == 2)
                        {
                            string dieselStr = segments[0].Trim();
                            string region = segments[1].Trim();

                            double diesel = 0;
                            if (double.TryParse(dieselStr, out diesel))
                                dieselLines.Add(new Tuple<string, double>(region, diesel));
                        }
                    }

                    return dieselLines;
                }
            }

            return null;
        }

        public const string DefaultDieselFuel = "U.S.";

        public static double GetDiesel(List<Tuple<string, double>> dieselLines, string dieselFuel = DefaultDieselFuel)
        {
            if (dieselLines == null || dieselLines.Count == 0)
                return 0;

            if (string.IsNullOrEmpty(dieselFuel))
                dieselFuel = DefaultDieselFuel;

            foreach (var dieselLine in dieselLines)
            {
                if (string.Compare(dieselLine.Item1, dieselFuel, true) == 0)
                    return dieselLine.Item2;
            }

            return 0;
        }
    }
}
