using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

using System.Text.RegularExpressions;
using Manatee.Trello;
using Manatee.Trello.ManateeJson;
using Manatee.Trello.WebApi;
using Excel = Microsoft.Office.Interop.Excel;


namespace TrelloTest
{
    class Program
    {
        static void Main(string[] args)
        {
            // Manatee configuration
            var serializer = new ManateeSerializer();
            var JsonFactory = new ManateeFactory();
            var RestClientProvider = new WebApiClientProvider();
            TrelloConfiguration.JsonFactory = JsonFactory;
            TrelloConfiguration.RestClientProvider = RestClientProvider;
            TrelloConfiguration.Serializer = serializer;
            TrelloConfiguration.Deserializer = serializer;
            //TrelloConfiguration.JsonFactory = new ManateeFactory();
            //TrelloConfiguration.RestClientProvider = new WebApiClientProvider();
            TrelloAuthorization.Default.AppKey = "ba914894214469f55add1219389d760d";
            TrelloAuthorization.Default.UserToken = "aec5644c8fdd4a89973eaf8cae448826e37abf3406758d9546c9796055a94a07";


        }
    }
}
