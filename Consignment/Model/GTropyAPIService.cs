using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Consignment.Model
{
    public class GTropyAPIService
    {
        // Ek hi shared HttpClient, 20 sec timeout ke saath - GTropy API slow ho to LR save hang na ho
        private static readonly HttpClient client = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };

        public async Task SendToGtropyAPI(string vehicle, string invoice, DateTime startDateTime, string origin, string destination)
        //public async Task SendToGtropyAPI()
        {
            var url = "https://ctyf.co.in/api/uploadjsontrip";

            var payload = new
            {
                Token = "41DDE3979A",
                Vehicle = vehicle,
                Invoice = invoice,
                Startdttime = startDateTime.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                Origin = origin,
                Destination = destination
            };

            var jsonPayload = JsonConvert.SerializeObject(payload);
            var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

            try
            {
                HttpResponseMessage response = await client.PostAsync(url, content);
                string responseBody = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    string[] parts = responseBody.Split('|');
                    if (parts[0] == "1")
                    {
                        Console.WriteLine("GTropy Upload Success.");
                    }
                    else
                    {
                        Console.WriteLine("GTropy API Error");
                    }
                }
                else
                {
                    Console.WriteLine("HTTP Error");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception in Gtropy API Call : " + ex.Message);
            }
        }
    }
}