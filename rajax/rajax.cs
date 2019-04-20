using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.IO;
using System.Net;
using Newtonsoft.Json;

namespace rajax
{
    public static class Rajax
    {
        [ExcelFunction(Description = "Perform a HTTP request.")]
        public static string RAJAX(string uri, string data = "", string contentType = "", string method = "GET")
        {
            byte[] dataBytes = Encoding.UTF8.GetBytes(data);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
            request.ContentLength = dataBytes.Length;
            if (contentType != "") request.ContentType = contentType;
            request.Method = method;
            request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36";

            if (data != "")
            {
                using (Stream requestBody = request.GetRequestStream())
                {
                    requestBody.Write(dataBytes, 0, dataBytes.Length);
                }
            }

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        [ExcelFunction(Description = "HTTP GET shorthand.")]
        public static string RGET(string uri)
        {
            return RAJAX(uri, "", "", "GET");
        }

        [ExcelFunction(Description = "HTTP GET JSON shorthand.")]
        public static object RGETJSON(string uri)
        {
            return JsonConvert.DeserializeObject(RAJAX(uri, contentType: "application/json"));
        }

        [ExcelFunction(Description = "HTTP POST shorthand.")]
        public static object RPOST(string uri, string data = "")
        {
            return RAJAX(uri, method: "POST", data: data);
        }

        public static async Task<string> GetAsync(string uri)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;

            using (HttpWebResponse response = (HttpWebResponse)await request.GetResponseAsync())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {
                return await reader.ReadToEndAsync();
            }
        }
    }
}

