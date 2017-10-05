using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace OutlookRESTAPI_UpdateCalendar
{
    class HttpHelper
    {
        /// <summary>
        /// Perform an HTTP GET request to a URL using an HTTP Authorization header
        /// </summary>
        /// <param name="method"></param>
        /// <param name="url">The URL</param>
        /// <param name="token">The token</param>
        /// <param name="requestContent"></param>
        /// <returns>String containing the results of the GET operation</returns>
        public static string SendAsync(HttpMethod method, string url, string token, object requestContent = null)
        {
            var httpClient = new HttpClient();
            try
            {
                var request = new HttpRequestMessage(method, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                if (requestContent != null)
                {
                    string contentString = JsonConvert.SerializeObject(requestContent, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                    request.Content = new StringContent(contentString, System.Text.Encoding.UTF8, "application/json");
                }

                var response = httpClient.SendAsync(request).Result;

                if (response.IsSuccessStatusCode)
                {
                    var content = response.Content.ReadAsStringAsync().Result;
                    return content;
                }

                Console.WriteLine("HTTP Request Failed.  Status Code : " + response.StatusCode);
                return null;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }
    }
}
