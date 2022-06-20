using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace ADPF.Utilities.Helpers
{
    public static class HttpClientFactory
    {
        public static async Task<HttpResponseMessage> GetAsyncWithBearerTokenAdded(string uri, string token, string headerTokenKeyName = "Bearer")
        {
            if (string.IsNullOrEmpty(uri) || string.IsNullOrEmpty(token))
            {
                return new HttpResponseMessage()
                {
                    StatusCode = System.Net.HttpStatusCode.BadRequest
                };

            }

            using (HttpClient httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Clear();

                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue(headerTokenKeyName, token);

                try
                {
                    var result = await httpClient.GetAsync(uri);


                    return result;
                }
                catch (Exception ex)
                {

                }
            }

            return new HttpResponseMessage();
        }

        public static async Task<HttpResponseMessage> PostAsyncWithToken(string requestURI, string token, StringContent content, string headerTokenKeyName = "Authorization")
        {
            if (string.IsNullOrEmpty(requestURI) || string.IsNullOrEmpty(token))
            {
                return new HttpResponseMessage()
                {
                    StatusCode = System.Net.HttpStatusCode.BadRequest
                };
            }
            try
            {
                using (HttpClient httpClient = new HttpClient())
                {
                    httpClient.DefaultRequestHeaders.Clear();

                    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    httpClient.DefaultRequestHeaders.TryAddWithoutValidation(headerTokenKeyName, token);

                    var result = await httpClient.PostAsync(requestURI, content);

                    return result;
                }
            }
            catch (Exception ex)
            {


                return new HttpResponseMessage();
            }
        }
    }
}
