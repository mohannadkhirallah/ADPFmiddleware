using ADPF.Business.Models.DTO;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace ADPF.Business.Business.BDirectDebit
{
    public class BDirectDebit
    {
        public BDirectDebit(IOrganizationService service)
        {

        }
        public BDirectDebit()
        {

        }

        #region Constancts 
        #endregion
            #region Method

        public  static EntryDto GetEntries()
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri("https://api.publicapis.org/");
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    //GET Method
                    HttpResponseMessage response = client.GetAsync("entries").Result;
                    if (response.IsSuccessStatusCode)
                    {
                        string apiResponse = response.Content.ReadAsStringAsync().Result;
                        var dto= JsonConvert.DeserializeObject<EntryDto>(apiResponse);
                        return dto;



                    }
                    else
                    {
                        throw new InvalidOperationException("The API is not working");
                    }

                }
               
                }
            catch (Exception)
            {

                throw;
            }

           
        }
        #endregion
    }
}
