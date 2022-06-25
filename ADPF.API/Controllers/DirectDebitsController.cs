using ADPF.API.Exceptions;
using ADPF.Business.Business.BDirectDebit;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Mvc;

namespace ADPF.API.Controllers
{
    public class DirectDebitsController : ApiController
    {
        // GET: DirectDebits
        [EnableCors(origins: "*", headers: "*", methods: "*")]
        public IHttpActionResult Get( int id)
        {
            try
            {
               var response= BDirectDebit.GetEntries();
                return Ok(response);
            }
            catch (CustomException ex)
            {
                return BadRequest(ex.Message);
            }
           
        }
        
        public IHttpActionResult Post(int id, int x)
        {
            try
            {
                return Ok();
            }
            catch (CustomException ex)
            {

                throw;
            }
            
        }
     
    }
}