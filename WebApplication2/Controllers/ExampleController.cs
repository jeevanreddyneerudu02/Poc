using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace WebApplication2.Controllers
{
    public class ExampleController : ApiController
    {
        public IHttpActionResult ProcessFile([FromBody]string fileName)
        {
            // do something with fileName parameter

            return Ok();
        }
    }
}
