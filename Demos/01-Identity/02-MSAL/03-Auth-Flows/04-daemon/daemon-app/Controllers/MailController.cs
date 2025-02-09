using System;
using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace MSALDaemon
{
    [Route("api/[controller]")]
    [ApiController]
    public class MailController : ControllerBase
    {
        AppConfig config { get; set; }
        public MailController(IOptions<AppConfig> cfg)
        {
            config = (AppConfig)cfg.Value;
        }

        //https://localhost:5001/api/mail
        [HttpGet]
        public ActionResult SendMail()
        {
            GraphHelper.SendMail("Hello World", "A msg from me", new[] { "alexander.pajer@integrations.at" }, config.GraphCfg);
            return Ok();
        }
    }
}