using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.SqlServer.Server;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using sistem_informasi_produksi_backend.Helper;
using System.Data;

namespace sistem_informasi_produksi_backend.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class PendaftaranPesertaController(IConfiguration configuration) : Controller
    {
        //readonly PolmanAstraLibrary.PolmanAstraLibrary lib = new(PolmanAstraLibrary.PolmanAstraLibrary.Decrypt(configuration.GetConnectionString("DefaultConnection"), "PoliteknikAstra_ConfigurationKey"));
        readonly PolmanAstraLibrary.PolmanAstraLibrary lib = new(configuration.GetConnectionString("DefaultConnection"));

        DataTable dt = new();

        //[Authorize]
        [HttpPost]
        public IActionResult CreatePendaftaranPeserta([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("sim_createPeserta", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }
    }
}
