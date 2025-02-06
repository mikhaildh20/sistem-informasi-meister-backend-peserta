using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using sistem_informasi_produksi_backend.Helper;
using System.Data;

namespace sistem_informasi_produksi_backend.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class SuratPenawaranController(IConfiguration configuration, IWebHostEnvironment hostingEnvironment) : Controller
    {
        readonly PolmanAstraLibrary.PolmanAstraLibrary lib = new(PolmanAstraLibrary.PolmanAstraLibrary.Decrypt(configuration.GetConnectionString("DefaultConnection"), "YourDecryptionKey"));
        readonly SendMail mail = new(configuration, hostingEnvironment);
        DataTable dt = new();

        [Authorize]
        [HttpPost]
        public IActionResult GetDataSuratPenawaran([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataSuratPenawaran", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataSuratPenawaranById([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataSuratPenawaranById", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataSuratPenawaranKonfirmasi([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataSuratPenawaranKonfirmasi", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataSuratPenawaranRiwayat([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataSuratPenawaranRiwayat", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataProdukByPermintaan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataProdukByPermintaan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CreateSuratPenawaran([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_createSuratPenawaran", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult DetailSuratPenawaran([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_detailSuratPenawaran", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult EditSuratPenawaran([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_editSuratPenawaran", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetListAlternatifKeuntunganDiskonRAK([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getListAlternatifKeuntunganDiskonRAK", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult DetailRAKbyPermintaanPelanggan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_detailRAKbyPermintaanPelanggan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CheckPrintSuratPenawaran([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_checkPrintSuratPenawaran", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public async Task<IActionResult> SentSuratPenawaran([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_sentSuratPenawaran", EncodeData.HtmlEncodeObject(value));

                if (dt.Rows[0][0].ToString()!.Equals("SUKSES"))
                {
                    string htmlBody = "";
                    string footer = "";
                    string result;

                    htmlBody += "Yth. Bapak/Ibu " + dt.Rows[0][4].ToString() + ",<br><br>";
                    htmlBody += "Berdasarkan permintaan terkait produk/jasa yang telah Bapak/Ibu sampaikan, kami dengan ini mengirimkan surat penawaran.<br>";
                    htmlBody += "Mohon untuk meninjau dokumen yang telah kami kirimkan. Apabila ada pertanyaan lebih lanjut, silakan menghubungi bagian Marketing Politeknik Astra.<br><br>";
                    htmlBody += "Demikian yang dapat kami sampaikan, terima kasih.<br><br><br>";

                    footer += "<b>SISFO Produksi<br>Politeknik Astra</b><br><br>";
                    footer += "Catatan:<br>Email ini dibuat otomatis oleh sistem, jangan membalas email ini.";

                    result = await mail.Send(
                        "[Politeknik Astra] Penawaran Produk/Jasa - " + dt.Rows[0][2].ToString(),
                        dt.Rows[0][3].ToString()!,
                        htmlBody + footer,
                        dt.Rows[0][1].ToString()!
                    );
                    if (!result.Equals("OK")) return BadRequest(result);

                    return Ok(JsonConvert.SerializeObject(dt));
                }
                return BadRequest();
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult ApproveSuratPenawaran([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_approveSuratPenawaran", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public async Task<IActionResult> PrintSuratPenawaran([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                string fileName = lib.CallProcedure("pro_printSuratPenawaran", [
                    .. EncodeData.HtmlEncodeObject(value),
                    .. new string[] {
                        Guid.NewGuid() +
                        "_" +
                        DateTime.Now.Year +
                        DateTime.Now.Month +
                        DateTime.Now.Day +
                        DateTime.Now.Hour +
                        DateTime.Now.Minute +
                        DateTime.Now.Second
                    }
                ]).Rows[0][0].ToString() ?? "";

                if (string.IsNullOrEmpty(fileName)) throw new Exception();

                string saveDirectory = Path.Combine(hostingEnvironment.WebRootPath, "Uploads");
                if (!Directory.Exists(saveDirectory))
                    Directory.CreateDirectory(saveDirectory);

                using HttpClient client = new();
                HttpResponseMessage response = await client.PostAsJsonAsync(configuration["Key:reportAPI"]! + "GetReportSuratPenawaran", value.ToString());

                if (response.IsSuccessStatusCode)
                {
                    byte[] content = await response.Content.ReadAsByteArrayAsync();
                    await System.IO.File.WriteAllBytesAsync(Path.Combine(saveDirectory, fileName), content);

                    return Ok(JsonConvert.SerializeObject(new { Hasil = fileName }));
                }
                else throw new Exception();
            }
            catch { return BadRequest(); }
        }
    }
}
