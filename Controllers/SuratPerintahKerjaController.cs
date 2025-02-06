using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using sistem_informasi_produksi_backend.Helper;
using System.Data;
using System.Runtime.Versioning;

namespace sistem_informasi_produksi_backend.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class SuratPerintahKerjaController(IConfiguration configuration, IWebHostEnvironment hostingEnvironment) : Controller
    {
        readonly PolmanAstraLibrary.PolmanAstraLibrary lib = new(PolmanAstraLibrary.PolmanAstraLibrary.Decrypt(configuration.GetConnectionString("DefaultConnection"), "YourDecryptionKey"));
        readonly LDAPAuthentication ldap = new(configuration);
        readonly SendMail mail = new(configuration, hostingEnvironment);
        DataTable dt = new();
        DataTable dtNotifikasi = new();

        [Authorize]
        [HttpPost]
        public IActionResult GetDataSuratPerintahKerja([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataSuratPerintahKerja", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataSuratPerintahKerjaById([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataSuratPerintahKerjaById", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataProdukByPermintaan2([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataProdukByPermintaan2", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CreateSuratPerintahKerja([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_createSuratPerintahKerja", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult DetailSuratPerintahKerja([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_detailSuratPerintahKerja", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult DetailPenawaranbyPermintaanPelanggan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_detailPenawaranbyPermintaanPelanggan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult EditSuratPerintahKerja([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_editSuratPerintahKerja", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [SupportedOSPlatform("windows")]
        [HttpPost]
        public async Task<IActionResult> SentSuratPerintahKerja([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_sentSuratPerintahKerja", EncodeData.HtmlEncodeObject(value));

                if (dt.Rows[0][0].ToString()!.Equals("SUKSES"))
                {
                    string htmlBody = "";
                    string footer = "";
                    string result;

                    htmlBody += "Yth. Bapak/Ibu <<TONAME>>,<br><br>";
                    htmlBody += "Sistem informasi produksi mencatat bahwa tim Marketing telah mengirimkan permintaan produksi baru dengan rincian sebagai berikut:<br><br>";
                    htmlBody += "Nomor Registrasi Marketing : <b>" + dt.Rows[0][1].ToString() + "</b><br><br>";

                    footer += "Silakan lakukan pengecekan dan konversi nomor SPK pada permintaan tersebut dengan cara masuk terlebih dahulu ke <a href='" + configuration["Key:linkRoot"]! + "'>Sistem Informasi Produksi</a>.<br><br>";
                    footer += "<b>SISFO Produksi<br>Politeknik Astra</b><br><br>";
                    footer += "Catatan:<br>Email ini dibuat otomatis oleh sistem, jangan membalas email ini.";

                    dtNotifikasi = lib.CallProcedure("all_createNotifikasiNew", [
                        "SENTSURATPERINTAHKERJA",
                        EncodeData.HtmlEncodeObject(value)[0],
                        "APP58",
                        ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[1]),
                        EncodeData.HtmlEncodeObject(value)[1],
                        "<a href='surat_perintah_kerja_operasional'>Permintaan Produksi Baru No. " + dt.Rows[0][1].ToString() + " - Terbit</a>",
                        "[SISFO Produksi] - " + ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[1]).ToUpper() + " Menerbitkan Permintaan Produksi Baru No. " + dt.Rows[0][1].ToString(),
                        htmlBody,
                        footer,
                        "Detail",
                        "Surat Perintah Kerja",
                        "",
                        ""
                    ]);

                    //SEMENTARA KIRIM EMAIL DISINI - BEGIN
                    for (int i = 0; i < dtNotifikasi.Rows.Count; i++)
                    {
                        result = await mail.Send(
                            "[SISFO Produksi] - " + ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[1]).ToUpper() + " Menerbitkan Permintaan Produksi Baru No. " + dt.Rows[0][1].ToString(),
                            ldap.GetMail(dtNotifikasi.Rows[i][0].ToString()!),
                            htmlBody.Replace("<<TONAME>>", ldap.GetDisplayName(dtNotifikasi.Rows[i][0].ToString()!)) + footer
                        );
                        if (!result.Equals("OK")) return BadRequest(result);
                    }
                    //SEMENTARA KIRIM EMAIL DISINI - END

                    return Ok(JsonConvert.SerializeObject(dt));
                }
                return BadRequest();
            }
            catch { return BadRequest(); }
        }
    }
}
