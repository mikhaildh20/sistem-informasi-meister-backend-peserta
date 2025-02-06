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
    public class RencanaAnggaranKegiatanController(IConfiguration configuration, IWebHostEnvironment hostingEnvironment) : Controller
    {
        readonly PolmanAstraLibrary.PolmanAstraLibrary lib = new(PolmanAstraLibrary.PolmanAstraLibrary.Decrypt(configuration.GetConnectionString("DefaultConnection"), "YourDecryptionKey"));
        readonly LDAPAuthentication ldap = new(configuration);
        readonly SendMail mail = new(configuration, hostingEnvironment);
        DataTable dt = new();
        DataTable dtNotifikasi = new();

        [Authorize]
        [HttpPost]
        public IActionResult GetDataRencanaAnggaranKegiatan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataRencanaAnggaranKegiatan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataRencanaAnggaranKegiatanById([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataRencanaAnggaranKegiatanById", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataProdukByRAK([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataProdukByRAK", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CreateRencanaAnggaranKegiatan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_createRencanaAnggaranKegiatan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult DetailRencanaAnggaranKegiatan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_detailRencanaAnggaranKegiatan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult EditRencanaAnggaranKegiatan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_editRencanaAnggaranKegiatan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [SupportedOSPlatform("windows")]
        [HttpPost]
        public async Task<IActionResult> SentRencanaAnggaranKegiatan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_sentRencanaAnggaranKegiatan", EncodeData.HtmlEncodeObject(value));

                if (dt.Rows[0][0].ToString()!.Equals("SUKSES"))
                {
                    string htmlBody = "";
                    string footer = "";
                    string result;
                    DataTable dataProduk = lib.CallProcedure("pro_getDataProdukByRAK", [EncodeData.HtmlEncodeObject(value)[0], "Detail"]);
                    DataTable dataBiaya = lib.CallProcedure("pro_detailRencanaAnggaranKegiatan", [EncodeData.HtmlEncodeObject(value)[0]]);
                    decimal totalBiaya = 0;
                    decimal totalAnggaran = 0;
                    decimal keuntunganAnggaran = 0;
                    decimal biayaTambahan = Convert.ToDecimal(dataBiaya.Rows[0][4].ToString()) + Convert.ToDecimal(dataBiaya.Rows[0][5].ToString()) + Convert.ToDecimal(dataBiaya.Rows[0][6].ToString()) + Convert.ToDecimal(dataBiaya.Rows[0][7].ToString());

                    htmlBody += "Yth. Bapak/Ibu <<TONAME>>,<br><br>";
                    htmlBody += "Sistem informasi produksi mencatat ada pengiriman rencana anggaran kegiatan (RAK) baru dengan rincian sebagai berikut:<br><br>";
                    htmlBody += "Nomor Registrasi Marketing : <b>" + dt.Rows[0][1].ToString() + "</b><br><br>";
                    htmlBody += "Daftar Produk/Jasa :<br><br>";
                    htmlBody += "<table border=1><tr><th>No</th><th>Nama Produk/Jasa</th><th>Jumlah</th><th>Total Biaya (Rp)</th></tr>";
                    for (int i = 0; i < dataProduk.Rows.Count; i++)
                    {
                        htmlBody += "<tr>";
                        htmlBody += "<td style='text-align:center;'>" + (i + 1).ToString() + "</td>";
                        htmlBody += "<td>" + dataProduk.Rows[i][2].ToString() + "</td>";
                        htmlBody += "<td style='text-align:center;'>" + dataProduk.Rows[i][3].ToString() + "</td>";
                        htmlBody += "<td style='text-align:right;'>" + Utilities.Separator(Convert.ToInt64(Convert.ToDecimal(dataProduk.Rows[i][9].ToString())).ToString()) + "</td>";
                        htmlBody += "</tr>";
                        totalBiaya += Convert.ToDecimal(dataProduk.Rows[i][9].ToString());
                    }
                    htmlBody += "</table><br>";
                    htmlBody += "Biaya Tambahan :<br><br>";
                    htmlBody += "<table border=1><tr><th>Biaya Tidak Langsung</th><th>Biaya Pengiriman</th><th>Biaya Garansi</th><th>Biaya Lainnya</th></tr>";
                    htmlBody += "<tr><td style='text-align:center;'>" + dataBiaya.Rows[0][4].ToString() + "%</td><td style='text-align:center;'>" + dataBiaya.Rows[0][5].ToString() + "%</td><td style='text-align:center;'>" + dataBiaya.Rows[0][6].ToString() + "%</td><td style='text-align:center;'>" + dataBiaya.Rows[0][7].ToString() + "%</td></tr>";
                    htmlBody += "</table><br>";

                    totalAnggaran += totalBiaya + (totalBiaya * biayaTambahan / 100);
                    htmlBody += "Total Anggaran (Rp) : <b>" + Utilities.Separator(Convert.ToInt64(totalAnggaran).ToString()) + "</b><br><br>";
                    htmlBody += "Simulasi Alternatif Keuntungan dan Diskon :<br><br>";
                    htmlBody += "<table border=1><tr><th>Keuntungan (%)</th><th>Diskon (%)</th><th>Total Harga Jual (Rp)</th></tr>";
                    for (int i = 8; i <= 8; i++)
                    {
                        keuntunganAnggaran = totalAnggaran * Convert.ToDecimal(dataBiaya.Rows[0][i].ToString()) / 100;
                        htmlBody += "<tr>";
                        htmlBody += "<td style='text-align:center;'>" + dataBiaya.Rows[0][i].ToString() + "</td>";
                        htmlBody += "<td style='text-align:center;'>" + dataBiaya.Rows[0][i + 3].ToString() + "</td>";
                        htmlBody += "<td style='text-align:right;'>" + Utilities.Separator(Convert.ToInt64((totalAnggaran + keuntunganAnggaran - (keuntunganAnggaran * Convert.ToDecimal(dataBiaya.Rows[0][i + 3].ToString()) / 100))).ToString()) + "</td>";
                        htmlBody += "</tr>";
                    }
                    htmlBody += "</table><br><br>";

                    footer += "Silakan lakukan persetujuan atau penolakan pada RAK tersebut dengan cara masuk terlebih dahulu ke <a href='" + configuration["Key:linkRoot"]! + "'>Sistem Informasi Produksi</a>.<br><br>";
                    footer += "<b>SISFO Produksi<br>Politeknik Astra</b><br><br>";
                    footer += "Catatan:<br>Email ini dibuat otomatis oleh sistem, jangan membalas email ini.";

                    string toWho = lib.CallProcedure("pro_sentRencanaAnggaranKegiatanDecision", [EncodeData.HtmlEncodeObject(value)[0], Convert.ToInt64(totalAnggaran).ToString()]).Rows[0][0].ToString()!;

                    dtNotifikasi = lib.CallProcedure("all_createNotifikasiNew", [
                        "SENTTO" + toWho,
                        EncodeData.HtmlEncodeObject(value)[0],
                        "APP58",
                        ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[1]),
                        EncodeData.HtmlEncodeObject(value)[1],
                        "<a href='rencana_anggaran_kegiatan'>Rencana Anggaran Kegiatan Baru No. " + dt.Rows[0][1].ToString() + " - Menunggu Persetujuan</a>",
                        "[SISFO Produksi] - " + ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[1]).ToUpper() + " Mengirimkan Rencana Anggaran Kegiatan No. " + dt.Rows[0][1].ToString() + " (Menunggu Persetujuan)",
                        htmlBody,
                        footer,
                        "Detail",
                        "Rencana Anggaran Kegiatan (Menunggu Persetujuan)",
                        "",
                        ""
                    ]);

                    //SEMENTARA KIRIM EMAIL DISINI - BEGIN
                    for (int i = 0; i < dtNotifikasi.Rows.Count; i++)
                    {
                        result = await mail.Send(
                            "[SISFO Produksi] - " + ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[1]).ToUpper() + " Mengirimkan Rencana Anggaran Kegiatan No. " + dt.Rows[0][1].ToString() + " (Menunggu Persetujuan)",
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

        [Authorize]
        [SupportedOSPlatform("windows")]
        [HttpPost]
        public async Task<IActionResult> ApproveRencanaAnggaranKegiatan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_approveRencanaAnggaranKegiatan", EncodeData.HtmlEncodeObject(value));

                string htmlBody = "";
                string footer = "";
                string result;

                htmlBody += "Yth. Bapak/Ibu <<TONAME>>,<br><br>";
                htmlBody += "Sistem informasi produksi mencatat bahwa rencana anggaran kegiatan (RAK) yang Bapak/Ibu ajukan telah <b>disetujui</b> oleh " + dt.Rows[0][2].ToString() + " dengan rincian sebagai berikut:<br><br>";
                htmlBody += "Nomor Registrasi Marketing : <b>" + dt.Rows[0][1].ToString() + "</b><br><br>";

                footer += "Silakan lakukan pengecekan pada RAK tersebut dengan cara masuk terlebih dahulu ke <a href='" + configuration["Key:linkRoot"]! + "'>Sistem Informasi Produksi</a>.<br><br>";
                footer += "<b>SISFO Produksi<br>Politeknik Astra</b><br><br>";
                footer += "Catatan:<br>Email ini dibuat otomatis oleh sistem, jangan membalas email ini.";

                dtNotifikasi = lib.CallProcedure("all_createNotifikasiNew", [
                    "SENTTOMARKETING",
                    EncodeData.HtmlEncodeObject(value)[0],
                    "APP58",
                    ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[1]),
                    EncodeData.HtmlEncodeObject(value)[1],
                    "<a href='rencana_anggaran_kegiatan'>Rencana Anggaran Kegiatan No. " + dt.Rows[0][1].ToString() + " - Disetujui " + dt.Rows[0][2].ToString() + "</a>",
                    "[SISFO Produksi] - " + dt.Rows[0][2].ToString() + " Mengirimkan Rencana Anggaran Kegiatan No. " + dt.Rows[0][1].ToString() + " (Disetujui)",
                    htmlBody,
                    footer,
                    "Detail",
                    "Rencana Anggaran Kegiatan (Disetujui)",
                    "",
                    ""
                ]);

                //SEMENTARA KIRIM EMAIL DISINI - BEGIN
                for (int i = 0; i < dtNotifikasi.Rows.Count; i++)
                {
                    result = await mail.Send(
                        "[SISFO Produksi] - " + dt.Rows[0][2].ToString() + " Mengirimkan Rencana Anggaran Kegiatan No. " + dt.Rows[0][1].ToString() + " (Disetujui)",
                        ldap.GetMail(dtNotifikasi.Rows[i][0].ToString()!),
                        htmlBody.Replace("<<TONAME>>", ldap.GetDisplayName(dtNotifikasi.Rows[i][0].ToString()!)) + footer
                    );
                    if (!result.Equals("OK")) return BadRequest(result);
                }
                //SEMENTARA KIRIM EMAIL DISINI - END

                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [SupportedOSPlatform("windows")]
        [HttpPost]
        public async Task<IActionResult> RejectRencanaAnggaranKegiatan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_rejectRencanaAnggaranKegiatan", EncodeData.HtmlEncodeObject(value));

                string htmlBody = "";
                string footer = "";
                string result;

                htmlBody += "Yth. Bapak/Ibu <<TONAME>>,<br><br>";
                htmlBody += "Sistem informasi produksi mencatat bahwa rencana anggaran kegiatan (RAK) yang Bapak/Ibu ajukan telah <b>ditolak</b> oleh " + dt.Rows[0][2].ToString() + " dengan rincian sebagai berikut:<br><br>";
                htmlBody += "Nomor Registrasi Marketing : <b>" + dt.Rows[0][1].ToString() + "</b><br>";
                htmlBody += "Alasan Tolak : <b>" + EncodeData.HtmlEncodeObject(value)[1] + "</b><br><br>";

                footer += "Silakan lakukan pengecekan pada RAK tersebut dengan cara masuk terlebih dahulu ke <a href='" + configuration["Key:linkRoot"]! + "'>Sistem Informasi Produksi</a>.<br><br>";
                footer += "<b>SISFO Produksi<br>Politeknik Astra</b><br><br>";
                footer += "Catatan:<br>Email ini dibuat otomatis oleh sistem, jangan membalas email ini.";

                dtNotifikasi = lib.CallProcedure("all_createNotifikasiNew", [
                    "SENTTOMARKETING",
                    EncodeData.HtmlEncodeObject(value)[0],
                    "APP58",
                    ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[2]),
                    EncodeData.HtmlEncodeObject(value)[2],
                    "<a href='rencana_anggaran_kegiatan'>Rencana Anggaran Kegiatan No. " + dt.Rows[0][1].ToString() + " - Ditolak " + dt.Rows[0][2].ToString() + "</a>",
                    "[SISFO Produksi] - " + dt.Rows[0][2].ToString() + " Mengirimkan Rencana Anggaran Kegiatan No. " + dt.Rows[0][1].ToString() + " (Ditolak)",
                    htmlBody,
                    footer,
                    "Detail",
                    "Rencana Anggaran Kegiatan (Ditolak)",
                    "",
                    ""
                ]);

                //SEMENTARA KIRIM EMAIL DISINI - BEGIN
                for (int i = 0; i < dtNotifikasi.Rows.Count; i++)
                {
                    result = await mail.Send(
                        "[SISFO Produksi] - " + dt.Rows[0][2].ToString() + " Mengirimkan Rencana Anggaran Kegiatan No. " + dt.Rows[0][1].ToString() + " (Ditolak)",
                        ldap.GetMail(dtNotifikasi.Rows[i][0].ToString()!),
                        htmlBody.Replace("<<TONAME>>", ldap.GetDisplayName(dtNotifikasi.Rows[i][0].ToString()!)) + footer
                    );
                    if (!result.Equals("OK")) return BadRequest(result);
                }
                //SEMENTARA KIRIM EMAIL DISINI - END

                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataRAKCostMaterial([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataRAKCostMaterial", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataRAKCostProses([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataRAKCostProses", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataRAKCostOther([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataRAKCostOther", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }
    }
}
