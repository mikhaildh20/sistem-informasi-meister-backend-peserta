using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using sistem_informasi_produksi_backend.Helper;
using System.Data;
using System.Runtime.Versioning;

namespace sistem_informasi_produksi_backend.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class PermintaanPelangganController(IConfiguration configuration, IWebHostEnvironment hostingEnvironment) : Controller
    {
        readonly PolmanAstraLibrary.PolmanAstraLibrary lib = new(PolmanAstraLibrary.PolmanAstraLibrary.Decrypt(configuration.GetConnectionString("DefaultConnection"), "YourDecryptionKey"));
        readonly LDAPAuthentication ldap = new(configuration);
        readonly SendMail mail = new(configuration, hostingEnvironment);
        DataTable dt = new();
        DataTable dtNotifikasi = new();

        [Authorize]
        [HttpPost]
        public IActionResult GetDataPermintaanPelanggan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataPermintaanPelanggan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataPermintaanPelangganById([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataPermintaanPelangganById", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetListPermintaanPelanggan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getListPermintaanPelanggan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetListPermintaanPelanggan2([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getListPermintaanPelanggan2", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetListPermintaanPelanggan3([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getListPermintaanPelanggan3", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CreatePermintaanPelanggan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_createPermintaanPelanggan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CreatePermintaanPelangganProduk([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_createPermintaanPelangganProduk", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult DetailPermintaanPelanggan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_detailPermintaanPelanggan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult DetailPermintaanPelangganProduk([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_detailPermintaanPelangganProduk", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult EditPermintaanPelanggan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_editPermintaanPelanggan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult DeletePermintaanPelanggan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_deletePermintaanPelanggan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CancelPermintaanPelanggan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_cancelPermintaanPelanggan", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [SupportedOSPlatform("windows")]
        [HttpPost]
        public async Task<IActionResult> SentPermintaanPelanggan([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_sentPermintaanPelanggan", EncodeData.HtmlEncodeObject(value));

                if (dt.Rows[0][0].ToString()!.Equals("SUKSES"))
                {
                    string htmlBody = "";
                    string footer = "";
                    string result;

                    htmlBody += "Yth. Bapak/Ibu <<TONAME>>,<br><br>";
                    htmlBody += "Sistem informasi produksi mencatat ada pengiriman permintaan pelanggan baru dengan rincian sebagai berikut:<br><br>";
                    htmlBody += "Nomor Registrasi Marketing : <b>" + dt.Rows[0][1].ToString() + "</b><br>";
                    htmlBody += "Status : <b>Menunggu Analisa PPIC dan Engineering</b><br><br>";

                    footer += "Silakan lakukan pengecekan pada permintaan pelanggan tersebut dengan cara masuk terlebih dahulu ke <a href='" + configuration["Key:linkRoot"]! + "'>Sistem Informasi Produksi</a>.<br><br>";
                    footer += "<b>SISFO Produksi<br>Politeknik Astra</b><br><br>";
                    footer += "Catatan:<br>Email ini dibuat otomatis oleh sistem, jangan membalas email ini.";

                    dtNotifikasi = lib.CallProcedure("all_createNotifikasiNew", [
                        "SENTPERMINTAANPELANGGAN",
                        EncodeData.HtmlEncodeObject(value)[0],
                        "APP58",
                        ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[1]),
                        EncodeData.HtmlEncodeObject(value)[1],
                        "<a href='permintaan_pelanggan'>Permintaan Pelanggan Baru No. " + dt.Rows[0][1].ToString() + " - Menunggu Analisa</a>",
                        "[SISFO Produksi] - " + ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[1]).ToUpper() + " Mengirimkan Permintaan Pelanggan No. " + dt.Rows[0][1].ToString() + " (Menunggu Analisa)",
                        htmlBody,
                        footer,
                        "Detail",
                        "Permintaan Pelanggan (Menunggu Analisa)",
                        "",
                        ""
                    ]);

                    //SEMENTARA KIRIM EMAIL DISINI - BEGIN
                    for (int i = 0; i < dtNotifikasi.Rows.Count; i++)
                    {
                        result = await mail.Send(
                            "[SISFO Produksi] - " + ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[1]).ToUpper() + " Mengirimkan Permintaan Pelanggan No. " + dt.Rows[0][1].ToString() + " (Menunggu Analisa)",
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
        public async Task<IActionResult> ApprovePermintaanPelangganProduk([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_approvePermintaanPelangganProduk", EncodeData.HtmlEncodeObject(value));

                if (dt.Rows[0][0].ToString()!.Equals("END"))
                {
                    string htmlBody = "";
                    string footer = "";
                    string result;

                    htmlBody += "Yth. Bapak/Ibu <<TONAME>>,<br><br>";
                    htmlBody += "Sistem informasi produksi mencatat ada permintaan pelanggan yang telah selesai dilakukan analisa oleh Engineering dan PPIC dengan rincian sebagai berikut:<br><br>";
                    htmlBody += "Nomor Registrasi Marketing : <b>" + dt.Rows[0][1].ToString() + "</b><br>";
                    htmlBody += "Status : <b>" + dt.Rows[0][2].ToString() + "</b><br><br>";

                    footer += "Silakan lakukan pengecekan pada permintaan pelanggan tersebut dengan cara masuk terlebih dahulu ke <a href='" + configuration["Key:linkRoot"]! + "'>Sistem Informasi Produksi</a>.<br><br>";
                    footer += "<b>SISFO Produksi<br>Politeknik Astra</b><br><br>";
                    footer += "Catatan:<br>Email ini dibuat otomatis oleh sistem, jangan membalas email ini.";

                    dtNotifikasi = lib.CallProcedure("all_createNotifikasiNew", [
                        "SENTTOMARKETING",
                        EncodeData.HtmlEncodeObject(value)[0],
                        "APP58",
                        ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[4]),
                        EncodeData.HtmlEncodeObject(value)[4],
                        "<a href='permintaan_pelanggan'>Permintaan Pelanggan No. " + dt.Rows[0][1].ToString() + " - Selesai Analisa</a>",
                        "[SISFO Produksi] - Tim Engineering dan PPIC Mengirimkan Permintaan Pelanggan No. " + dt.Rows[0][1].ToString() + " (Selesai Analisa)",
                        htmlBody,
                        footer,
                        "Detail",
                        "Permintaan Pelanggan (Selesai Analisa)",
                        "",
                        ""
                    ]);

                    //SEMENTARA KIRIM EMAIL DISINI - BEGIN
                    for (int i = 0; i < dtNotifikasi.Rows.Count; i++)
                    {
                        result = await mail.Send(
                            "[SISFO Produksi] - Tim Engineering dan PPIC Mengirimkan Permintaan Pelanggan No. " + dt.Rows[0][1].ToString() + " (Selesai Analisa)",
                            ldap.GetMail(dtNotifikasi.Rows[i][0].ToString()!),
                            htmlBody.Replace("<<TONAME>>", ldap.GetDisplayName(dtNotifikasi.Rows[i][0].ToString()!)) + footer
                        );
                        if (!result.Equals("OK")) return BadRequest(result);
                    }
                    //SEMENTARA KIRIM EMAIL DISINI - END
                }

                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [SupportedOSPlatform("windows")]
        [HttpPost]
        public async Task<IActionResult> RejectPermintaanPelangganProduk([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_rejectPermintaanPelangganProduk", EncodeData.HtmlEncodeObject(value));

                if (dt.Rows[0][0].ToString()!.Equals("END"))
                {
                    string htmlBody = "";
                    string footer = "";
                    string result;

                    htmlBody += "Yth. Bapak/Ibu <<TONAME>>,<br><br>";
                    htmlBody += "Sistem informasi produksi mencatat ada permintaan pelanggan yang telah selesai dilakukan analisa oleh Engineering dan PPIC dengan rincian sebagai berikut:<br><br>";
                    htmlBody += "Nomor Registrasi Marketing : <b>" + dt.Rows[0][1].ToString() + "</b><br>";
                    htmlBody += "Status : <b>" + dt.Rows[0][2].ToString() + "</b><br><br>";

                    footer += "Silakan lakukan pengecekan pada permintaan pelanggan tersebut dengan cara masuk terlebih dahulu ke <a href='" + configuration["Key:linkRoot"]! + "'>Sistem Informasi Produksi</a>.<br><br>";
                    footer += "<b>SISFO Produksi<br>Politeknik Astra</b><br><br>";
                    footer += "Catatan:<br>Email ini dibuat otomatis oleh sistem, jangan membalas email ini.";

                    dtNotifikasi = lib.CallProcedure("all_createNotifikasiNew", [
                        "SENTTOMARKETING",
                        EncodeData.HtmlEncodeObject(value)[0],
                        "APP58",
                        ldap.GetDisplayName(EncodeData.HtmlEncodeObject(value)[4]),
                        EncodeData.HtmlEncodeObject(value)[4],
                        "<a href='permintaan_pelanggan'>Permintaan Pelanggan No. " + dt.Rows[0][1].ToString() + " - Selesai Analisa</a>",
                        "[SISFO Produksi] - Tim Engineering dan PPIC Mengirimkan Permintaan Pelanggan No. " + dt.Rows[0][1].ToString() + " (Selesai Analisa)",
                        htmlBody,
                        footer,
                        "Detail",
                        "Permintaan Pelanggan (Selesai Analisa)",
                        "",
                        ""
                    ]);

                    //SEMENTARA KIRIM EMAIL DISINI - BEGIN
                    for (int i = 0; i < dtNotifikasi.Rows.Count; i++)
                    {
                        result = await mail.Send(
                            "[SISFO Produksi] - Tim Engineering dan PPIC Mengirimkan Permintaan Pelanggan No. " + dt.Rows[0][1].ToString() + " (Selesai Analisa)",
                            ldap.GetMail(dtNotifikasi.Rows[i][0].ToString()!),
                            htmlBody.Replace("<<TONAME>>", ldap.GetDisplayName(dtNotifikasi.Rows[i][0].ToString()!)) + footer
                        );
                        if (!result.Equals("OK")) return BadRequest(result);
                    }
                    //SEMENTARA KIRIM EMAIL DISINI - END
                }

                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult SetGambarValidasi([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_setGambarValidasi", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult ResetBaseCostProduk([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_resetBaseCostProduk", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CreateBaseCostMaterial([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_createBaseCostMaterial", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CreateBaseCostProses([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_createBaseCostProses", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult CreateBaseCostOther([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_createBaseCostOther", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataBaseCostMaterial([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataBaseCostMaterial", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataBaseCostProses([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataBaseCostProses", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult GetDataBaseCostOther([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                dt = lib.CallProcedure("pro_getDataBaseCostOther", EncodeData.HtmlEncodeObject(value));
                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public async Task<IActionResult> CreateTemplateProses()
        {
            try
            {
                string saveDirectory = Path.Combine(hostingEnvironment.WebRootPath, "Uploads");
                if (!Directory.Exists(saveDirectory))
                    Directory.CreateDirectory(saveDirectory);
                string filePathBase = Path.Combine(saveDirectory, "Template_List_Proses_Base.xlsx");
                string filePathOutput = Path.Combine(saveDirectory, "Template_List_Proses.xlsx");

                using (ExcelPackage pck = new(new FileStream(filePathBase, FileMode.Open)))
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.First();

                    dt = lib.CallProcedure("pro_getListProses", []);
                    ws.Cells[1, 11].LoadFromDataTable(dt, true);

                    var validation = ws.Cells["A:A"].DataValidation.AddListDataValidation();
                    validation.ShowErrorMessage = true;
                    validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                    validation.ErrorTitle = "Error";
                    validation.Error = "Mohon pilih dari daftar yang sudah ada";
                    validation.Formula.ExcelFormula = "$L$2:$L$" + (dt.Rows.Count + 1).ToString();
                    ws.Column(11).Hidden = true;
                    ws.Column(12).Hidden = true;

                    using var stream = new FileStream(filePathOutput, FileMode.Create);
                    await pck.SaveAsAsync(stream);
                }

                return Ok(JsonConvert.SerializeObject(new { Hasil = "Template_List_Proses.xlsx" }));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult ImporBaseCostMaterial([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                string locationDirectory = Path.Combine(hostingEnvironment.WebRootPath, "Uploads");
                string filePath = Path.Combine(locationDirectory, EncodeData.HtmlEncodeObject(value)[0]);

                using (ExcelPackage pck = new(new FileStream(filePath, FileMode.Open)))
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.First();
                    dt = WorksheetToDataTable(ws, "Material");
                }

                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        [Authorize]
        [HttpPost]
        public IActionResult ImporBaseCostProses([FromBody] dynamic data)
        {
            try
            {
                JObject value = JObject.Parse(data.ToString());
                string locationDirectory = Path.Combine(hostingEnvironment.WebRootPath, "Uploads");
                string filePath = Path.Combine(locationDirectory, EncodeData.HtmlEncodeObject(value)[0]);

                using (ExcelPackage pck = new(new FileStream(filePath, FileMode.Open)))
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.First();
                    dt = WorksheetToDataTable(ws, "Proses");
                }

                return Ok(JsonConvert.SerializeObject(dt));
            }
            catch { return BadRequest(); }
        }

        private DataTable WorksheetToDataTable(ExcelWorksheet ws, string jenisTemplate)
        {
            int totalRows = ws.Dimension.End.Row;
            int startRow = 2;
            DataRow dr;
            DataTable dtConverted = new();
            DataTable dtTemp;

            if (jenisTemplate.Equals("Material"))
            {
                dtConverted.Columns.Add("Key");
                dtConverted.Columns.Add("Nama Material/Standar Part");
                dtConverted.Columns.Add("Jenis");
                dtConverted.Columns.Add("Dimensi");
                dtConverted.Columns.Add("Harga Satuan (Rp)");
                dtConverted.Columns.Add("Jumlah");
                dtConverted.Columns.Add("Total Biaya (Rp)");
                dtConverted.Columns.Add("Count");
            }
            else if (jenisTemplate.Equals("Proses"))
            {
                dtConverted.Columns.Add("Key");
                dtConverted.Columns.Add("Nama Proses");
                dtConverted.Columns.Add("Harga Satuan (Rp)");
                dtConverted.Columns.Add("Jumlah");
                dtConverted.Columns.Add("Total Biaya (Rp)");
                dtConverted.Columns.Add("Count");
            }

            for (int rowNum = startRow; rowNum <= totalRows; rowNum++)
            {
                dr = dtConverted.NewRow();

                if (jenisTemplate.Equals("Material"))
                {
                    if (!ws.Cells[rowNum, 1].Text.ToString()!.Trim().Equals("") &&
                        !ws.Cells[rowNum, 2].Text.ToString()!.Trim().Equals("") &&
                        !ws.Cells[rowNum, 3].Text.ToString()!.Trim().Equals("") &&
                        !ws.Cells[rowNum, 4].Text.ToString()!.Trim().Equals("") &&
                        !ws.Cells[rowNum, 5].Text.ToString()!.Trim().Equals(""))
                    {
                        dr.ItemArray = [
                            Guid.NewGuid(),
                            ws.Cells[rowNum, 1].Value.ToString(),
                            ws.Cells[rowNum, 2].Value.ToString(),
                            ws.Cells[rowNum, 3].Value.ToString(),
                            ws.Cells[rowNum, 4].Value.ToString(),
                            ws.Cells[rowNum, 5].Value.ToString(),
                            Convert.ToInt64(ws.Cells[rowNum, 4].Value.ToString()) * Convert.ToDecimal(ws.Cells[rowNum, 5].Value.ToString()),
                            0
                        ];
                        dtConverted.Rows.Add(dr);
                    }
                }
                else if (jenisTemplate.Equals("Proses"))
                {
                    if (!ws.Cells[rowNum, 1].Text.ToString()!.Trim().Equals("") &&
                        !ws.Cells[rowNum, 2].Text.ToString()!.Trim().Equals(""))
                    {
                        DataRow[] existingRows = dtConverted.Select($"[Nama Proses] = '{ws.Cells[rowNum, 1].Value}'");

                        if (existingRows.Length > 0)
                        {
                            existingRows[0]["Jumlah"] = Convert.ToDecimal(existingRows[0]["Jumlah"]) + Convert.ToDecimal(ws.Cells[rowNum, 2].Value.ToString());
                            existingRows[0]["Total Biaya (Rp)"] = Convert.ToDecimal(existingRows[0]["Harga Satuan (Rp)"]) * Convert.ToDecimal(existingRows[0]["Jumlah"]);
                        }
                        else
                        {
                            dtTemp = lib.CallProcedure("pro_getHargaLamaByNamaProses", [ws.Cells[rowNum, 1].Value.ToString()!]);
                            dr.ItemArray = [
                                dtTemp.Rows[0][0].ToString(),
                                ws.Cells[rowNum, 1].Value.ToString(),
                                dtTemp.Rows[0][1].ToString(),
                                ws.Cells[rowNum, 2].Value.ToString(),
                                Convert.ToInt64(dtTemp.Rows[0][1].ToString()) * Convert.ToDecimal(ws.Cells[rowNum, 2].Value.ToString()),
                                0
                            ];
                            dtConverted.Rows.Add(dr);
                        }
                    }
                }
            }

            return dtConverted;
        }
    }
}
