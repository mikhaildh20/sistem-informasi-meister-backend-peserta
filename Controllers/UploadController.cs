using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;

namespace sistem_informasi_produksi_backend.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class UploadController(IWebHostEnvironment hostingEnvironment) : Controller
    {
        //[Authorize]
        //[HttpPost]
        //public async Task<IActionResult> UploadFile(IFormFile file)
        //{
        //    try
        //    {
        //        string fileName = "FILE_" +
        //            Guid.NewGuid() +
        //            "_" +
        //            DateTime.Now.Year +
        //            DateTime.Now.Month +
        //            DateTime.Now.Day +
        //            DateTime.Now.Hour +
        //            DateTime.Now.Minute +
        //            DateTime.Now.Second;

        //        if (file == null || file.Length == 0)
        //            return BadRequest("Berkas tidak ada/tidak valid.");

        //        string uploadDirectory = Path.Combine(hostingEnvironment.WebRootPath, "Uploads");
        //        if (!Directory.Exists(uploadDirectory))
        //            Directory.CreateDirectory(uploadDirectory);
        //        string filePath = Path.Combine(uploadDirectory, fileName + Path.GetExtension(file.FileName));

        //        using var stream = new FileStream(filePath, FileMode.Create);
        //        await file.CopyToAsync(stream);

        //        return Ok(JsonConvert.SerializeObject(new { Hasil = fileName + Path.GetExtension(file.FileName) }));
        //    }
        //    catch { return BadRequest(); }
        //}

        [HttpPost]
        public async Task<IActionResult> UploadFile(List<IFormFile> files)
        {
            try
            {
                if (files == null || files.Count == 0)
                    return BadRequest("Berkas tidak ada/tidak valid.");

                string[] imageExtensions = { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp" };

                string uploadRoot = Path.Combine(hostingEnvironment.WebRootPath, "Uploads");
                string photoPath = Path.Combine(uploadRoot, "PasPhoto");
                string sertifikatPath = Path.Combine(uploadRoot, "Sertifikat");

                // Buat folder jika belum ada
                if (!Directory.Exists(photoPath)) Directory.CreateDirectory(photoPath);
                if (!Directory.Exists(sertifikatPath)) Directory.CreateDirectory(sertifikatPath);

                string fotoPeserta = "";
                List<string> sertifikatFiles = new List<string>();

                foreach (var file in files)
                {
                    string fileExt = Path.GetExtension(file.FileName).ToLower();
                    string fileName = "FILE_" + Guid.NewGuid() + fileExt;
                    string savePath = "";

                    if (imageExtensions.Contains(fileExt))
                    {
                        savePath = Path.Combine(photoPath, fileName);
                        fotoPeserta = $"PasPhoto/{fileName}"; // Hanya simpan 1 foto terakhir
                    }
                    else
                    {
                        savePath = Path.Combine(sertifikatPath, fileName);
                        sertifikatFiles.Add($"Sertifikat/{fileName}");
                    }

                    // Simpan file
                    using var stream = new FileStream(savePath, FileMode.Create);
                    await file.CopyToAsync(stream);
                }

                return Ok(JsonConvert.SerializeObject(new
                {
                    FotoPeserta = fotoPeserta,
                    SertifikatPeserta = string.Join(",", sertifikatFiles)
                }));
            }
            catch
            {
                return BadRequest("Terjadi kesalahan saat mengunggah file.");
            }
        }
    }
}
