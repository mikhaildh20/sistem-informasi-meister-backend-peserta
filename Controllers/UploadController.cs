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
        [HttpPost]
        public async Task<IActionResult> UploadFile(IFormFile file, [FromQuery] string folder = "Uploads")
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return BadRequest("Berkas tidak ada/tidak valid.");
                }

                /*Console.WriteLine($"File diterima: {file.FileName}, Ukuran: {file.Length} bytes");*/

                string fileName = "FILE_" + Guid.NewGuid() + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(file.FileName);

                string uploadDirectory = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", folder);
                if (!Directory.Exists(uploadDirectory))
                    Directory.CreateDirectory(uploadDirectory);

                string filePath = Path.Combine(uploadDirectory, fileName);
                using var stream = new FileStream(filePath, FileMode.Create);
                await file.CopyToAsync(stream);

                /*Console.WriteLine($"File berhasil disimpan di: {filePath}");*/

                return Ok(new { Hasil = fileName });
            }
            catch (Exception ex)
            {
                return BadRequest("Terjadi kesalahan saat mengunggah berkas.");
            }
        }

        [HttpPost]
        public async Task<IActionResult> UploadMultipleFiles([FromForm] List<IFormFile> files, [FromQuery] string folder = "Uploads")
        {
            if (files == null || files.Count == 0)
                return BadRequest("Tidak ada file yang dikirim.");

            List<string> fileNames = new List<string>();
            string uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", folder);

            if (!Directory.Exists(uploadPath))
                Directory.CreateDirectory(uploadPath);

            foreach (var file in files)
            {
                if (file.Length > 0)
                {
                    string fileName = "FILE_" + Guid.NewGuid() + Path.GetExtension(file.FileName);
                    string filePath = Path.Combine(uploadPath, fileName);

                    using var stream = new FileStream(filePath, FileMode.Create);
                    await file.CopyToAsync(stream);

                    fileNames.Add(fileName);
                }
            }

            return Ok(new { hasil = string.Join(",", fileNames) }); // Return file names as a comma-separated string
        }


    }
}
