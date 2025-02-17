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
        public async Task<IActionResult> UploadFile(IFormFile file, [FromQuery] string folder = "Uploads")
        {
            try
            {
                if (file == null || file.Length == 0)
                    return BadRequest("Berkas tidak ada/tidak valid.");

                // Generate a unique filename
                string fileName = "FILE_" + Guid.NewGuid() + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(file.FileName);

                // Ensure directory exists inside wwwroot
                string uploadDirectory = Path.Combine(hostingEnvironment.WebRootPath, folder);
                if (!Directory.Exists(uploadDirectory))
                    Directory.CreateDirectory(uploadDirectory);

                // Define the file path
                string filePath = Path.Combine(uploadDirectory, fileName);

                // Save file
                using var stream = new FileStream(filePath, FileMode.Create);
                await file.CopyToAsync(stream);

                return Ok(new { Hasil = fileName });
            }
            catch
            {
                return BadRequest("Terjadi kesalahan saat mengunggah berkas.");
            }
        }

    }
}
