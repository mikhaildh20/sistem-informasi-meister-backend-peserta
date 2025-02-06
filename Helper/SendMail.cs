using System.Net;
using System.Net.Mail;

namespace sistem_informasi_produksi_backend.Helper
{
    public class SendMail(IConfiguration configuration, IWebHostEnvironment hostingEnvironment)
    {
        public async Task<string> Send(string subjek, string to, string body, string attachment = "")
        {
            try
            {
                MailMessage mail = new()
                {
                    Subject = subjek,
                    Body = body,
                    IsBodyHtml = true,
                    From = new MailAddress(configuration["Key:linkSender"]!)
                };
                mail.To.Add(new MailAddress(to));

                if (attachment != "")
                {
                    mail.Attachments.Add(new Attachment(Path.Combine(hostingEnvironment.WebRootPath, "Uploads", attachment), "application/pdf"));
                }

                SmtpClient message = new(configuration["Key:linkSMTPServer"], Convert.ToInt32(configuration["Key:linkPort"]))
                {
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(configuration["Key:linkUserMail"], configuration["Key:linkPasswordMail"]),
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    EnableSsl = true
                };

                if (configuration["Key:isCanSend"]!.ToString().Equals("1"))
                    await message.SendMailAsync(mail);

                return "OK";
            }
            catch (Exception ex) { return ex.Message; }
        }
    }
}
