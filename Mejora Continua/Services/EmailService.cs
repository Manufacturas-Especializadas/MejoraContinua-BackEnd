using Mejora_Continua.Data;
using Microsoft.Extensions.Options;
using System.Net;
using System.Net.Mail;

namespace Mejora_Continua.Services
{
    public class EmailService
    {
        public EmailSettings Setting { get; }

        public EmailService(IOptions<EmailSettings> options)
        {
            Setting = options.Value;
        }

        public async Task SendEmailAsync(string toEmail, string subject, string body)
        {
            var mailMessage = new MailMessage
            {
                From = new MailAddress(Setting.SenderEmail, Setting.SenderName),
                Subject = subject,
                Body = body,
                IsBodyHtml = true,
                Priority = MailPriority.High,
            };

            mailMessage.To.Add(toEmail);

            using var client = new SmtpClient
            {
                Host = Setting.Host,
                Port = Setting.Port,
                EnableSsl = Setting.UseSSL,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(Setting.Username, Setting.Password)
            };

            await client.SendMailAsync(mailMessage);
        }
    }
}