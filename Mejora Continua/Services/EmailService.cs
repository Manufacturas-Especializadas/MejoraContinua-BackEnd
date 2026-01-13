using Mejora_Continua.Data;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Extensions.Options;
using MimeKit;
using System.Diagnostics;

namespace Mejora_Continua.Services
{
    public class EmailService
    {
        private readonly EmailSettings _settings;

        public EmailService(IOptions<EmailSettings> settings)
        {
            _settings = settings.Value;
        }

        public async Task SendEmailAsync(string toEmail, string subject, string messageBody)
        {
            var email = new MimeMessage();

            email.From.Add(new MailboxAddress(_settings.SenderName, _settings.SenderEmail));
            email.To.Add(MailboxAddress.Parse(toEmail));
            email.Subject = subject;

            var builder = new BodyBuilder { HtmlBody = messageBody };
            email.Body = builder.ToMessageBody();

            await SendViaSmtpAsync(email);
        }

        public async Task SendGlobalNotificationAsync(string subject, string messageBody, List<string> recipients)
        {
            var email = new MimeMessage();
            email.From.Add(new MailboxAddress(_settings.SenderName, _settings.SenderEmail));

            foreach (var address in recipients)
            {
                if (!string.IsNullOrWhiteSpace(address))
                {
                    email.Bcc.Add(MailboxAddress.Parse(address));
                }
            }

            email.Subject = subject;
            var builder = new BodyBuilder { HtmlBody = messageBody };
            email.Body = builder.ToMessageBody();

            await SendViaSmtpAsync(email);
        }

        private async Task SendViaSmtpAsync(MimeMessage email)
        {
            using var smtp = new SmtpClient();
            smtp.Timeout = 15000;

            try
            {
                await smtp.ConnectAsync(_settings.Host, _settings.Port, SecureSocketOptions.StartTls);

                await smtp.AuthenticateAsync(_settings.Username, _settings.Password);

                await smtp.SendAsync(email);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ERROR EMAIL]: {ex.Message}");
                throw new Exception($"Error enviando correo a través de {_settings.Host}: {ex.Message}");
            }
            finally
            {
                await smtp.DisconnectAsync(true);
            }
        }
    }
}