using System;
using System.Configuration;
using System.IO;
using System.Net.Mail;

namespace AllSalesReport
{
    static class MailSender
    {
        public static void SendMail(string subject, string body, string mailToAddress, Stream contentStream, string fileName, string mediaType)
        {
            string fromAddress = ConfigurationManager.AppSettings["Mail_From_Address"];
            string fromDisplayName = ConfigurationManager.AppSettings["Mail_From_Display_Name"];
            string[] toAddresses = mailToAddress.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            string smtpServer = ConfigurationManager.AppSettings["Mail_Smtp_Server"];

            MailMessage mailMessage = new MailMessage();
            mailMessage.From = new MailAddress(fromAddress, fromDisplayName);
            foreach (var toAddress in toAddresses)
                mailMessage.To.Add(new MailAddress(toAddress.Trim()));
            mailMessage.Subject = subject;
            mailMessage.Priority = MailPriority.Normal;

            if (string.IsNullOrEmpty(body) == false)
                mailMessage.Body = body;

            Attachment attachment = null;
            if (contentStream != null)
            {
                attachment = new Attachment(contentStream, fileName, mediaType);
                mailMessage.Attachments.Add(attachment);
            }

            SmtpClient smtp = new SmtpClient(smtpServer);
            if (smtpServer != "192.168.78.232")
                smtp.UseDefaultCredentials = true;
            smtp.Send(mailMessage);

            if (attachment != null)
                attachment.Dispose();
            mailMessage.Dispose();
        }
    }
}
