using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Net.Mail;

namespace PilotDispatchExtensions
{
    public class Utils
    {
        public Utils()
        {
            
        }

        public string[] SendEmails(string fromEmailAddress, string hostAddress, string port,
            string ssl, string login, string password, string title,
            string body, string messageType, ArrayList alEmail,
            ArrayList alAttachments)
        {
            var sw = new Stopwatch();
            var m = new List<MailMessage>();
            sw.Start();

            try
            {
                var smtpFromEmail = (fromEmailAddress ?? string.Empty);
                var smtpHost = (hostAddress ?? string.Empty);
                var smtpPort = (port ?? string.Empty);
                var smtpEnableSll = (ssl == "T" ? true : false);
                var smtpLogin = (login ?? string.Empty);
                var smtpPassword = (password ?? string.Empty);
                var s = 0;

                //foreach (var mail in from object t in alEmail
                //    select t.ToString()
                //    into item
                //    where !string.IsNullOrEmpty(item)
                //    select new MailMessage(smtpFromEmail, item, title, body))
                //{

                MailMessage mail = new MailMessage();

                mail.From = new MailAddress(fromEmailAddress);
                mail.Body = body;
                mail.Subject = title;
                mail.IsBodyHtml = true;

                for (int i = 0; i <= alEmail.Count - 1; i++)
                {
                    if (alEmail[i] != null)
                    {
                        mail.To.Add(new MailAddress(alEmail[i].ToString()));
                        s++;
                    }  
                }

                if (messageType.Trim() == "E")
                {
                    foreach (var t in alAttachments)
                    {
                        var filepath = t.ToString().Trim();
                        if (string.IsNullOrEmpty(filepath))
                        {
                            continue;
                        }

                        var data = new Attachment(filepath);

                        // Add time stamp information for the file.
                        var disposition = data.ContentDisposition;
                        disposition.CreationDate = System.IO.File.GetCreationTime(filepath);
                        disposition.ModificationDate = System.IO.File.GetLastWriteTime(filepath);
                        disposition.ReadDate = System.IO.File.GetLastAccessTime(filepath);

                        // Add the file attachment to this e-mail message.
                        mail.Attachments.Add(data);
                        }
                    }

                m.Add(mail);
 
                var smtp = new SmtpClient();

                if (smtpHost != "")
                    smtp.Host = smtpHost;
                if (smtpPort != "")
                    smtp.Port = Convert.ToInt32(smtpPort);
                if (smtpLogin != "" & smtpPassword != "")
                    smtp.Credentials = new System.Net.NetworkCredential(smtpLogin, smtpPassword);
                smtp.EnableSsl = smtpEnableSll;

                
                foreach (var t in m)
                {
                    smtp.Send(t);
                }

                sw.Stop();
                var te = sw.ElapsedMilliseconds;

                var results = new String[]
                {s.ToString(CultureInfo.InvariantCulture), te.ToString(CultureInfo.InvariantCulture)};

                return results;
            }
            catch (Exception e)
            {
                var results = new String[]
                {
                    e.Message.ToString(CultureInfo.InvariantCulture),
                    e.InnerException.ToString()
                };

                return results;
            }
        }
    }
}
