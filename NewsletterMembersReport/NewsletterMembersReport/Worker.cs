using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MailChimp.Net;
using MailChimp.Net.Interfaces;
using MailChimp.Net.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using NewsletterMembersReport.Models;

namespace NewsletterMembersReport
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly MailChimpSettings _mailChimpSettings;
        private readonly SmtpMailSettings _smtpMailSettings;
        private readonly ReportSettings _reportSettings;

        public Worker(ILogger<Worker> logger, IConfiguration configuration)
        {
            _logger = logger;
            _mailChimpSettings = configuration.GetSection("MailChimp").Get<MailChimpSettings>();
            _smtpMailSettings = configuration.GetSection("SmtpMailSettings").Get<SmtpMailSettings>();
            _reportSettings = configuration.GetSection("Report").Get<ReportSettings>();
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (stoppingToken.IsCancellationRequested == false)
            {
                // Calculate the next execution time
                DateTime nextExecution = CalculateNextExecutionTime();

                // Execute task if next execution time is already in the past
                if (nextExecution <= DateTime.Now.Date)
                {
                    await CreateCsv().ContinueWith((csvFile) =>
                    {
                        SendMail(csvFile.Result);

                        // Save the current date as the last execution time
                        Properties.Settings.Default.LastExecutionTime = DateTime.Now.Date.ToString();
                        Properties.Settings.Default.Save();

                    }, TaskContinuationOptions.OnlyOnRanToCompletion);
                }

                // Delay next execution for x days
                int daysUntilNextExecution = (nextExecution - DateTime.Now.Date).Days;
                for (int i = 0; i < daysUntilNextExecution; i++)
                {// Workaround for maximum int value
                    await Task.Delay(TimeSpan.FromDays(1), stoppingToken);
                }
            }
        }

        private DateTime CalculateNextExecutionTime()
        {
            DateTime output = DateTime.Now;

            if (string.IsNullOrEmpty(Properties.Settings.Default.LastExecutionTime) == false)
            {
                DateTime lastExecutionTime = DateTime.Parse(Properties.Settings.Default.LastExecutionTime);
                output = lastExecutionTime.AddDays(_reportSettings.IntervalInDays);
            }

            return output.Date;
        }

        private async Task<string> CreateCsv()
        {
            _logger.LogInformation($"Start CSV Creation - {DateTime.Now}");

            // Initialize MailChimp Manager
            IMailChimpManager manager = new MailChimpManager(_mailChimpSettings.ApiKey);
            var members = await manager.Members.GetAllAsync(_mailChimpSettings.AudienceListId, new MailChimp.Net.Core.MemberRequest { Status = Status.Subscribed, Limit = 10000 });

            // Get csv file path
            string localAppData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NewsletterMembersReport");
            string csvFilePath = Path.Combine(localAppData, "newsletter_members.csv");

            // Create directory
            if (Directory.Exists(localAppData) == false)
            {
                Directory.CreateDirectory(localAppData);
            }

            // Create csv info
            List<string> csvLines = new List<string>();
            csvLines.Add($"Name,Given Name,Family Name,E-mail 1 - Type,E-mail 1 - Value,Notes");
            foreach (var member in members)
            {
                string firstname = member.MergeFields.First(x => x.Key == "FNAME").Value.ToString();
                string lastname = member.MergeFields.First(x => x.Key == "LNAME").Value.ToString();
                string optIn = DateTime.Parse(member.TimestampOpt).ToString("G");
                csvLines.Add($"{firstname} {lastname},{firstname},{lastname},Home,{member.EmailAddress},Status: {member.Status} - Opt-in: {optIn}");
            }
            File.WriteAllLines(csvFilePath, csvLines, Encoding.UTF8);

            _logger.LogInformation($"Finished CSV Creation - {DateTime.Now}");

            return csvFilePath;
        }

        private void SendMail(string csvFile)
        {
            _logger.LogInformation($"Start Mail Creation - {DateTime.Now}");

            var from = _smtpMailSettings.SendFrom;
            var to = _smtpMailSettings.SendTo;
            var cc = _smtpMailSettings.SendCc;
            var bcc = _smtpMailSettings.SendBcc;
            var pw = _smtpMailSettings.Password;
            var subject = _smtpMailSettings.MailSubject;
            var body = _smtpMailSettings.MailBody;
            var attachment = new Attachment(csvFile);

            var smtp = new SmtpClient
            {
                Host = _smtpMailSettings.Host,
                Port = _smtpMailSettings.Port,
                EnableSsl = _smtpMailSettings.EnableSsl,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(from, pw)
            };

            using (var message = new MailMessage(from, to, subject, body))
            {
                if (string.IsNullOrEmpty(cc) == false)
                {
                    message.CC.Add(cc);
                }
                if (string.IsNullOrEmpty(bcc) == false)
                {
                    message.Bcc.Add(bcc);
                }
                message.Attachments.Add(attachment);
                message.IsBodyHtml = true;
                smtp.Send(message);
            }            

            _logger.LogInformation($"Finished Mail Creation. Mail was sent - {DateTime.Now}");
        }
    }
}
