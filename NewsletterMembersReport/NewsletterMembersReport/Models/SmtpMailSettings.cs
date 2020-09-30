namespace NewsletterMembersReport.Models
{
    public class SmtpMailSettings
    {
        public string Host { get; set; }
        public int Port { get; set; }
        public bool EnableSsl { get; set; }
        public string Password { get; set; }
        public string SendFrom { get; set; }
        public string SendTo { get; set; }
        public string SendCc { get; set; }
        public string SendBcc { get; set; }
        public string MailSubject { get; set; }
        public string MailBody { get; set; }
    }
}
