# mailchimp-member-report
The mailchimp-member-report is a worker service that generates a monthly report of the currently subscribed members of a mailchimp newsletter.

The report is automatically send via email.

## Motivation
I created this service to automate the generation of a newsletter member report, I otherwise would need to do manually.

## Technologies used
- C#
- .NET Core 3.1
- Worker Service (Microsoft.Extensions.Hosting.BackgroundService)
- MailChimp API
- System.Net.Mail (SmtpClient)

## How to use
The service can be installed as a Windows Service. For further information please visit: [Microsoft Docs](https://docs.microsoft.com/en-us/dotnet/framework/windows-services/how-to-install-and-uninstall-services)

### Service configuration
You can change the **appsettings.json** file to configure the MailChimp API, the SMTP Mail settings and the interval of the report.

Change the report interval under the section **Report**.

```
"Report": {
    "IntervalInDays": 30 
}
```

Add your MailChimp API settings under the section **MailChimp**.

```
"MailChimp": {
    "ApiKey": "",
    "AudienceListId": ""
}
```

Add your SMTP Mail settings under the section **SmtpMailSettings**.

```
"SmtpMailSettings": {
    "Host": "",
    "Port": "",
    "EnableSsl": true,
    "Password": "",
    "SendFrom": "",
    "SendTo": "",
    "SendCc": "",
    "SendBCc": "",
    "MailSubject": "",
    "MailBody": ""
}
```

The *MailBody* also supports HTML tags.

**Tip:** If you are using GMail you can create a new Application Key and use this as the SMTP Password.