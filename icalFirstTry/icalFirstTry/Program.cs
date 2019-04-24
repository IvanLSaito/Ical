using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace icalFirstTry
{
    class Program
    {
        static void Main(string[] args)
        {
            string calendar = CreateCalendar();
            SendEMail(calendar);
        }

        private static void SendEMail(string calendar)
        {
            string to = "jordi.arenas.vigo@gmail.com";
            string from = "ivan.lopez.saito@gmail.com";
            string subject = "Using the new SMTP client.";
            string body = @"Using this new feature, you can send an e-mail message from an application very easily.";
            MailMessage message = new MailMessage(from, to, subject, body);

            string str = CreateCalendar();
            //StringBuilder str = new StringBuilder();
            //str.AppendLine("BEGIN:VCALENDAR");
            //str.AppendLine("PRODID:-//Schedule a Meeting");
            //str.AppendLine("VERSION:2.0");
            //str.AppendLine("METHOD:REQUEST");
            //str.AppendLine("BEGIN:VEVENT");
            //str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", DateTime.Now.AddMinutes(+330)));
            //str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
            //str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", DateTime.Now.AddMinutes(+660)));
            //str.AppendLine("LOCATION: " + "abcd");
            //str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
            //str.AppendLine(string.Format("DESCRIPTION:{0}", msg.Body));
            //str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", msg.Body));
            //str.AppendLine(string.Format("SUMMARY:{0}", msg.Subject));
            //str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", msg.From.Address));

            //str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", msg.To[0].DisplayName, msg.To[0].Address));

            //str.AppendLine("BEGIN:VALARM");
            //str.AppendLine("TRIGGER:-PT15M");
            //str.AppendLine("ACTION:DISPLAY");
            //str.AppendLine("DESCRIPTION:Reminder");
            //str.AppendLine("END:VALARM");
            //str.AppendLine("END:VEVENT");
            //str.AppendLine("END:VCALENDAR");

            byte[] byteArray = Encoding.ASCII.GetBytes(str.ToString());
            MemoryStream stream = new MemoryStream(byteArray);

            System.Net.Mail.Attachment attach = new System.Net.Mail.Attachment(stream, "test.ics");

            message.Attachments.Add(attach);

            System.Net.Mime.ContentType contype = new System.Net.Mime.ContentType("text/calendar");
            contype.Parameters.Add("method", "REQUEST");
            ////  contype.Parameters.Add("name", "Meeting.ics");
            //AlternateView avCal = AlternateView.CreateAlternateViewFromString(str.ToString(), contype);
            //msg.AlternateViews.Add(avCal);

            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Host = "smtp.gmail.com";
            smtpClient.Port = 587;
            //smtpClient.EnableSsl = true;
            smtpClient.Credentials = new System.Net.NetworkCredential("ivan.lopez.saito@gmail.com", "ivanjackson");
            smtpClient.Send(message);
            //https://stackoverflow.com/questions/49358161/create-ics-file-and-send-email-with-attachment-using-c-sharp
            //https://esausilva.com/2016/11/17/create-ical-ics-files-in-c-asp-net-mvc-several-methods/
        }

        private static string CreateCalendar()
        {
            var now = DateTime.Now;
            var later = now.AddHours(1);

            var e = new CalendarEvent
            {
                Start = new CalDateTime(now),
                End = new CalDateTime(later),
            };

            var attendee = new Attendee
            {
                CommonName = "Ivan Saito",
                Rsvp = true,
                Value = new Uri("mailto:ivan.lopez-saito@sogeti.com")
            };
            e.Attendees = new List<Attendee> { attendee };


            var calendar = new Calendar();
            calendar.Events.Add(e);

            var serializer = new CalendarSerializer();
            var icalString = serializer.SerializeToString(calendar);
            return icalString;
        }
    }
}
