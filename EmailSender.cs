using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.IO;


namespace EmailSender
{
    /* Example of how this is used. */
    //class testProgram
    //{
    //    static void Main(string[] args)
    //    {
    //        Email email = new Email();
    //        email.Address = "test@fakeemail.com";
    //        email.Subject = "testSubject";
    //        email.ToName = "Mr.Testofferson";
    //        email.FromName = "Mr.Bombtastic";
    //        email.Message = "This is a test.<br />This should be on a new line.";
    //        email.AttachmentLocation = "\\\\Location\\filepath\\ScheduledTasks\\Email.cs"; //change this to reflect a file you wish to attach
    //        email.Send();
    //    }
    //}

    public class Email
    {
        private string _address;
        private string _subject;
        private string _toName;
        private string _fromName;
        private string _message;
        private string _attachmentLocation;

        public string Address
        {
            get
            {
                return _address;
            }
            set
            {
                try
                {
                    _address = new MailAddress(value.ToString()).Address;
                }
                catch (FormatException)
                {
                    Console.WriteLine("No Email Address Assigned");
                }
            }
        }

        public string Subject
        {
            get
            {
                return _subject;
            }
            set
            {
                if (value.Length > 0)
                {
                    _subject = value;
                }
                else
                {
                    throw new Exception("No Subject");
                }
            }
        }

        public string ToName
        {
            get
            {
                return _toName;
            }
            set
            {
                if (value.Length > 0)
                {
                    _toName = value;
                }
                else
                {
                    throw new Exception("No To Name Assigned");
                }
            }
        }

        public string FromName
        {
            get
            {
                return _fromName;
            }
            set
            {
                if (value.Length > 0)
                {
                    _fromName = value;
                }
                else
                {
                    throw new Exception("No From Name Assigned");
                }
            }
        }

        public string Message
        {
            get
            {
                return _message;
            }
            set
            {
                if (value.Length > 0)
                {
                    _message = value;
                }
                else
                {
                    throw new Exception("No message Assigned");
                }
            }
        }

        public string AttachmentLocation
        {
            get
            {
                return _attachmentLocation;
            }
            set
            {
                if (File.Exists(value.ToString()))
                {
                    _attachmentLocation = value;
                }
                else
                {
                    throw new Exception("No File at Location");
                }
            }
        }


        public void Send()
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.domain"); // change to apprpriate smtp client

                mail.From = new MailAddress("NoReply_Automated@fakeemail.com"); // recipient will see this email
                mail.To.Add(Address);
                mail.Subject = Subject;
                mail.IsBodyHtml = true;

                string body = string.Format(@"
                                    <p>To: {0},</p>
                                    <br />
                                    <p>{1}</p>
                                    <br />
                                    <p>Thank You,</p>
                                    <p>{2}</p>
                                    <br /><br />
                                    <p>This is an automated message. For any inqueries, please contact xxxxxxxx@xxxxxx.com</p>
                                    ", ToName, Message, FromName);

                var view = AlternateView.CreateAlternateViewFromString(body, null, "text/html");

                mail.AlternateViews.Add(view);

                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(AttachmentLocation);
                mail.Attachments.Add(attachment);

                SmtpServer.Port = 587;// change to appropriate port
                SmtpServer.UseDefaultCredentials = false;


                SmtpServer.Credentials = new System.Net.NetworkCredential("emailaddress@fakeemail.com", "fakeEmailPassword1"); //account and password
                SmtpServer.EnableSsl = false; //internal does not require SSL.

                SmtpServer.Send(mail);
                Console.WriteLine("File: " + AttachmentLocation + ", sent to " + ToName + " at " + Address);
                Console.Read();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Console.Read();
            }

        }
    }
}
