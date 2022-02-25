using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace MailService
{
    public partial class Service1 : ServiceBase
    {
        private Timer timer;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            if (timer == null)
            {
                timer = new Timer();
            }
            timer.Interval = 60*1000;
            timer.Elapsed += TimerElapsed;
            LogToFile("Windows service is started at " + DateTime.Now);
            timer.Start();
        }

        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            LogToFile("Windows service is running " + DateTime.Now);
            SendMail();
        }

        protected override void OnStop()
        {
            LogToFile("Windows service is stopped at " + DateTime.Now);
        }

        public void SendMail()
        {
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Host = "smtp.gmail.com";
            smtpClient.Port = 587;
            smtpClient.Credentials = new NetworkCredential("senderMailAdress@gmail.com", "senderMailPassword");
            smtpClient.EnableSsl = true;

            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("senderMailAdress@gmail.com");
            mail.To.Add("receiverMailAdress1@windowslive.com");
            mail.To.Add("receiverMailAdress2@gmail.com");
            mail.Subject = "Windows Service Test";
            mail.Body = $"This message was sent by windows service.";

            try
            {
                smtpClient.Send(mail);
                LogToFile("Mail sent successfully " + DateTime.Now); 
            }
            catch (Exception ex)
            {
                LogToFile("ERROR: " + ex.Message);
            }
        }

        public void LogToFile(string message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + @"\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + @"\Logs\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/','_') + ".txt";
            if (!File.Exists(filepath))
            {
                using (StreamWriter streamWriter = File.CreateText(filepath))
                {
                    streamWriter.WriteLine(message);
                }
            }
            else
            {
                using (StreamWriter streamWriter = File.AppendText(filepath))
                {
                    streamWriter.WriteLine(message);
                }
            }
        }
    }
}
