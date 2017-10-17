using System.Net;
using System.Net.Mail;

namespace NewEplazaTest
{
    public class SendMail
    {
        public void Mail()
        {
            MailAddress fromMailAddress = new MailAddress("boteplaza@gmail.com", "Test");
            MailAddress toAddress = new MailAddress("sag@m-st.ru", "Uncle Bob");
            using (MailMessage mailMessage = new MailMessage(fromMailAddress, toAddress))
            using (SmtpClient smtpClient = new SmtpClient())
            {
                mailMessage.Subject = "Отчет по автотестированию Еплазы";
                mailMessage.Body = "Откройте документ";
                //прикрепляем вложение
                Attachment attData = new Attachment("C:/Users/new/Documents/Visual Studio 2017/Projects/NewEplazaTest/NewEplazaTest/bin/Debug/Test.xlsx");
                mailMessage.Attachments.Add(attData);

                smtpClient.Host = "smtp.gmail.com";
                smtpClient.Port = 587;
                smtpClient.EnableSsl = true;
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new NetworkCredential(fromMailAddress.Address, "123456eplaza");
                smtpClient.Send(mailMessage);
            }
        }
    }
}
