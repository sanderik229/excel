using OfficeOpenXml;
using Spire.Doc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;


namespace WordExceel
{
    /// <summary>
    /// Логика взаимодействия для SendFile.xaml
    /// </summary>
    public partial class SendFile : Window
    {
        private string filename1;
        private string filename2;

        public SendFile(string filename)
        {
            InitializeComponent();
            this.filename1 = filename;
            this.filename2 = filename;

        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Word word = new Word();
            Exel exel = new Exel();
            WordEmail wordEmail = new WordEmail(filename1);
            ExcelEmail excelEmail = new ExcelEmail(filename2);

            if (word.IsActive == true || wordEmail.IsActive == true)
            {
                Document doc = new Document();
                doc.LoadFromFile(filename1);
                doc.SaveToFile(filename1, FileFormat.Docx);
                doc.Close();
                Send(filename1);
                MessageBox.Show("Файл отправлен");
                Close();
            }
            else if (exel.IsActive == true || excelEmail.IsActive == true)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


                using (var package = new ExcelPackage(new System.IO.FileInfo(filename2)))
                {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets[0];
                    package.Save();
                }

                Send(filename2);
                MessageBox.Show("Файл отправлен");
                Close();
            }


        }





        private void Send(string filename)
        {
            MailMessage message = new MailMessage(LoginBx.Text, To.Text, Theme.Text, null);

            SmtpClient smtpClient = new SmtpClient(CheckMail());


            System.Net.Mail.Attachment attachment = new System.Net.Mail.Attachment(filename);

            message.Attachments.Add(attachment);

            smtpClient.Credentials = new NetworkCredential(LoginBx.Text, PasswordBx.Text);
            smtpClient.EnableSsl = true;
            smtpClient.Send(message);

        }


        private string CheckMail()
        {
            if (Combo.SelectedIndex == 1)
            {
                return "smtp.mail.ru";
            }
            else if (Combo.SelectedIndex == 2)
            {
                return "993";
            }
            else if (Combo.SelectedIndex == 3)
            {
                return "imap.gmail.com";
            }
            else if (Combo.SelectedIndex == 0)
            {
                return "imap.rambler.ru";
            }

            return null;
        }
    }
}
