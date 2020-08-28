using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Net.Mail;

namespace Jobs
{
    class Program
    {
        static string fileLocation = @"C:\Users\jepp6960\OneDrive\programmering\praktikpladsensøgeresultat.xlsx";
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage(new FileInfo(fileLocation));
            List<Company> companies = GetAllCompanys(excelPackage);
            string path = "C:/Users/jepp6960/OneDrive/programmering/skp efter h2/Jobs/Jobs/TextFile1.txt";
            List<string> CompSentTo = new List<string>();
            CompSentTo = EmailSent(path);
            companies = Checkcompanies(companies,CompSentTo);

            while (true)
            {
                Console.Clear();
                if (companies.Count>0)
                {
                   
                    using (StreamWriter writer = new StreamWriter(path, true))
                    {
                        writer.WriteLine(companies[0].Name);
                    }

                    Console.WriteLine($"{companies[0].Name}\nAdresse: {companies[0].Adress} {companies[0].Postal}\nTelefon: {companies[0].Phone}\nEmail: {companies[0].Email}");

                    Console.WriteLine("\nÅbn hjemmeside: W \nNæste firma: Enter");
                    Console.WriteLine(companies.Count);
                    companies.RemoveAt(0);
                    
                }
                else
                {
                    Console.WriteLine("no more companies");
                }
                Console.ReadLine();


            }
        }
        private void SendMail()
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

                mail.From = new MailAddress("Jeppe.dehli.sorensen@gmail.com");
                mail.To.Add("to_address");
                mail.Subject = "Test Mail";
                mail.Body = "This is for testing SMTP mail from GMAIL";

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("Jeppe.dehli.sorensen@gmail.com", "password");
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
               Debug.WriteLine("mail Send");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
        }
        private static List<Company> Checkcompanies(List<Company>companies, List<string>Sent)
        {
            for (int i = companies.Count-1; i >= 0; i--)
            {
                for (int j = 0; j < Sent.Count; j++)
                {
                    if (companies[i].Name==Sent[j])
                    {
                        companies.RemoveAt(i);

                    }
                }
            }
            return companies;
        }
        private static List<string> EmailSent(string path)
        {
            string line;
            StreamReader streamReader = new StreamReader(path);
            List<string> companies = new List<string>();
            string all = File.ReadAllText(path);
           
  
            while ((line = streamReader.ReadLine()) != null)
            {
                companies.Add(line);
            }

            streamReader.Close();
            return companies;
            
        }

        private static List<Company> GetAllCompanys(ExcelPackage excelPackage)
        {
            List<Company> Companies = new List<Company>();
            //Finds the first sheet inside the file
            var firstSheet = excelPackage.Workbook.Worksheets.FirstOrDefault();

            int rows = firstSheet.Dimension.Rows;
            string[] information = new string[13];

            for (int i = 2; i < rows; i++)
            {
                for (int j = 0; j < information.Length; j++)
                {
                    information[j] = firstSheet.Cells[i, j + 1].Value.ToString();

                }
                Company company = new Company(information[0], information[1], information[2], information[3], information[4], information[5], information[6], information[7], information[11], Int32.Parse(information[12]));
                string firstNum = company.Postal.Substring(0,1);
                if (string.IsNullOrEmpty(company.Email) == false)
                {
                    if (firstNum == "1"|| firstNum == "2" || firstNum == "4" )
                    {
                        Companies.Add(company);
                    }
                }
            }
            return Companies;
        }


        private static Company GetRandomCompany(ExcelPackage excelPackage)
        {
            //Finds the first sheet inside the file
            var firstSheet = excelPackage.Workbook.Worksheets.FirstOrDefault();

            int rows = firstSheet.Dimension.Rows;
            string[] information = new string[13];

            Random random = new Random();
            int rnd = random.Next(2, rows);
            //Make it easier to insert into company instructor
            for (int i = 0; i < information.Length; i++)
            {
                information[i] = firstSheet.Cells[rnd, i + 1].Value.ToString();

            }
            Company company = new Company(information[0], information[1], information[2], information[3], information[4], information[5], information[6], information[7], information[11], Int32.Parse(information[12]));
            return company;
        }









        //.NET Core wont launch browser with process.start. This method fixes it
        public static void OpenBrowser(string url)
        {
            try
            {
                Process.Start(url);
            }
            catch
            {

                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    url = url.Replace("&", "^&");
                    Process.Start(new ProcessStartInfo("cmd", $"/c start {url}") { CreateNoWindow = true });
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Process.Start("xdg-open", url);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Process.Start("open", url);
                }
                else
                {
                    throw;
                }
            }
        }
    }
}
