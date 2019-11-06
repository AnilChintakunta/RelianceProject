using MetroFramework.Forms;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using Reliance.PDF_Handler;



using OpenQA.Selenium;

using DuoVia.FuzzyStrings;

using System.Windows.Forms;

using System.Text.RegularExpressions;
using AxAcroPDFLib;

namespace Reliance
{

    public partial class Status : MetroForm
    {
        private static IWebBrowser webBrowser;
        private static IEmailHandler emailHandler;
        private static IPDFHandler pdfHandler;
        private static FileHandler fileHandler;
        string browserName = "chrome";

        public static string ServiceResult = "InProgress";
        string statusText = "";
        List<string> paths = new List<string>();
        bool showPdf = false;
        int UnreadmailCount = 0;
        int count = 0;
        bool isMinimize = false;
        string[] statusValues = new string[] { "Checking Policy Number", "Checking Insured Name", "Checking Address", "Checking Phone", "Checking Email", "Checking Registration Number", "Checking Engine Number", "Checking Chassis", "Checking Make And Model", "Checking Body Type", "Checking Year Of Manufacture", "Checking Seating Capacity", "Checking Previous Policy Year", "Checking Previous Policy Company Name", "Checking Vehicle Category", "Checking Vehicle Sub-Category", "Checking Previous Policy RED", "Checking Salutation", "Checking Hypothecation", "Checking RTO" };
        public Status()
        {
            webBrowser = new WebBrowser(browserName);
            emailHandler = new EmailHandler(webBrowser);
            pdfHandler = new PDFHandler(webBrowser);
            fileHandler = new FileHandler(webBrowser);

            InitializeComponent();
        }

        public static List<string> myPaths;
        public Status(List<string> _paths)
        {
            this.paths = _paths;
            myPaths = _paths;
            InitializeComponent();
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            Thread.Sleep(2000);
            this.WindowState = FormWindowState.Minimized;
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            pdfHandler.DeleteFile(webBrowser.downloadPath);
            Dictionary<int, int> matchFields = new Dictionary<int, int>();
           
            emailHandler.LoginAndNavigateToInbox("kartik.rpa@gmail.com", "Kartik1a$");
            webBrowser.Click(WebElements.BtnInsuranceLabel);
            List<IWebElement> unreadMails = emailHandler.GetUnreadEmails();
            UnreadmailCount = unreadMails.Count;
            foreach (IWebElement item in unreadMails)
            {
                count++;
                CustomerDetailsModel cModel = new CustomerDetailsModel();
                PolicyDetailsModel pModel = new PolicyDetailsModel();
                webBrowser.ClickWithJavaScript(item.FindElement(By.CssSelector(WebElements.ChildElementUnreadEmails)));
                Dictionary<string, string> paths = emailHandler.DownloadAndGetFilePaths();
                Thread.Sleep(2000);
                isMinimize = false;
                backgroundWorker1.ReportProgress(1);
                myPaths = paths.Values.ToList();
                showPdf = true;
                backgroundWorker1.ReportProgress(1);

                CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
                CancellationToken token = cancellationTokenSource.Token;
                Task task = Task.Run(() =>
                {
                    while (!token.IsCancellationRequested)
                    {
                        foreach (string status in statusValues)
                        {
                            Thread.Sleep(500);
                            statusText = status;
                            backgroundWorker1.ReportProgress(1);
                        }
                    }
                    //while (token.IsCancellationRequested)
                    //{
                    //    token.ThrowIfCancellationRequested();
                    //}

                }, token);

                //CancellationTokenSource tokenSource = new CancellationTokenSource();
                //CancellationToken token = tokenSource.Token;
                //Task t = Task.Run(() =>
                //  {
                //      foreach (string status in statusValues)
                //      {
                //          Thread.Sleep(500);
                //          statusText = status;
                //          backgroundWorker1.ReportProgress(1);
                //      }

                //  }, token);

                foreach (KeyValuePair<string, string> path in paths)
                {
                    if (path.Value.Contains("Customer"))
                    {
                        cModel = GetPDFFromKYCForm(path.Value);
                    }
                    else
                    {
                        pModel = GetDetailsFromInsuranceCopy(path.Value);
                    }
                }
                matchFields = MatchingFields(cModel, pModel);
                string result = $"Using {paths.Keys.ToArray()[0]} and {paths.Keys.ToArray()[1]}, {matchFields.Keys.FirstOrDefault()} fields matching out of 22 fields and {matchFields.Values.FirstOrDefault()} are empty";
                cancellationTokenSource.Cancel();
                task.Wait();
                // MessageBox.Show(result);
                statusText = result;
                backgroundWorker1.ReportProgress(1);
                Thread.Sleep(1000);
                isMinimize = true;
                backgroundWorker1.ReportProgress(1);
                // fileHandler.Close("AcroRd32");
                webBrowser.Click(WebElements.BtnReply);
                webBrowser.EnterText(WebElements.TxtBoxReplyMessage, "Hello," + "\n" + result + "\nThank You\nReliance Bot");
                webBrowser.Click(WebElements.BtnSend);
                webBrowser.Pause(3);
                webBrowser.Click(WebElements.BtnInsuranceLabel);
                pdfHandler.DeleteFile(webBrowser.downloadPath);
                Thread.Sleep(3000);
              

            }
            ServiceResult = "Completed";
            backgroundWorker1.ReportProgress(1);
            webBrowser.Dispose();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (myPaths.Count > 0 && showPdf == true)
            {
                showPdf = false;
                axAcroPDF1.src = myPaths[0];
                Thread.Sleep(2000);
                axAcroPDF2.src = myPaths[1];
                lblMailCount.Text = "Processing Unread mail " + count + " of " + UnreadmailCount;
                this.WindowState = FormWindowState.Minimized;
            }
            if (isMinimize == true)
            {
                this.WindowState = FormWindowState.Minimized;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
            }
            content.Text = statusText;
            if (ServiceResult == "Completed")
            {
                this.Close();
            }
        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // this.Close();
        }
      
        public static Dictionary<string, string> GetFilePaths(string username, string password)
        {
            pdfHandler.DeleteFile(webBrowser.downloadPath);
            Dictionary<int, int> matchFields = new Dictionary<int, int>();
            emailHandler.LoginAndNavigateToInbox(username, password);
            webBrowser.Click(WebElements.BtnInsuranceLabel);
            List<IWebElement> unreadMails = emailHandler.GetUnreadEmails();
            Dictionary<string, string> paths = new Dictionary<string, string>();
            foreach (IWebElement item in unreadMails)
            {
                CustomerDetailsModel cModel = new CustomerDetailsModel();
                PolicyDetailsModel pModel = new PolicyDetailsModel();
                webBrowser.ClickWithJavaScript(item.FindElement(By.CssSelector(WebElements.ChildElementUnreadEmails)));
                paths = emailHandler.DownloadAndGetFilePaths();
            }
            return paths;
        }

        public static CustomerDetailsModel GetPDFFromKYCForm(string path)
        {
            Stream stream = File.OpenRead(path);
            List<string> pdf = pdfHandler.GetPdfPagesContent(stream);
            var details = pdf.FirstOrDefault();

            int policyFrom = details.IndexOf("Policy Number: ") + "Policy Number: ".Length;
            int policyTo = details.IndexOf("\nInsured");
            string policyNumber = details.Substring(policyFrom, policyTo - policyFrom);


            int insuredFrom = details.IndexOf("Insured Name: ") + "Insured Name: ".Length;
            int insuredTo = details.IndexOf("\nAddress");
            string insuredname = details.Substring(insuredFrom, insuredTo - insuredFrom);

            int addressFrom = details.IndexOf("Address: ") + "Address: ".Length;
            int addressTo = details.IndexOf("\nPhone");
            string address = details.Substring(addressFrom, addressTo - addressFrom);

            int phoneFrom = details.IndexOf("Phone: ") + "Phone: ".Length;
            int phoneTo = details.IndexOf("\nEmail");
            string phone = details.Substring(phoneFrom, phoneTo - phoneFrom);

            int emailFrom = details.IndexOf("Email: ") + "Email: ".Length;
            int emailTo = details.IndexOf("\nRegistration");
            string email = details.Substring(emailFrom, emailTo - emailFrom);

            int registrationFrom = details.IndexOf("Registration Number: ") + "Registration Number: ".Length;
            int registrationTo = details.IndexOf("\nEngine");
            string registrationNumber = details.Substring(registrationFrom, registrationTo - registrationFrom);

            int engineFrom = details.IndexOf("Engine Number: ") + "Engine Number: ".Length;
            int engineTo = details.IndexOf("\nChassis");
            string engineNumber = details.Substring(engineFrom, engineTo - engineFrom);

            int chassisFrom = details.IndexOf("Chassis: ") + "Chassis: ".Length;
            int chassisTo = details.IndexOf("\nMake/Model");
            string chassis = details.Substring(chassisFrom, chassisTo - chassisFrom);

            int makeModelFrom = details.IndexOf("Make/Model: ") + "Make/Model: ".Length;
            int makeModelTo = details.IndexOf("\nType");
            string makeModel = details.Substring(makeModelFrom, makeModelTo - makeModelFrom);

            int typeOfBodyFrom = details.IndexOf("Type of Body: ") + "Type of Body: ".Length;
            int typeOfBodyTo = details.IndexOf("\nYear");
            string typeOfBody = details.Substring(typeOfBodyFrom, typeOfBodyTo - typeOfBodyFrom);

            int yearOfManufactureFrom = details.IndexOf("Year of Manufacture: ") + "Year of Manufacture: ".Length;
            int yearOfManufactureTo = details.IndexOf("\nSeating ");
            string yearOfManufacture = details.Substring(yearOfManufactureFrom, yearOfManufactureTo - yearOfManufactureFrom);

            int seatingCapacityFrom = details.IndexOf("Seating Capacity: ") + "Seating Capacity: ".Length;
            int seatingcapacityTo = details.IndexOf("\nPrevious policy year");
            string seatingCapacity = details.Substring(seatingCapacityFrom, seatingcapacityTo - seatingCapacityFrom);

            int previousPolicyYearFrom = details.IndexOf("Previous policy year: ") + "Previous policy year: ".Length;
            int previousPolicyYearTo = details.IndexOf("\nPrevious policy company name");
            string previousPolicyYear = details.Substring(previousPolicyYearFrom, previousPolicyYearTo - previousPolicyYearFrom);

            int previousPolicyCompanyNameFrom = details.IndexOf("Previous policy company name: ") + "Previous policy company name: ".Length;
            int previousPolicyCompanyNameTo = details.IndexOf("\nVehicle Category");
            string previousPolicyCompamnyName = details.Substring(previousPolicyCompanyNameFrom, previousPolicyCompanyNameTo - previousPolicyCompanyNameFrom);

            int vehicleCategoryFrom = details.IndexOf("Vehicle Category: ") + "Vehicle Category: ".Length;
            int vehicleCategoryTo = details.IndexOf("\nVehicle Sub-Category");
            string vehicleCategory = details.Substring(vehicleCategoryFrom, vehicleCategoryTo - vehicleCategoryFrom);

            int vehicleSubCategoryFrom = details.IndexOf("Vehicle Sub-Category: ") + "Vehicle Sub-Category: ".Length;
            int vehicleSubCategoryTo = details.IndexOf("\nPrevious Policy RED");
            string vehicleSubCategory = details.Substring(vehicleSubCategoryFrom, vehicleSubCategoryTo - vehicleSubCategoryFrom);

            int previousPolicyREDFrom = details.IndexOf("Previous Policy RED: ") + "Previous Policy RED: ".Length;
            int previousPolicyREDTo = details.IndexOf("\nSalutation");
            string previousPolicyRED = details.Substring(previousPolicyREDFrom, previousPolicyREDTo - previousPolicyREDFrom);

            int hypothecationFrom = details.IndexOf("Salutation: ") + "Salutation: ".Length;
            int hypothecationTo = details.IndexOf("\nHypothecation");
            string hypothecation = details.Substring(hypothecationFrom, hypothecationTo - hypothecationFrom);

            int rtoFrom = details.IndexOf("Hypothecation: ") + "Hypothecation: ".Length;
            int rtoTo = details.IndexOf("\nRTO");
            string rto = details.Substring(rtoFrom, rtoTo - rtoFrom);

            int NCB_NCDFrom = details.IndexOf("RTO: ") + "RTO: ".Length;
            int NCB_NCDTo = details.IndexOf("\nNCB/NCD");
            string ncb_ncd = details.Substring(NCB_NCDFrom, NCB_NCDTo - NCB_NCDFrom);


            CustomerDetailsModel cModel = new CustomerDetailsModel();
            cModel.PolicyNumber = policyNumber;
            cModel.InsuredName = insuredname;
            cModel.Address = address;
            cModel.Phone = phone;
            cModel.Email = email;
            cModel.RegistrationNumber = registrationNumber;
            cModel.EngineNumber = engineNumber;
            cModel.Chassis = chassis;
            cModel.Make_Model = makeModel;
            cModel.TypeOfBody = typeOfBody;
            cModel.YearOfManufacture = yearOfManufacture;
            cModel.SeatingCapacity = seatingCapacity;
            cModel.PreviousPolicyYear = previousPolicyYear;
            cModel.PreviousPolicyCompanyName = previousPolicyCompamnyName;
            cModel.VehicleCategory = vehicleCategory;
            cModel.VehicleSubCategory = vehicleSubCategory;
            cModel.PreviousPolicyRED = previousPolicyRED;
            cModel.Hypothecation = hypothecation;
            cModel.RTO = rto;
            cModel.NCB = ncb_ncd;

            return cModel;

        }

        public static PolicyDetailsModel GetDetailsFromInsuranceCopy(string path)
        {
            var client = new RestClient("https://cognitiveservicesprodfun.azurewebsites.net/api/Insurancepolicy/extract?code=D5LEk84GbaOqP1jenKZq6HVrBXVXeKkxeJDPVqoth0WAfYnPJNUlLA==");
            var request = new RestRequest(Method.POST);
            request.AddParameter("Content-Type", "multipart/form-data");
            request.AddFile("file", path, "application/pdf");
            IRestResponse response = client.Execute(request);

            PolicyDetailsModel pModel = JsonConvert.DeserializeObject<PolicyDetailsModel>(response.Content);
            return pModel;
        }

        private static double FuzzyMatch(string word, string targetWord)
        {
            word = string.IsNullOrEmpty(word) ? string.Empty : word.ToLower();
            targetWord = string.IsNullOrEmpty(targetWord) ? string.Empty : targetWord.ToLower();

            int match = word.LevenshteinDistance(targetWord);

            return word.Length >= 1 && targetWord.Length >= 1 ? 1 - (match / Math.Max(word.Length, targetWord.Length)) : double.MinValue;

        }

        public static Dictionary<int, int> MatchingFields(CustomerDetailsModel cModel, PolicyDetailsModel pModel)
        {
            int i = 0;
            int j = 0;

            foreach (var item in cModel.GetType().GetProperties())
            {
                if (item.PropertyType == typeof(string))
                {
                    string value = (string)item.GetValue(cModel);
                    if (string.IsNullOrEmpty(value))
                    {
                        j++;
                    }
                }
            }

            i += Convert.ToInt32(FuzzyMatch(cModel.InsuredName, pModel.Insured) >= 0.7);
            i += Convert.ToInt32(FuzzyMatch(cModel.PolicyNumber, pModel.Number) == 1);
            i += Convert.ToInt32(FuzzyMatch(cModel.Address, pModel.Address) >= 0.7);
            i += Convert.ToInt32(FuzzyMatch(cModel.Phone, pModel.Phone) >= 0.8);
            i += Convert.ToInt32(FuzzyMatch(cModel.Email, pModel.Email) >= 0.8);
            i += Convert.ToInt32(FuzzyMatch(cModel.RegistrationNumber, pModel.RegistrationNumber) == 1);
            i += Convert.ToInt32(FuzzyMatch(cModel.EngineNumber, pModel.EngineNumber) >= 0.8);
            i += Convert.ToInt32(FuzzyMatch(cModel.Chassis, pModel.Chassis) >= 0.8);
            i += Convert.ToInt32(FuzzyMatch(cModel.Make_Model, pModel.Make) >= 0.8);
            i += Convert.ToInt32(FuzzyMatch(cModel.TypeOfBody, pModel.TypeOfBody) >= 0.8);
            i += Convert.ToInt32(FuzzyMatch(cModel.YearOfManufacture, pModel.YearOfManufacturing) == 1);
            i += Convert.ToInt32(FuzzyMatch(cModel.SeatingCapacity, pModel.SeatingCapacity) == 1);
            i += Convert.ToInt32(FuzzyMatch(cModel.PreviousPolicyYear, string.Empty) == 1);
            i += Convert.ToInt32(FuzzyMatch(cModel.PreviousPolicyCompanyName, pModel.CompanyName) >= 0.7);
            i += Convert.ToInt32(FuzzyMatch(cModel.VehicleCategory, pModel.VehicleCategory) >= 0.8);
            i += Convert.ToInt32(FuzzyMatch(cModel.VehicleSubCategory, string.Empty) == 1);
            i += Convert.ToInt32(FuzzyMatch(cModel.PreviousPolicyRED, pModel.PreviousPolicyRED) == 1);
            i += Convert.ToInt32(FuzzyMatch(cModel.Salutation, pModel.Salutation) == 1);
            i += Convert.ToInt32(FuzzyMatch(cModel.Hypothecation, pModel.Hypothecation) >= 0.7);
            i += Convert.ToInt32(FuzzyMatch(cModel.RTO, pModel.RTO) >= 0.8);
            i += Convert.ToInt32(FuzzyMatch(cModel.NCB, pModel.NCB) == 1);

            Dictionary<int, int> result = new Dictionary<int, int>();
            result.Add(i, j);
            return result;
        }

        private int GetNumberOfUnreadEmails()
        {
            string title = "";
            bool status = false;
            while (!title.Contains("Insurance"))
            {
                title = webBrowser.GetWindowTitle();
                Task.Delay(1500);
                break;
            }

            var number = Regex.Match(title, @"[\d]");
            return Convert.ToInt32(number.Value);
        }
    }
}
