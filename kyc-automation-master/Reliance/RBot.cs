
using Newtonsoft.Json.Linq;
using System.Linq;
using OpenQA.Selenium;
using Reliance.PDF_Handler;
using System.Collections.Generic;
using DuoVia.FuzzyStrings;
using System.IO;
using RestSharp;
using Newtonsoft.Json;
using System;
using System.Windows.Forms;
using System.Threading;
using OpenQA.Selenium.Interactions;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Reliance
{
    public class RBot
    {
        private readonly IWebBrowser webBrowser;
        private readonly IEmailHandler emailHandler;
        private IPDFHandler pdfHandler;
        private readonly FileHandler fileHandler;
        public RBot()
        {
          
        }
        public RBot(string browserName)
        {
            webBrowser = new WebBrowser(browserName);
            emailHandler = new EmailHandler(webBrowser);
            pdfHandler = new PDFHandler(webBrowser);
            fileHandler = new FileHandler(webBrowser);
        }
        public void RunTheBot(string username, string password)
        {
            pdfHandler.DeleteFile(webBrowser.downloadPath);
            Dictionary<int, int> matchFields = new Dictionary<int, int>();
            emailHandler.LoginAndNavigateToInbox(username, password);
            webBrowser.Click(WebElements.BtnInsuranceLabel);
            List<IWebElement> unreadMails = emailHandler.GetUnreadEmails();
            foreach (IWebElement item in unreadMails)
            {
                CustomerDetailsModel cModel = new CustomerDetailsModel();
                PolicyDetailsModel pModel = new PolicyDetailsModel();
                webBrowser.ClickWithJavaScript(item.FindElement(By.CssSelector(WebElements.ChildElementUnreadEmails)));
                Dictionary<string, string> paths = emailHandler.DownloadAndGetFilePaths();
                //fileHandler.Open(paths.Values.ToList());

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                var statusFrom = new Status(paths.Values.ToList());
                Application.Run(statusFrom);

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
                // statusFrom.resultText = result;

                // fileHandler.Close("AcroRd32");
                webBrowser.Click(WebElements.BtnReply);
                webBrowser.EnterText(WebElements.TxtBoxReplyMessage, "Hello," + "\n" + result + "\nThank You\nReliance Bot");
                webBrowser.Click(WebElements.BtnSend);
                webBrowser.Pause(3);
                webBrowser.Click(WebElements.BtnInsuranceLabel);
                pdfHandler.DeleteFile(webBrowser.downloadPath);
            }
            webBrowser.Dispose();
        }

        public CustomerDetailsModel GetPDFFromKYCForm(string path)
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

        public PolicyDetailsModel GetDetailsFromInsuranceCopy(string path)
        {
            var client = new RestClient("https://cognitiveservicesprodfun.azurewebsites.net/api/Insurancepolicy/extract?code=D5LEk84GbaOqP1jenKZq6HVrBXVXeKkxeJDPVqoth0WAfYnPJNUlLA==");
            var request = new RestRequest(Method.POST);
            request.AddParameter("Content-Type", "multipart/form-data");
            request.AddFile("file", path, "application/pdf");
            IRestResponse response = client.Execute(request);

            PolicyDetailsModel pModel = JsonConvert.DeserializeObject<PolicyDetailsModel>(response.Content);
            return pModel;
        }

        private double FuzzyMatch(string word, string targetWord)
        {
            word = string.IsNullOrEmpty(word) ? string.Empty : word.ToLower();
            targetWord = string.IsNullOrEmpty(targetWord) ? string.Empty : targetWord.ToLower();

            int match = word.LevenshteinDistance(targetWord);

            return word.Length >= 1 && targetWord.Length >= 1 ? 1 - (match / Math.Max(word.Length, targetWord.Length)) : double.MinValue;

        }

        public Dictionary<int, int> MatchingFields(CustomerDetailsModel cModel, PolicyDetailsModel pModel)
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
