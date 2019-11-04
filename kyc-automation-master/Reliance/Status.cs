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


namespace Reliance
{

    public partial class Status : MetroForm
    {
        public static string ServiceResult = "InProgress";
        string statusText = "";
        List<string> paths = new List<string>();
        string[] statusValues = new string[] { "Checking Policy Number", "Checking Insured Name", "Checking Address", "Checking Phone", "Checking Email", "Checking Registration Number", "Checking Engine Number", "Checking Chassis", "Checking Make And Model", "Checking Body Type", "Checking Year Of Manufacture", "Checking Seating Capacity", "Checking Previous Policy Year", "Checking Previous Policy Company Name", "Checking Vehicle Category", "Checking Vehicle Sub-Category", "Checking Previous Policy RED", "Checking Salutation", "Checking Hypothecation", "Checking RTO" };
        public Status()
        {

            InitializeComponent();
        }
        public static List<string> myPaths;
        public Status(List<string> _paths)
        {
            this.paths = _paths;
            myPaths = _paths;
            InitializeComponent();
        }
        public string resultText
        {
            get
            {
                return content.Text;
            }

            set { label1.Text = value; }
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            axAcroPDF1.src = paths[0];
            axAcroPDF2.src = paths[1];
            Thread.Sleep(2000);
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            Task<bool> task = Method1();
            foreach (string status in statusValues)
            {
                Thread.Sleep(1000);
                statusText = status;
                backgroundWorker1.ReportProgress(1);
            }

        }
        public static async Task<bool> Method1()
        {
            await Task.Run(() =>
            {
                //CustomerDetailsModel cModel = new CustomerDetailsModel();
                //PolicyDetailsModel pModel = new PolicyDetailsModel();
                //RBot rbt = new RBot();
                //foreach (string txt in myPaths)
                //{
                //    if (txt.Contains("Customer"))
                //    {
                //        cModel = rbt.GetPDFFromKYCForm(txt);
                //    }
                //    else
                //    {
                //        pModel = rbt.GetDetailsFromInsuranceCopy(txt);
                //    }
                //}
                //Dictionary<int, int> matchFields = new Dictionary<int, int>();
                //matchFields = rbt.MatchingFields(cModel, pModel);

                //ServiceResult = $"Using {myPaths[0]} and {myPaths[1]}, {matchFields.Keys.FirstOrDefault()} fields matching out of 22 fields and {matchFields.Values.FirstOrDefault()} are empty";
               
            });
            return true;
        }
 
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            content.Text = statusText;

            if (ServiceResult != "InProgress")
            {
                this.Close();
            }
        }
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // this.Close();
        }
    }
}
