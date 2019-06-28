using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Configuration;
using GLAReportMVC;


namespace SpreadsheetProcessor
{
    public partial class Scheduler : ServiceBase
    {
        private Timer timer = null;
        public Scheduler()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            timer = new Timer(); 
            //this.timer.Interval = 3600000; //hourly
            this.timer.Interval = 500;
            this.timer.Elapsed += new System.Timers.ElapsedEventHandler(this.timer_Tick);
            timer.Enabled = true;
            Library.WriteLog("Spreadsheet processor service started");
        }

        protected override void OnStop()
        {
        }

        private void timer_Tick(object sender, ElapsedEventArgs e)
        {
            //Service action
            var fileP = ConfigurationManager.AppSettings["SpreadsheetFolderOriginPath"];
            var filePath = @"C:\Users\mrutt\Documents\GLASpreadsheets";
            string[] files = Directory.GetFiles(filePath);

            foreach (var fi in files)
            {
                //Perform the operations based on file type               
                switch(fi)
                {
                    case var fl when (fi.Contains("MIN")):
                        ProcessMIN(fi);
                        break;
                    case var fl when (fi.Contains("MAX")):
                        ProcessMAX(fi);
                        break;
                    case var fl when (fi.Contains("Avg")):
                        ProcessAvg(fi);
                        break;
                }
            }
        }

        private void ProcessMIN(string fileName)
        {
            //Execute the code for the MIN spreadsheet
            var usage = new DataSources.CMDB();
            try
            {
                usage.ConvertUsageSheetToList(fileName, null, null);
            }
            catch (Exception ex)
            {
                Library.WriteLog(ex.ToString());
            }
        }


        private void ProcessMAX(string fileName)
        {
            //Execute the code for the MAX readsheet
            var usage = new DataSources.CMDB();
            try
            {
                usage.ConvertUsageSheetToList(null, fileName, null);
            }
            catch (Exception ex)
            {
                Library.WriteLog(ex.InnerException.ToString());
            }
        }

        private void ProcessAvg(string fileName)
        {
            //Execute the code for the MAX readsheet
            var usage = new DataSources.CMDB();
            try
            {
                usage.ConvertUsageSheetToList(null, null, fileName);
            }
            catch (Exception ex)
            {
                Library.WriteLog(ex.ToString());
            }
        }

    }
}
