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
            this.timer.Interval = 10000;
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
            var filePath = ConfigurationManager.AppSettings["SpreadsheetFolderOriginPath"];
            var destPath = ConfigurationManager.AppSettings["SpreadsheetFolderArchivePath"];
            string[] files = Directory.GetFiles(filePath);

            foreach (var fi in files)
            {
                //Perform the operations based on file type               
                switch (fi)
                {
                    case var fl when (fi.Contains("Avg")):
                        ProcessAvg(filePath, destPath, fi);
                        break;
                    case var fl when (fi.Contains("MAX")):
                        ProcessMAX(filePath, destPath, fi);
                        break;
                    case var fl when (fi.Contains("MIN")):
                        ProcessMIN(filePath, destPath, fi);
                        break;
                    case var fl when (fi.Contains("Detections")):
                        ProcessEvents(filePath, destPath, fi);
                        break;
                    case var fl when (fi.Contains("Availability")):
                        ProcessAvailability(filePath, destPath, fi);
                        break;
                    case var fl when (fi.Contains("patching")):
                        ProcessPatching(filePath, destPath, fi);
                        break;
                    case var fl when (fi.Contains("File_System")):
                        ProcessFileSystem(filePath, destPath, fi);
                        break;
                }
            }
        }


        static void ProcessMIN(string origin, string dest, string fileName)
        {
            //Execute the code for the MIN spreadsheet
            var usage = new DataSources.CMDB();
            try
            {
                usage.ConvertUsageSheetToList(fileName, null, null);
                MoveFile(origin, dest, fileName);
            }
            catch (Exception ex)
            {
                Library.WriteLog(ex.ToString());
            }
        }


        static void ProcessMAX(string origin, string dest, string fileName)
        {
            //Execute the code for the MAX spreadsheet
            var usage = new DataSources.CMDB();
            try
            {
                usage.ConvertUsageSheetToList(null, fileName, null);
                MoveFile(origin, dest, fileName);
            }
            catch (Exception ex)
            {
                Library.WriteLog(ex.InnerException.ToString());
            }
        }


        static void ProcessAvg(string origin, string dest, string fileName)
        {
            //Execute the code for the Avg spreadsheet
            var usage = new DataSources.CMDB();
            try
            {
                usage.ConvertUsageSheetToList(null, null, fileName);
                MoveFile(origin, dest, fileName);
            }
            catch (Exception ex)
            {
                Library.WriteLog(ex.ToString());
            }
        }


        static void ProcessEvents(string origin, string dest, string fileName)
        {
            //Execute the code for the Events spreadsheet
            var events = new DataSources.CMDB();
            try
            {
                events.ConvertEventSheetToList(fileName);
                MoveFile(origin, dest, fileName);
            }
            catch (Exception ex)
            {
                Library.WriteLog(ex.ToString());
            }
        }


        static void ProcessAvailability(string origin, string dest, string fileName)
        {
            //Execute the code for the Availability spreadsheet
            var avail = new DataSources.CMDB();
            try
            {
                avail.ConvertSciLogicSheetToList(fileName);
                MoveFile(origin, dest, fileName);
            }
            catch (Exception ex)
            {
                Library.WriteLog(ex.ToString());
            }
        }


        static void ProcessPatching(string origin, string dest, string fileName)
        {
            //Execute the code for the Patching spreadsheet
            var patch = new DataSources.CMDB();
            try
            {
                patch.ConvertPatchSheetToList(fileName);
                MoveFile(origin, dest, fileName);
            }
            catch (Exception ex)
            {
                Library.WriteLog(ex.ToString());
            }
        }


        static void ProcessFileSystem(string origin, string dest, string fileName)
        {
            //Execute the code for the File System spreadsheet
            var filesys = new DataSources.CMDB();
            try
            {
                filesys.ConvertFileSystemSheetToList(fileName);
                MoveFile(origin, dest, fileName);
            }
            catch (Exception ex)
            {
                Library.WriteLog(ex.ToString());
            }
        }


        //Handle moving the file after it has been processed
        static void MoveFile(string originFolder, string destFolder, string fileName)
        {
            try
            {
                var fileOnly = Path.GetFileName(fileName);
                var destFile = destFolder + "\\" + fileOnly;
                // Ensure that the target does not exist.
                if (File.Exists(destFile))
                    File.Delete(destFile);

                // Move the file.
                File.Move(fileName, destFile);
                //Library.WriteLog(string.Format("{0} was moved to {1}.", originFolder, destFolder));
                Library.WriteLog($"File {fileName} was moved to {destFolder}");

                // See if the original exists now.
                if (File.Exists(originFolder + "\\" + fileName))
                {
                    Library.WriteLog("The original file still exists, which is unexpected.");
                }
                else
                {
                    Library.WriteLog("The original file no longer exists, as expected.");
                }
            }
            catch (Exception e)
            {
                Library.WriteLog($"The process failed: {e.ToString()}");
            }

        }

    }
}
