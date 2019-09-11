using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ClearInfopathCache
{
    public partial class Form1 : Form
    {
        Process proc = new Process();
        ProcessStartInfo procInfo = new ProcessStartInfo();


        public Form1()
        {
            InitializeComponent();
        }


        public void Form1_Shown(object sender, EventArgs e)
        {
            Task.Factory.StartNew(() =>
            {

                System.Threading.Thread.Sleep(2000);
                Invoke(new Action(RunProgram));

            });
        }


        private async void RunProgram()
        {
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;

            await Task.Run(() => EndSubmitInfopath("INFOPATH"));

            System.Threading.Thread.Sleep(2000);

            this.label1.Invoke((MethodInvoker)delegate
            {
                label1.Text = "Staring Delete of Infopath Cache";
            });

            await Task.Run(() => DeleteInfoCache("<path location to infopath cache>"));

            await Task.Run(() => getLatestTemplates("ASO7"));
            await Task.Run(() => getLatestTemplates("HRO"));
            await Task.Run(() => getLatestTemplates("HnB7"));
            await Task.Run(() => getLatestTemplates("mms"));
            await Task.Run(() => getLatestTemplates("PBA"));
            await Task.Run(() => getLatestTemplates("PACO7"));
            await Task.Run(() => getLatestTemplates("peo"));
            await Task.Run(() => getLatestTemplates("Questionnaire7"));
            await Task.Run(() => getLatestTemplates("esr"));
            await Task.Run(() => getLatestTemplates("CommuterBenefits"));

            this.label1.Invoke((MethodInvoker)delegate
            {
                label1.Text = "Finished Deleting Infopath Cache";
            });

            System.Threading.Thread.Sleep(2000);

            await Task.Run(() => EndSubmitInfopath("SubmitInfoPath"));

            System.Threading.Thread.Sleep(2000);

            await Task.Run(() => setProcessStart());
           


            progressBar1.Visible = false;

        }


        private void DeleteInfoCache(string directory)
        {
            // add in deleting of updated templates and checking for redownload

            System.Threading.Thread.Sleep(1000);

            DirectoryInfo dirInfo = new DirectoryInfo(directory);

            if (Directory.Exists(dirInfo.ToString()))
            {


                foreach (DirectoryInfo dir in dirInfo.GetDirectories())
                {

                    DeleteInfoCache(dirInfo + "\\" + dir.ToString());

                    dir.Delete();
                }

                foreach (FileInfo file in dirInfo.GetFiles())
                {
                    file.Delete();
                }


            }

        }


        private void EndSubmitInfopath(string processName)
        {

            this.label1.Invoke((MethodInvoker)delegate
            {
                label1.Text = "Ending " + processName + " Process";
            });

            System.Threading.Thread.Sleep(4000);

            Process[] process = Process.GetProcessesByName(processName);

            if (process.Length > 0)
            {
                if (process[0].ProcessName == processName)
                {
                    process[0].Kill();
                }
            }

            System.Threading.Thread.Sleep(4000);

            this.label1.Invoke((MethodInvoker)delegate
            {
                label1.Text = "Ended " + processName + "Process";

            });
        }

        private void setProcessStart()
        {

            this.label1.Invoke((MethodInvoker)delegate
            {
                label1.Text = "Starting eForm Applicatoin";
            });

            procInfo.UseShellExecute = false;
            procInfo.FileName = "cmd.exe";
            procInfo.Arguments = "/c \"<start path for application>\"";
            procInfo.CreateNoWindow = true;

            proc.StartInfo = procInfo;

            proc.Start();

            System.Threading.Thread.Sleep(4000);

            this.label1.Invoke((MethodInvoker)delegate
            {
                label1.Text = "eForm Applicatoin has started";
            });

            System.Threading.Thread.Sleep(2000);

            this.label1.Invoke((MethodInvoker)delegate
            {
                label1.Text = "Infopath Clear has completed.";
            });

        }




        public void getLatestTemplates(string tool)
        {
            string[] templatePaths = Directory.GetFiles(@"<path to tools>");
            string latestTemplate = "";
            string delimiter = "_";


            if (tool == "esr" || tool == "CommuterBenefits")
            {
                List<int> templateVersions = new List<int>();

                foreach (string file in templatePaths)
                {
                    Regex pattern = new Regex(@"v\d{1,2}");
                    Match match = pattern.Match(file);

                    string versionNumber = match.Value.Substring(1, match.Value.Length - 1);

                    templateVersions.Add(Int32.Parse(versionNumber));
                }

                int maxVersion = 0;
                int count = 0;
                foreach (int version in templateVersions)
                {
                    if (version > maxVersion)
                    {
                        maxVersion = version;
                        latestTemplate = templatePaths[count];
                    }
                    count++;
                }

                if (File.Exists(latestTemplate))
                {
                    FileInfo tempfile = new FileInfo(latestTemplate);
                    tempfile.Delete();
                }
            }
            else
            {

                if (tool == "PACO7" || tool == "HRO" || tool == "HnB7" || tool == "PBA")
                {
                    delimiter = "-";
                }

                string[] dateFormat = {
                "MM" + delimiter + "dd" + delimiter + "yy",
                "M" + delimiter + "d" + delimiter + "yy",
                "M" + delimiter + "dd" + delimiter + "yy",
                "MM" + delimiter + "d" + delimiter + "yy" };

                List<DateTime> dates = new List<DateTime>();
                foreach (string file in templatePaths)
                {
                    try
                    {
                        //Regex pattern = new Regex(@"(?<month>\d{1,2}$+)_(?<date>\d{1,2})+_(?<year>\d{1,2}+).xsn");
                        Regex pattern = new Regex(@"_\d{1,2}" + delimiter + @"\d{1,2}" + delimiter + @"\d{1,2}");
                        Match match = pattern.Match(file);
                        string date = match.Value.Substring(1, match.Value.Length - 1);

                        dates.Add(DateTime.ParseExact(date, dateFormat, CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None));
                    }
                    catch (Exception)
                    {
                        //Acts as a placeholder to preserve the List index for invalid templates
                        dates.Add(DateTime.ParseExact("01/01/1900", "MM/dd/yyyy", CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None));
                    }
                }

                DateTime minDate = DateTime.MaxValue;
                DateTime maxDate = DateTime.MinValue;
                int count = 0;
                foreach (DateTime date in dates)
                {
                    if (date > maxDate)
                    {
                        maxDate = date;
                        latestTemplate = templatePaths[count];
                    }
                    count++;

                }

                if (File.Exists(latestTemplate))
                {
                    FileInfo tempfile = new FileInfo(latestTemplate);
                    tempfile.Delete();
                }
            }
        }

        
    }
}
