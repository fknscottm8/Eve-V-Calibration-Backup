using System;
using System.Data.SqlServerCe;
using System.IO;
using System.IO.Compression;
using System.Management;
using System.Windows.Forms;

namespace Eve_V_Calibration_Backup
{
    public partial class Form1 : Form
    {
        public string CalMANpath = @"C:\ProgramData\SpectraCal\CalMAN Client 3";

        public Form1()
        {
            InitializeComponent();

            foreach (ManagementObject mo in new ManagementObjectSearcher("SELECT PNPDeviceID FROM Win32_DesktopMonitor WHERE DeviceID='DesktopMonitor1'").Get())
                if (mo["PNPDeviceID"].ToString() != null && mo["PNPDeviceID"].ToString().Contains("\\"))
                    label1.Text = mo["PNPDeviceID"].ToString().Split('\\')[2];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(CalMANpath))
            {
                SaveFileDialog sfd = new SaveFileDialog { Filter = "ZIP (*.zip)| *.zip", RestoreDirectory = true };

                if (sfd.ShowDialog() == DialogResult.OK)
                    ZipFile.CreateFromDirectory(CalMANpath, sfd.FileName);
            }
            else
                MessageBox.Show("CalMAN folder missing, client probably not installed. Failed to retreive database.", "Backup Error");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(CalMANpath))
            {
                OpenFileDialog ofd = new OpenFileDialog { Filter = "ZIP (*.zip)| *.zip", RestoreDirectory = true };

                if (ofd.ShowDialog() == DialogResult.OK)
                    ZipFile.ExtractToDirectory(ofd.FileName, CalMANpath);
            }
            else
                MessageBox.Show("CalMAN folder exists, please delete/rename/backup before restore.", "Restore Error");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SqlCeConnection con = new SqlCeConnection(string.Format(@"Data Source={0}\Data\MonitorsDB.2.3.sdf", CalMANpath));

            con.Open();

            new SqlCeCommand(
                string.Format("update Monitors set pnpid = '{0}' where MonitorID = {1}",
                label1.Text,
                new SqlCeCommand("select CreatedMonitorID from Monitors join Profiles on Profiles.CreatedMonitorID = Monitors.MonitorID join (select top 1 ICC_DataID from ICCs order by CreationDate desc) ICCs on ICCs.ICC_DataID = Profiles.ICC_DataID",
                    con).ExecuteScalar().ToString()),
                con).ExecuteScalar();

            con.Close();
        }
    }
}
