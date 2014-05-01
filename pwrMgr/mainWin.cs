using System;
using System.Management; //WMI
using System.Collections; //arraylist 
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace pwrMgr
{
    public partial class mainWin : Form
    {
        public mainWin()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Location = new Point(Screen.PrimaryScreen.WorkingArea.Width - this.Width, Screen.PrimaryScreen.WorkingArea.Height-this.Height);
            this.Hide();

            // Add all power plans to tray context menu and combo box
            foreach (PowerPlan p in PowerPlan.getAll())
            {
                trayContextMenuStrip.Items.Add(
                    new System.Windows.Forms.ToolStripMenuItem(p.elementName, null,
                        new EventHandler(trayActivatePlan)));
                if (p.isActive)
                {
                    statusLabel.Text = "Active Plan : " + p.elementName +"\n\n"+ p.description;
                    trayIcon.ShowBalloonTip(20, "Active Power Plan", p.elementName, ToolTipIcon.Info);
                }
            }
            // Sort tray context menu
            ToolStripMenuItem[] tmp = new ToolStripMenuItem[trayContextMenuStrip.Items.Count];
            trayContextMenuStrip.Items.CopyTo(tmp, 0);
            Array.Sort(tmp,
                delegate(ToolStripMenuItem tsiA, ToolStripMenuItem tsiB)
                {
                    return tsiA.Text.Replace("&", string.Empty).CompareTo(tsiB.Text.Replace("&", string.Empty));
                });
            trayContextMenuStrip.Items.Clear();
            trayContextMenuStrip.Items.AddRange(tmp);

            // Add exit to tray context menu
            System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            exitToolStripMenuItem.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            exitToolStripMenuItem.Text = "Exit";
            exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            trayContextMenuStrip.Items.Add(exitToolStripMenuItem);
        }

        private void psMonitorBw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            trayIcon.ShowBalloonTip(20, "Power Plan Changed", (string)e.UserState, ToolTipIcon.Info);
        }
        // end of background worker

        public void trayActivatePlan(Object sender, EventArgs e)
        {
            foreach (PowerPlan p in PowerPlan.getAll())
            {
                if (String.Compare(p.elementName, sender.ToString(), true) == 0
                    && p.isActive == false) // don't switch if already active
                {
                    p.activate();
                    statusLabel.Text = "Active Plan : " + p.elementName + "\n\n" + p.description;
                    trayIcon.ShowBalloonTip(20, "Power Plan Changed", p.elementName, ToolTipIcon.Info);
                    break;
                }
            }
        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        // tray icon click to maximize/minimize
        private void trayIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (this.WindowState == FormWindowState.Minimized)
                {
                    BringToFront(); //place over other windows
                    Show();
                    WindowState = FormWindowState.Normal;
                }
                else if (this.WindowState == FormWindowState.Normal)
                {
                    this.WindowState = FormWindowState.Minimized;
                    Hide();
                }
            }
        }
    }

    /*
     * 
     * 
     *             //background process monitor
            BackgroundWorker psMonitorBw = new BackgroundWorker();
            psMonitorBw.WorkerSupportsCancellation = true;
            psMonitorBw.WorkerReportsProgress = true;
            psMonitorBw.ProgressChanged += new ProgressChangedEventHandler(psMonitorBw_ProgressChanged);
            psMonitorBw.DoWork += new DoWorkEventHandler(psMonitorBw_DoWork);
            if (psMonitorBw.IsBusy != true)
                psMonitorBw.RunWorkerAsync();
/// <summary>
/// Background Worker to monitor for opening processes and apply associated power plan
/// </summary>
/// <param name="sender"></param>
/// <param name="e"></param>
private void psMonitorBw_DoWork(object sender, DoWorkEventArgs e)
{
    BackgroundWorker senderAsBw = sender as BackgroundWorker;
    while (senderAsBw.CancellationPending == false)
    {
        System.Threading.Thread.Sleep(3000);
    }
}*/

    public class PowerPlan
    {
        private string name;
        private string id;
        private string Description;
        private bool active;

        public string elementName
        {
            get
            {
                return name;
            }
        }
        public string instanceID
        {
            get
            {
                return id;
            }
        }
        public string description
        {
            get
            {
                return Description;
            }
        }
        public bool isActive
        {
            get
            {
                return active;
            }
        }

        // constructor
        private PowerPlan(string Name, string ID, string description, bool Active)
        {
            name = Name;
            id = ID;
            Description = description;
            active = Active;
        }

        /// <summary>
        /// Gets active power plan
        /// </summary>
        /// <returns></returns>
        public static PowerPlan getActive()
        {
            string Name = "", ID = "", Description = "";
            try
            {
                ManagementObjectSearcher searcher =
                    new ManagementObjectSearcher("root\\CIMV2\\power",
                    "SELECT * FROM Win32_PowerPlan");

                foreach (ManagementObject queryObj in searcher.Get())
                {
                    if (Convert.ToBoolean(queryObj.Properties["IsActive"].Value) == true)
                    {
                        Name = queryObj.Properties["ElementName"].Value.ToString();
                        ID = queryObj.Properties["InstanceID"].Value.ToString();
                        if (queryObj.Properties["Description"].Value == null)
                            Description = "<No Description>";
                        else
                            Description = queryObj.Properties["Description"].Value.ToString();

                        return new PowerPlan(Name, ID, Description, true);
                    }
                }
            }
            catch (ManagementException e)
            {
                MessageBox.Show(e.ToString());
                return null;
            }
            MessageBox.Show("No active power plan found. WTF?");
            return null; //shld never get here, there is always an active plan :P
        }

        /// <summary>
        /// Returns all power plans in an ArrayList
        /// </summary>
        /// <returns></returns>
        public static ArrayList getAll()
        {
            ArrayList allPlans = new ArrayList();
            try  // WMI Query
            {
                ManagementObjectSearcher searcher =
                    new ManagementObjectSearcher("root\\CIMV2\\power",
                    "SELECT * FROM Win32_PowerPlan");

                foreach (ManagementObject queryObj in searcher.Get())
                {
                    string Name = "", ID = "", Description = "";
                    bool Active;

                    Active = Convert.ToBoolean(queryObj.Properties["IsActive"].Value);
                    Name = queryObj.Properties["ElementName"].Value.ToString();
                    ID = queryObj.Properties["InstanceID"].Value.ToString();
                    if (queryObj.Properties["Description"].Value == null)
                        Description = "<No Description>";
                    else
                        Description = queryObj.Properties["Description"].Value.ToString();

                    allPlans.Add(new PowerPlan(Name, ID, Description, Active));
                }
            }
            catch (ManagementException e)
            {
                MessageBox.Show("An error occurred while querying for WMI data: " + e.Message);
                return null;
            }
            return allPlans;
        }
        /// <summary>
        ///  Activates selected power plan.
        /// </summary>
        /// <param name="property"></param>
        /// <returns></returns>
        public void activate()
        {
            try // WMI Query
            {
                ManagementObject classInstance =
                    new ManagementObject("root\\CIMV2\\power",
                    "Win32_PowerPlan.InstanceID='" + id + "'",
                    null);

                // Execute the method and obtain the return values.
                ManagementBaseObject outParams =
                    classInstance.InvokeMethod("Activate", null, null);
            }
            catch (ManagementException err)
            {
                MessageBox.Show("An error occurred while trying to execute the WMI method: " + err.Message);
            }

        }
    }
}
