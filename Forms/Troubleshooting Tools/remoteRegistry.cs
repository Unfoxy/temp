using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace desktopDashboard___Y_Lee.Forms
{
    public partial class remoteRegistry : Form
    {
        public remoteRegistry()
        {
            InitializeComponent();
        }
        private void btnRemoteRegistryOK_Click(object sender, EventArgs e)
        {
            string hostname = txtRemoteRegistry.Text;
            string item = comboxRemoteRegistry.Text.ToString();

            try
            {
                if (item == "")
                    rtxtRemoteRegistry.Text = "Please Select Object\n" + "\n1. Fix Globe Sign";
                else
                {
                    string answer = MessageBox.Show("Please Confirm Again " + "PC Name: " + hostname.ToUpper() + "\nCategory: " + item,
                "Desktop Dashboard", MessageBoxButtons.YesNo, MessageBoxIcon.Warning).ToString();
                    if (answer == "Yes")
                    {
                        string[] pingTesting = Functions.pingHostname(hostname);

                        if (pingTesting != null && pingTesting[0] != "TimeOut")
                        {
                            rtxtRemoteRegistry.Text = item + "Request Start                ... 0/6";
                            rtxtRemoteRegistry.AppendText(Environment.NewLine + "Change RemoteRegistry Startup to Manual    ... 1/6");
                            Functions.startupChange(hostname);
                            rtxtRemoteRegistry.AppendText(Environment.NewLine + "Start RemoteRegistry Service               ... 2/6");
                            Functions.startStopService(hostname);
                            rtxtRemoteRegistry.AppendText(Environment.NewLine + "Edit 'IPChecksumOffloadIPv4' Value         ... 3/6");
                            Functions.editRegistry(hostname);
                            rtxtRemoteRegistry.AppendText(Environment.NewLine + "Disable RemoteRegistry Startup to Disabled ... 4/6");
                            Functions.startupChange(hostname);
                            rtxtRemoteRegistry.AppendText(Environment.NewLine + "Stop RemoteRegistry Service                ... 5/6");
                            Functions.startStopService(hostname);
                            rtxtRemoteRegistry.AppendText(Environment.NewLine + item + " Request Completed!          ... 6/6");
                        }
                        else if (pingTesting == null)
                            rtxtRemoteRegistry.Text = "PC Status Error\n" + "\nPC: '" + hostname.ToUpper() + "' Not Online";
                        else
                            rtxtRemoteRegistry.Text = "Invalid Entry\n" + "\nPC: '" + hostname.ToUpper() + "' Not Valid Entry";
                    }
                    else
                        rtxtRemoteRegistry.Text = item + " Request Cancelled";
                }
            }
            catch
            {
                rtxtRemoteRegistry.Text = "Invalid Entry";
            }
        }
    }
}