using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace desktopDashboard___Y_Lee.Forms
{
    public partial class pingPc : Form
    {
        public pingPc()
        {
            InitializeComponent();
        }
        private void btnPingPcOK_Click(object sender, EventArgs e)
        {
            string hostname = txtPingPc.Text;
            try
            {
                string[] result = Functions.pingHostname(hostname);
                if (result != null && result[0] != "TimeOut")
                {
                    rtxtPingPc.Text = "Status   : Online"
                                    + "\nHostname : " + hostname.ToUpper()
                                    + "\nIP       : " + result[0];
                }
                else if (result == null)
                    rtxtPingPc.Text = "'" + hostname.ToUpper() + "' is Offline";
                else if (result[0] == "TimeOut")
                    rtxtPingPc.Text = "'" + hostname +  "'" + " is Invalid Entry";
                else
                    rtxtPingPc.Text = "Invalid Entry";
            }
            catch
            {
                rtxtPingPc.Text = "Invalid Entry";
            }
        }
    }
}
