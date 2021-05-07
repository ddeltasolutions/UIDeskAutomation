using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace UIDeskAutomationLib
{
    internal partial class FrmDummy : Form
    {
        public FrmDummy()
        {
            InitializeComponent();

            timer.AutoReset = false;
            timer.Elapsed += timer_Elapsed;
        }

        private void FrmDummy_Load(object sender, EventArgs e)
        {
            
        }

        private void SetTimer()
        {
            if (timer.Enabled == true)
            {
                return;
            }

            try
            {
                timer.Start();
            }
            catch { }
        }

        void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                this.Close();
            }
            catch { }
        }

        System.Timers.Timer timer = new System.Timers.Timer(200);
    }
}
