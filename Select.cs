using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Ecoview_Normal
{
    public partial class Select : Form
    {
  
        public Select()
        {
            InitializeComponent();
 
        }
        bool click = false;
        int selet_rezim;
        private void button2_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                selet_rezim = 1;
            }
            else
            {
                if(radioButton2.Checked == true)
                {
                    selet_rezim = 2;
                }
                else
                {
                    if(radioButton3.Checked == true)
                    {
                        selet_rezim = 3;
                    }
                    else
                    {
                        if(radioButton4.Checked == true)
                        {
                            selet_rezim = 4;
                        }
                        else
                        {
                            if(radioButton5.Checked == true)
                            {
                                selet_rezim = 9;
                            }
                        }
                    }
                }
            }
            Hide();
            Ecoview f2 = new Ecoview(selet_rezim);
            f2.ShowDialog();
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            click = false;
            Application.Exit();
        }

        private void Select_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (click != true)
            {
                
                System.Windows.Forms.Application.ExitThread();

            }
        }
    }
}
