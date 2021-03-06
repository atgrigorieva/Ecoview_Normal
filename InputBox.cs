﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Ecoview_Normal
{
    public partial class InputBox : Form
    {
        Ecoview _Analis;
        public InputBox(Ecoview parent)
        {
            InitializeComponent();
            this._Analis = parent;
            textBox1.Focus();
            textBox1.Select();
            if (_Analis.CellOpt != 0)
            {
                textBox1.Text = Convert.ToString(_Analis.CellOpt);
            }
            else
            {
                textBox1.Text = Convert.ToString(_Analis.CellOpt);
            }
        }

        private void Save_Click(object sender, EventArgs e)
        {

            if (textBox1.Text != "")
            {
                _Analis.CellOpt = Convert.ToDouble(textBox1.Text);
                Close();
            }
            else
            {
                MessageBox.Show("Введите значение!");
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {

            Close();
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (number == 13)
            {
                Save_Click(sender, e);
            }
            if (number == 27)
            {
                Cancel_Click(sender, e);
            }
            if (e.KeyChar == 46 && textBox1.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && textBox1.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && textBox1.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && textBox1.Text.IndexOf('-') != -1) || (number == 43 && textBox1.Text.IndexOf('+') != -1))
            {
                e.Handled = true;
                return;
            }
            if ((e.KeyChar <= 42 || e.KeyChar >= 58 || e.KeyChar == 43 || e.KeyChar == 47) && number != 8 && number != 44 && number != 13 && number != 27) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '-', '.'");
            }

        }
    }
}
