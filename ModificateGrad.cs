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
    public partial class ModificateGrad : Form
    {
        Ecoview _Analis;
        public int oldValue = 3;
        public int old = 3;
        int index1 = 9;
        public ModificateGrad(Ecoview parent)
        {
            InitializeComponent();

            this._Analis = parent;
            Description.Text = _Analis.Description;
            //    DateTime = _Analis.DateTime;
            Ispolnitel.Text = _Analis.Ispolnitel;
            textBox1.Text = _Analis.direction;
            textBox2.Text = _Analis.code;
            k0Text.Text = _Analis.AgroText0.Text; k1Text.Text = _Analis.AgroText1.Text; k2Text.Text = _Analis.AgroText2.Text;
            // USE_KO = _Analis.USE_KO;
            if (_Analis.USE_KO == true)
            {
                USE_KO.Checked = true;

            }
            else
            {
                USE_KO.Checked = false;
            }

            if (_Analis.Zavisimoct == "A(C)")
            {
                radioButton4.Checked = true;

            }
            else
            {
                radioButton5.Checked = true;
            }

            if (_Analis.radioButton1.Checked == true)
            {
                radioButton1.Checked = true;

                k1Text.Text = string.Format("{0:0.0000}", _Analis.AgroText1.Text);
                //k2Text.Text = string.Format("{0:0.0000}", _Analis.textBox6.Text);
            }
            else
            {
                if (_Analis.radioButton2.Checked == true)
                {
                    radioButton2.Checked = true;

                    k1Text.Text = string.Format("{0:0.0000}", _Analis.AgroText1.Text);
                    k0Text.Text = string.Format("{0:0.0000}", _Analis.AgroText0.Text);
                }
                else
                {
                    radioButton3.Checked = true;

                    k1Text.Text = string.Format("{0:0.0000}", _Analis.AgroText1.Text);
                    k2Text.Text = string.Format("{0:0.0000}", _Analis.AgroText2.Text);
                    k0Text.Text = string.Format("{0:0.0000}", _Analis.AgroText0.Text);
                }
            }
            if (_Analis.SposobZadan == "По СО")
            {
                k0Text.Enabled = false;
                k1Text.Enabled = false;
                k2Text.Enabled = false;
            }
            else
            {
                numericUpDown3.Enabled = false;
                numericUpDown4.Enabled = false;
            }
            WL_grad.Text = _Analis.wavelength1;
            index1 = Ed.FindString(_Analis.edconctr);

            Ed.SelectedIndex = index1;
            int index = Opt_dlin_cuvet.FindString(_Analis.WidthCuvette);
            Opt_dlin_cuvet.SelectedIndex = index;
            Down.Text = _Analis.BottomLine;
            Up.Text = _Analis.TopLine;
            ND.Text = _Analis.ND;
            Description.Text = _Analis.Description;
            dateTimePicker1.Text = _Analis.DateTime;
            numericUpDown1.Value = _Analis.Days;
            Ispolnitel.Text = _Analis.Ispolnitel;
            if (_Analis.CountSeriya != null)
            {
                numericUpDown3.Value = Convert.ToInt32(_Analis.CountSeriya);
            }
            if (_Analis.CountInSeriya != null)
            {
                numericUpDown4.Value = Convert.ToInt32(_Analis.CountInSeriya);
            }
            Veshestvo.Text = _Analis.Veshestvo1;
            textBox2.Text = _Analis.code;
            textBox4.Text = _Analis.textBox3.Text;

            if (_Analis.ComPort == true)
            {
                _Analis.GWNew.Text = WL_grad.Text;
            }
            _Analis.textBox10.Text = WL_grad.Text;
            _Analis.wavelength1 = WL_grad.Text;
            _Analis.Veshestvo1 = Veshestvo.Text;
            _Analis.WidthCuvette = Opt_dlin_cuvet.Text;

            _Analis.textBox3.Text = textBox4.Text;
            _Analis.textBox11.Text = Veshestvo.Text;
            _Analis.wavelength1 = WL_grad.Text;

            switch (_Analis.aproksim) {
                case "Линейная через 0":           
                    radioButton1.Checked = true;
                    break;

                case "Линейная":
                    radioButton2.Checked = true;
                    break;
                default:
                    radioButton3.Checked = true;
                    break;
            }        

            if (_Analis.ComPort == true)
            {
                WL_grad.Text = _Analis.GWNew.Text;
            }
            var height = 22;
            var labelx = 6;
            this.radioButton1.CheckedChanged += new EventHandler(radioButton1_CheckedChanged);
            this.radioButton2.CheckedChanged += new EventHandler(radioButton2_CheckedChanged);
            this.radioButton3.CheckedChanged += new EventHandler(radioButton3_CheckedChanged);
            this.radioButton6.CheckedChanged += new EventHandler(radioButton6_CheckedChanged);
            this.radioButton7.CheckedChanged += new EventHandler(radioButton7_CheckedChanged);
            // numericUpDown4.Value = oldValue;
            for (int i = 0; i <= 9; i++)
            {
                var label = new Label();
                label.Name = "CO" + i++.ToString();
                label.Text = "CO " + i-- + " =";
                label.Width = 40;
                label.Location = new Point(labelx, height);
                height += label.Height;
                groupBox6.Controls.Add(label);
            }
            var height1 = 19;
            var textBoxx = 52;


            for (int i = 0; i <= 9; i++)
            {
                _Analis.textBoxCO[i] = new TextBox();
                _Analis.textBoxCO[i].Name = "COtext" + i++.ToString();
                i--;
                _Analis.textBoxCO[i].Text = Convert.ToString("0,00");
                _Analis.textBoxCO[i].Width = 100;
                _Analis.textBoxCO[i].Height = 20;
                _Analis.textBoxCO[i].Location = new Point(textBoxx, height1);
                height1 += _Analis.textBoxCO[i].Height + 3;
                _Analis.textBoxCO[i].Enabled = false;
                groupBox6.Controls.Add(_Analis.textBoxCO[i]);
                _Analis.textBoxCO[i].KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);
            }
            var height2 = 22;
            var labelx1 = 198;
            for (int i = 10; i <= 19; i++)
            {
                var label = new Label();
                label.Name = "CO" + i++.ToString();
                label.Text = "CO " + i-- + " =";
                label.Width = 40;
                label.Location = new Point(labelx1, height2);
                height2 += label.Height;
                this.Controls.Add(label);
                groupBox6.Controls.Add(label);
            }
            var height3 = 19;
            var textBoxx3 = 244;
            for (int i = 10; i <= 19; i++)
            {
                _Analis.textBoxCO[i] = new TextBox();
                _Analis.textBoxCO[i].Name = "COtext" + i++.ToString();
                i--;
                _Analis.textBoxCO[i].Text = Convert.ToString("0,00");
                _Analis.textBoxCO[i].Width = 100;
                _Analis.textBoxCO[i].Height = 20;
                _Analis.textBoxCO[i].Location = new Point(textBoxx3, height3);
                height3 += _Analis.textBoxCO[i].Height + 3;
                _Analis.textBoxCO[i].Enabled = false;
                groupBox6.Controls.Add(_Analis.textBoxCO[i]);
                _Analis.textBoxCO[i].KeyPress += new System.Windows.Forms.KeyPressEventHandler(txt_KeyPress);
            }
            if (_Analis.SposobZadan != "Ввод коэффициентов")
            {
                //   numericUpDown4.Value = 3;
                for (int i = Convert.ToInt32(numericUpDown4.Value) - 1; i >= 0; i--)
                {
                    this._Analis.textBoxCO[i].Enabled = true;
                }
                if (_Analis.USE_KO == true)
                {
                    USE_KO.Checked = true;
                    for (int j = 0; j < numericUpDown4.Value; j++)
                    {
                        if (_Analis.Stolbec != null)
                        {
                            _Analis.textBoxCO[j].Text = _Analis.Stolbec[j + 1, 1];
                        }
                        if (_Analis.Table1.Rows[j + 1].Cells[1].Value != null)
                        {
                            _Analis.textBoxCO[j].Text = _Analis.Table1.Rows[j + 1].Cells[1].Value.ToString();
                        }
                    }
                }
                else
                {
                    for (int j = 0; j < numericUpDown4.Value; j++)
                    {

                        if (_Analis.Stolbec != null)
                        {
                            _Analis.textBoxCO[j].Text = _Analis.Stolbec[j, 1];
                        }
                        if (_Analis.Table1.Rows[j].Cells[1].Value != null)
                        {
                            _Analis.textBoxCO[j].Text = _Analis.Table1.Rows[j].Cells[1].Value.ToString();
                        }
                    }
                }
            }
            else
            {
                radioButton7.Checked = true;
                switch (_Analis.aproksim)
                {
                    case "Линейная через 0":
                        radioButton1.Checked = true;
                        break;

                    case "Линейная":
                        radioButton2.Checked = true;
                        break;
                    default:
                        radioButton3.Checked = true;
                        break;
                }
            }
        }
        void txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("Только цифры!");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            bool Save = false;
            double f = 0.0;
            for (int i = 0; i < Convert.ToInt32(numericUpDown4.Value); ++i)
            {
                if (Convert.ToDouble(_Analis.textBoxCO[i].Text) <= f && radioButton6.Checked == true && radioButton7.Checked != true)
                {
                    MessageBox.Show("Значение CO не может быть МЕНЬШЕ или РАВНО Нулю!");
                    Save = false;
                    break;
                }
                else
                {
                    Save = true;

                }
                if (WL_grad.Text == "")
                {
                    MessageBox.Show("Заполните поле Длина волны");
                    Save = false;
                    break;
                }
                else
                {
                    Save = true;

                }
            }
            if (Save != false)
            {
                DialogResult result = MessageBox.Show(
                    "Все текущие параметры и данные градуировки будут потеряны. Продолжить?",
                     "Подтверждение",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information,
                     MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                if (result == DialogResult.Yes)
                {
                    _Analis.DateTime = dateTimePicker1.Value.Date.ToString("dd.MM.yyyy");
                    _Analis.dateTimePicker1.Text = dateTimePicker1.Text;
                    _Analis.Ispolnitel = Ispolnitel.Text;
                    _Analis.Description = Description.Text;
                    _Analis.direction = textBox1.Text;
                    _Analis.code = textBox2.Text;
                    _Analis.BottomLine = Down.Text;
                    _Analis.TopLine = Up.Text;
                    _Analis.ND = ND.Text;
                    _Analis.edconctr = Ed.Text;
                    _Analis.Days = Convert.ToInt32(numericUpDown1.Value);
                    _Analis.numericUpDown1.Text = Convert.ToString(numericUpDown1.Value);
                    _Analis.CountSeriya = Convert.ToString(numericUpDown3.Value);
                    _Analis.CountInSeriya = Convert.ToString(numericUpDown4.Value);
                    if(_Analis.ComPort == true)
                    {
                        WL_grad_Leave(sender, e);
                        _Analis.GWNew.Text = string.Format("{0:0.0}", WL_grad.Text);
                        _Analis.wavelength1 = string.Format("{0:0.0}", WL_grad.Text);
                        _Analis.SWСhange();
                        
                    }
                    _Analis.textBox10.Text = string.Format("{0:0.0}", WL_grad.Text);
                    _Analis.wavelength1 = string.Format("{0:0.0}", WL_grad.Text);
                    _Analis.Veshestvo1 = Veshestvo.Text;
                    _Analis.WidthCuvette = Opt_dlin_cuvet.Text;

                    _Analis.textBox3.Text = textBox4.Text;
                    _Analis.textBox11.Text = Veshestvo.Text;
                    _Analis.wavelength1 = WL_grad.Text;


                    switch (radioButton4.Checked)
                    {
                        case true:
                            _Analis.Zavisimoct = "A(C)";
                            break;
                        case false:
                            _Analis.Zavisimoct = "C(A)";
                            break;

                    }
                    switch (radioButton7.Checked)
                    {
                        case true:
                            _Analis.SposobZadan = "Ввод коэффициентов";
                            double k0 = Convert.ToDouble(k0Text.Text);
                            double k1 = Convert.ToDouble(k1Text.Text);
                            double k2 = Convert.ToDouble(k2Text.Text);

                            break;
                        case false:
                            _Analis.SposobZadan = "По СО";
                            break;

                    }
                    if (radioButton1.Checked == true)
                    {
                        _Analis.aproksim = "Линейная через 0";
                    }
                    if (radioButton2.Checked == true)
                    {
                        _Analis.aproksim = "Линейная";
                    }
                    if (radioButton3.Checked == true)
                    {
                        _Analis.aproksim = "Квадратичная";
                    }
                }
                if (USE_KO.Checked == true)
                {
                    _Analis.USE_KO = true;
                }
                else
                {
                    _Analis.USE_KO = false;
                }
                _Analis.label27.Enabled = false;
                _Analis.GradTable();
                Close();
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }
        public void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

            if (radioButton7.Checked == true)
            {
                if (radioButton1.Checked == true && radioButton2.Checked == false && radioButton3.Checked == false)
                {
                    k0Text.Enabled = false;
                    k1Text.Enabled = true;
                    k2Text.Enabled = false;
                    k0Text.Text = string.Format("{0:0.0000}", 0);
                    k2Text.Text = string.Format("{0:0.0000}", 0);

                }
                else
                {
                    if (radioButton2.Checked == true && radioButton1.Checked == false && radioButton3.Checked == false)
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = false;

                        k2Text.Text = string.Format("{0:0.0000}", 0);


                    }
                    else
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = true;



                    }
                }
            }

        }
        public void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

            if (radioButton7.Checked == true)
            {
                if (radioButton1.Checked == true && radioButton2.Checked == false && radioButton3.Checked == false)
                {
                    k0Text.Enabled = false;
                    k1Text.Enabled = true;
                    k2Text.Enabled = false;
                    k0Text.Text = string.Format("{0:0.0000}", 0);
                    k2Text.Text = string.Format("{0:0.0000}", 0);


                }
                else
                {
                    if (radioButton2.Checked == true && radioButton1.Checked == false && radioButton3.Checked == false)
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = false;

                        k2Text.Text = string.Format("{0:0.0000}", 0);


                    }
                    else
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = true;



                    }
                }
            }
        }

        public void radioButton3_CheckedChanged(object sender, EventArgs e)
        {

            if (radioButton7.Checked == true)
            {
                if (radioButton1.Checked == true && radioButton2.Checked == false && radioButton3.Checked == false)
                {
                    k0Text.Enabled = false;
                    k1Text.Enabled = true;
                    k2Text.Enabled = false;
                    k0Text.Text = string.Format("{0:0.0000}", 0);
                    k2Text.Text = string.Format("{0:0.0000}", 0);


                }
                else
                {
                    if (radioButton2.Checked == true && radioButton1.Checked == false && radioButton3.Checked == false)
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = false;

                        k2Text.Text = string.Format("{0:0.0000}", 0);


                    }
                    else
                    {
                        k0Text.Enabled = true;
                        k1Text.Enabled = true;
                        k2Text.Enabled = true;



                    }
                }
            }

        }


        public void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            k0Text.Enabled = false;
            k1Text.Enabled = false;
            k2Text.Enabled = false;
            for (int i1 = 0; i1 < numericUpDown4.Value; i1++)
            {
                _Analis.textBoxCO[i1].Enabled = true;
            }
            numericUpDown3.Enabled = true;
            numericUpDown4.Enabled = true;
            USE_KO.Enabled = true;
        }

        public void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            USE_KO.Enabled = false;

            numericUpDown3.Enabled = false;
            numericUpDown4.Enabled = false;

            for (int i1 = 0; i1 <= numericUpDown4.Value; i1++)
            {
                _Analis.textBoxCO[i1].Enabled = false;
            }

            if (radioButton1.Checked == true && radioButton2.Checked == false && radioButton3.Checked == false)
            {
                k0Text.Enabled = false;
                k1Text.Enabled = true;
                k2Text.Enabled = false;

                k0Text.Text = string.Format("{0:0.0000}", 0);
                k2Text.Text = string.Format("{0:0.0000}", 0);
                k1Text.Text = string.Format("{0:0.0000}", 0);


            }
            else
            {
                if (radioButton2.Checked == true && radioButton1.Checked == false && radioButton3.Checked == false)
                {
                    k0Text.Enabled = true;
                    k1Text.Enabled = true;
                    k2Text.Enabled = false;


                    k0Text.Text = string.Format("{0:0.0000}", 0);
                    k2Text.Text = string.Format("{0:0.0000}", 0);
                    k1Text.Text = string.Format("{0:0.0000}", 0);

                    double k0 = Convert.ToDouble(k0Text.Text);
                    double k1 = Convert.ToDouble(k1Text.Text);
                    double k2 = Convert.ToDouble(k2Text.Text);

                }
                else
                {
                    k0Text.Enabled = true;
                    k1Text.Enabled = true;
                    k2Text.Enabled = true;

                    k0Text.Text = string.Format("{0:0.0000}", 0);
                    k2Text.Text = string.Format("{0:0.0000}", 0);
                    k1Text.Text = string.Format("{0:0.0000}", 0);


                }
            }

        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            _Analis.NoCoIzmer = Convert.ToInt32(numericUpDown4.Value);
            if (numericUpDown4.Value == 1)
            {
                radioButton2.Enabled = false;
                radioButton3.Enabled = false;
            }
            if (numericUpDown4.Value == 2)
            {
                radioButton2.Enabled = true;
                radioButton3.Enabled = false;
            }
            if (numericUpDown4.Value >= 3)
            {
                radioButton2.Enabled = true;
                radioButton3.Enabled = true;
            }



            if (this.oldValue > numericUpDown4.Value)
            {

                for (int i1 = 0; i1 <= 19; i1++)
                {
                    _Analis.textBoxCO[i1].Enabled = false;
                }

                for (int i = _Analis.NoCoIzmer - 1; i >= 0; i--)
                {
                    _Analis.textBoxCO[i].Enabled = true;

                }
            }
            else
            {
                for (int i = _Analis.NoCoIzmer - 1; i >= 1; i--)
                {
                    _Analis.textBoxCO[i].Enabled = true;

                }
            }
            oldValue = _Analis.NoCoIzmer;
        }

        private void WL_grad_Leave(object sender, EventArgs e)
        {
            if (_Analis.ComPort == true)
            {
                if (WL_grad.Text != "")
                {
                    if (_Analis.versionPribor.Contains("V"))
                    {
                        if (Convert.ToDouble(WL_grad.Text.Replace(".", ",")) < 315)
                        {
                            WL_grad.Text = Convert.ToString(315);
                        }
                        if (Convert.ToDouble(WL_grad.Text.Replace(".", ",")) > 1050)
                        {
                            WL_grad.Text = Convert.ToString(1050);
                        }
                    }
                    else
                    {
                        if (_Analis.versionPribor.Contains("U") && _Analis.versionPribor.Contains("2"))
                        {
                            if (Convert.ToDouble(WL_grad.Text.Replace(".", ",")) < 190)
                            {
                                WL_grad.Text = Convert.ToString(190);
                            }
                            if (Convert.ToDouble(WL_grad.Text.Replace(".", ",")) > 1050)
                            {
                                WL_grad.Text = Convert.ToString(1050);
                            }
                        }
                        else
                        {
                            if (Convert.ToDouble(WL_grad.Text.Replace(".", ",")) < 200)
                            {
                                WL_grad.Text = Convert.ToString(200);
                            }
                            if (Convert.ToDouble(WL_grad.Text.Replace(".", ",")) > 1050)
                            {
                                WL_grad.Text = Convert.ToString(1050);
                            }
                        }
                    }
                }
            }
        }
        private void k0Text_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && k0Text.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && k0Text.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && k0Text.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && k0Text.Text.IndexOf('-') != -1) || (number == 43 && k0Text.Text.IndexOf('+') != -1))
            {
                e.Handled = true;
                return;
            }
            if ((e.KeyChar <= 42 || e.KeyChar >= 58 || e.KeyChar == 43 || e.KeyChar == 47) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '-', '.'");
            }
        }

        private void k1Text_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && k1Text.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && k1Text.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && k1Text.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && k1Text.Text.IndexOf('-') != -1) || (number == 43 && k1Text.Text.IndexOf('+') != -1))
            {
                e.Handled = true;
                return;
            }
            if ((e.KeyChar <= 42 || e.KeyChar >= 58 || e.KeyChar == 43 || e.KeyChar == 47) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '-', '.'");
            }
        }

        private void k2Text_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && k2Text.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && k2Text.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && k2Text.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }
            if ((number == 45 && k2Text.Text.IndexOf('-') != -1) || (number == 43 && k2Text.Text.IndexOf('+') != -1))
            {
                e.Handled = true;
                return;
            }
            if ((e.KeyChar <= 42 || e.KeyChar >= 58 || e.KeyChar == 43 || e.KeyChar == 47) && number != 8 && number != 44) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '-', '.'");
            }
        }

        private void WL_grad_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && WL_grad.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && WL_grad.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && WL_grad.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }

            if ((e.KeyChar >= 58 || e.KeyChar <= 47) && number != 8 && number != 44 && number != 46) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '.'");
            }
        }

        private void Down_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && WL_grad.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && WL_grad.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && WL_grad.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }

            if ((e.KeyChar >= 58 || e.KeyChar <= 47) && number != 8 && number != 44 && number != 46) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '.'");
            }
        }

        private void Up_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && WL_grad.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && WL_grad.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && WL_grad.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }

            if ((e.KeyChar >= 58 || e.KeyChar <= 47) && number != 8 && number != 44 && number != 46) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '.'");
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (e.KeyChar == 46 && WL_grad.Text.IndexOf(',') == -1)
            {
                e.KeyChar = ',';

            }
            else
            {

                if (e.KeyChar == 46 && WL_grad.Text.IndexOf(',') != -1)
                {
                    e.Handled = true;
                    return;
                }

            }
            if (number == 44 && WL_grad.Text.IndexOf(',') != -1)
            {
                e.Handled = true;
                return;
            }

            if ((e.KeyChar >= 58 || e.KeyChar <= 47) && number != 8 && number != 44 && number != 46) //цифры, клавиша BackSpace и запятая а ASCII
            {
                e.Handled = true;
                MessageBox.Show("В данное поле можно вводить цифры, знаки '.'");
            }
        }
    }
}
