﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Ecoview_Normal
{
    class Izmerenie
    {
        Ecoview _Analis;
        public Izmerenie(Ecoview parent)
        {
            this._Analis = parent;
            _Analis.button1.Enabled = false;
            try
            {
                switch (_Analis.selet_rezim)
                {
                    case 1:
                        if (_Analis.IzmerenieFR_Table.RowCount > 1)
                        {
                            _Analis.IzmerenieFr_izmer();
                            _Analis.label27.Visible = true;
                            _Analis.label28.Visible = false;

                            _Analis.label59.Visible = false;
                            _Analis.Podskazka.Text = "Сохраните измерение";
                            _Analis.button1.Enabled = true;
                        }
                        else
                        {
                            MessageBox.Show("Данная опреция невозможна! Создайте новое измерение!");
                        }
                        break;
                    case 2:
                        if (_Analis.tabControl2.SelectedIndex == 0)
                        {
                            if (_Analis.Table1.RowCount > 1)
                            {
                                if (_Analis.textBox10.Text != _Analis.GWNew.Text)
                                {
                                    MessageBox.Show("Длина волны градуировки отличается от длины волны, установленной на приборе!\rИзмените настройки градуировки!");
                                }
                                _Analis.Graduirovka();

                                _Analis.label59.Visible = false;
                                _Analis.label27.Visible = true;
                                _Analis.label28.Visible = false;
                                _Analis.Podskazka.Text = "Сохраните измерение";
                                _Analis.button1.Enabled = true;
                            }
                            else
                            {
                                MessageBox.Show("Создайте градуировку по СО");
                            }
                        }
                        else
                        {
                            /*if (Table2.RowCount > 1)
                            {
                                if (textBox10.Text != GWNew.Text)
                                {
                                    MessageBox.Show("Длина волны градуировки отличается от длины волны, установленной на приборе!\rИзмените настройки градуировки!");
                                }
                                Izmerenie(sender, e);
                                label27.Visible = true;
                                label28.Visible = false;
                                Podskazka.Text = "Сохраните измерение";
                                button1.Enabled = true;
                            }
                            else
                            {
                                MessageBox.Show("Создайте измерение");
                            }*/
                        }
                        break;
                    case 3:
                        if (_Analis.scan_mass != null)
                        {
                            Application.DoEvents();

                            if (_Analis.dataGridView5.Rows.Count <= 35)
                            {
                                _Analis.dataGridView5.Rows.Add(_Analis.dataGridView5.Rows.Count, "Образец " + _Analis.dataGridView5.Rows.Count);
                                Array.Resize<double[]>(ref _Analis.massGEMultiAbs, _Analis.dataGridView5.Rows.Count - 1);
                                _Analis.massGEMultiAbs[_Analis.massGEMultiAbs.Length - 1] = new double[_Analis.dataGridView5.ColumnCount - 2];
                                Array.Resize<double[]>(ref _Analis.massGEMultiT, _Analis.dataGridView5.Rows.Count - 1);
                                _Analis.massGEMultiT[_Analis.massGEMultiAbs.Length - 1] = new double[_Analis.dataGridView5.ColumnCount - 2];
                                _Analis.StopSpectr = false;
                                _Analis.button14.Enabled = false;
                                _Analis.button11.Enabled = true;
                                _Analis.label33.Visible = true;
                                _Analis.label28.Visible = false;
                                _Analis.label27.Visible = false;
                                _Analis.Podskazka.Text = "Можно остановить";
                                _Analis.button6.Enabled = false;
                                _Analis.button5.Enabled = false;
                                _Analis.button3.Enabled = false;
                                _Analis.button7.Enabled = false;
                                _Analis.button12.Enabled = false;
                                _Analis.button8.Enabled = false;

                                _Analis.TableMultiScan();
                            }
                            else
                            {
                                MessageBox.Show("Достигнут максимум измерений! Создайте новое измерение!");
                            }
                            Application.DoEvents();
                            _Analis.label27.Visible = true;
                            _Analis.Podskazka.Text = "Сохраните измерение";
                            _Analis.label33.Visible = false;
                            _Analis.button7.Enabled = true;
                            _Analis.button3.Enabled = true;
                            //   button8.Enabled = true;
                            _Analis.button1.Enabled = true;
                            _Analis.button6.Enabled = true;
                            _Analis.button5.Enabled = true;
                            //  button3.Enabled = true;
                            //  button7.Enabled = true;
                            _Analis.button12.Enabled = true;
                            _Analis.button8.Enabled = true;

                        }
                        else
                        {
                            MessageBox.Show("Вы забыли откалиброваться! Откалибруйтесь!");
                        }
                        break;
                    case 4:
                        _Analis.ChartGraf();
                        _Analis.countscan = 0;
                        _Analis.timer2.Enabled = false;
                        if (_Analis.timer2.Enabled == false)
                        {
                            if (_Analis.delay > 0)
                            {
                                _Analis.timer1.Interval = Convert.ToInt32(1000); // 500 миллисекунд
                                _Analis.timer1.Enabled = true;
                                _Analis.timer1.Tick += _Analis.TimerTick1;
                            }
                            else {

                                _Analis.label33.Visible = true;
                                _Analis.label28.Visible = false;
                                _Analis.Podskazka.Text = "Можно остановить";
                                _Analis.timeLeft = Convert.ToInt32(_Analis.start);
                                _Analis.TableKinetica1.Rows.Clear();
                                _Analis.label27.Visible = false;
                                _Analis.button14.Enabled = false;
                                _Analis.dataGridView3.Rows.Clear();
                                _Analis.dataGridView4.Rows.Clear();
                                _Analis.timer2.Start();
                                _Analis.timer2.Enabled = true;
                                _Analis.TableKinetica();
                                _Analis.button11.Enabled = true;
                                _Analis.button6.Enabled = false;
                                _Analis.button5.Enabled = false;
                                _Analis.button3.Enabled = false;
                                _Analis.button7.Enabled = false;
                                _Analis.button12.Enabled = false;
                                _Analis.button8.Enabled = false;
                            }

                        }
                        break;
                    case 5:
                        break;
                    case 9:
                        _Analis.label33.Visible = false;
                        _Analis.StopSpectr = false;
                        // Application. = false;
                        //   timePlay = 1;
                        _Analis.Play_Ecxel();
                        _Analis.button1.Enabled = true;
                        break;
                }
            }
            catch
            {


            }
        }

    }
}
