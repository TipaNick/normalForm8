﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Bytescout.Spreadsheet;
using System.IO;

namespace normalForm8
{
    public partial class Form1 : Form
    {
        string examplePath = "C:\\Users\\TipaNick\\Desktop\\normalForm\\unform8.xls";
        string forCreatePath = "C:\\Users\\TipaNick\\Desktop\\normalForm\\finish.xls";
        Spreadsheet spreadsheet = new Spreadsheet();
        string employee1FIO = "";
        string employee1Job = "";
        string employee2FIO = "";
        string employee2Job = "";
        string employee3FIO = "";
        string employee3Job = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Загрузка файла
            spreadsheet.LoadFromFile(examplePath);
            Worksheet worksheet = spreadsheet.Workbook.Worksheets.ByName("стр1");
            //Заполнение организации
            worksheet.Cell("A6").Value = comboBox2.Text;
            worksheet.Cell("BQ6").Value = textBox3.Text;
            //Структурное подразделение
            worksheet.Cell("A8").Value = comboBox3.Text;
            worksheet.Cell("BQ7").Value = textBox4.Text;
            //Вид деятельности
            worksheet.Cell("BQ9").Value = textBox5.Text;
            //Вид операции
            worksheet.Cell("BQ10").Value = textBox6.Text;
            //Номер документа
            worksheet.Cell("AK14").Value = textBox1.Text;
            //Дата составления
            worksheet.Cell("AS14").Value = dateTimePicker1.Value.ToString("dd.MM.yyyy");
            //Отчетный период С
            worksheet.Cell("BA14").Value = dateTimePicker2.Value.ToString("dd.MM.yyyy");
            //Отчетный период До 
            worksheet.Cell("BF14").Value = dateTimePicker3.Value.ToString("dd.MM.yyyy");
            //Мат лицо Должность
            worksheet.Cell("U18").Value = comboBox1.Text.ToString();
            //Мат Лицо ФИО
            worksheet.Cell("AM18").Value = textBox2.Text; 
            //Наименование
            int row = dataGridView1.Rows.Count - 1;
            for (int i = 0; i < row; i++)
            {
                //Номер по порядку
                worksheet.Cell(27 + i, 0).Value = i + 1;
                //Наименование
                worksheet.Cell(27 + i, 4).Value = dataGridView1[0, i].Value;
                //Код
                worksheet.Cell(27 + i, 14).Value = dataGridView1[1, i].Value;
                //Цена
                worksheet.Cell(27 + i, 19).Value = dataGridView1[2, i].Value;
            }
            //Бой,лом
            row = dataGridView2.Rows.Count - 1;
            for (int i = 0; i < row; i++)
            {
                //Количество
                worksheet.Cell(27 + i, 25).Value = dataGridView2[0, i].Value;
                //Сумма
                worksheet.Cell(27 + i, 29).Value = dataGridView2[1, i].Value;
            }
            //Утрачено
            row = dataGridView3.Rows.Count - 1;
            for (int i = 0; i < row; i++)
            {
                //Количество
                worksheet.Cell(27 + i, 34).Value = dataGridView3[0, i].Value;
                //Сумма
                worksheet.Cell(27 + i, 38).Value = dataGridView3[1, i].Value;
            }
            //Всего
            row = dataGridView4.Rows.Count - 1;
            for (int i = 0; i < row; i++)
            {
                //Количество
                worksheet.Cell(27 + i, 43).Value = dataGridView4[0, i].Value;
                //Сумма
                worksheet.Cell(27 + i, 47).Value = dataGridView4[1, i].Value;
            }
            //Примечание
            row = dataGridView5.Rows.Count - 1;
            for (int i = 0; i < row; i++)
            {
                //Количество
                worksheet.Cell(27 + i, 52).Value = dataGridView5[0, i].Value;
                //Сумма
                worksheet.Cell(27 + i, 69).Value = dataGridView5[1, i].Value;
            }
            //Расшифровка подписей
            worksheet.Cell("N65").Value = employee1Job;
            worksheet.Cell("AQ65").Value = employee1FIO;
            worksheet.Cell("N67").Value = employee2Job;
            worksheet.Cell("AQ67").Value = employee2FIO;
            worksheet.Cell("N69").Value = employee3Job;
            worksheet.Cell("AQ69").Value = employee3FIO;
            //Сохранение файла
            if (File.Exists(forCreatePath))
            {
                File.Delete(forCreatePath);
            }
            spreadsheet.SaveAs(forCreatePath);
            //MessageBox.Show(dataGridView1.RowCount.ToString());
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox3.Text = (comboBox2.SelectedIndex+1).ToString();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox4.Text = (comboBox3.SelectedIndex+1).ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
            employee1FIO = form2.employee1FIO;
            employee1Job = form2.employee1Job;
            employee2FIO = form2.employee2FIO;
            employee2Job = form2.employee2Job;
            employee3FIO = form2.employee3FIO;
            employee3Job = form2.employee3Job;

    }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            dataGridView5.Rows.Clear();


        }
    }
}
