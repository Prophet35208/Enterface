using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1: Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Получаем индекс выбранной вкладки
            int selectedIndex = tabControl1.SelectedIndex;

            // В зависимости от индекса выполняем разные действия
            switch (selectedIndex)
            {
                case 0: // Выбрана первая вкладка (индекс 0)
                        // Код для первой вкладки
                    tableLayoutPanel3.Visible = true;
                    tableLayoutPanel12.Visible = false;
                    break;
                case 1: // Выбрана вторая вкладка (индекс 1)
                        // Код для второй вкладки
                    tableLayoutPanel3.Visible = false;
                    tableLayoutPanel12.Visible = true;
                    break;
                default: 
                    break;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str = comboBox1.SelectedItem.ToString(); 

            switch (str)
            {
                case "ООО \"СтройМаш\"":
                    textBoxOrgan.Text = "00001";
                    break;
                case "ООО \"ТвояМашинПочинилз\"":
                    textBoxOrgan.Text = "00002";
                    break;
                case "ООО \"КредитВКредит\"":
                    textBoxOrgan.Text = "00003";
                    break;
                default:
                    textBox1.Text = "";
                    break;
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str = comboBox2.SelectedItem.ToString();


            switch (str)
            {
                case "Бухгалтерия":
                    textBoxStruct.Text = "00001";
                    break;
                case "Кулинарный цех":
                    textBoxStruct.Text = "00002";
                    break;
                case "Линия раздачи":
                    textBoxStruct.Text = "00003";
                    break;
                default:
                    textBoxStruct.Text = "";
                    break;
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str = comboBox3.SelectedItem.ToString();

            switch (str)
            {
                case "Продукты пищевые":
                    textBoxDeyet.Text = "00001";
                    break;
                case "Напитки":
                    textBoxDeyet.Text = "00002";
                    break;
                case "Изделия табачные":
                    textBoxDeyet.Text = "00003";
                    break;
                default:
                    textBoxDeyet.Text = "";
                    break;
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str = comboBox4.SelectedItem.ToString();

            switch (str)
            {
                case "Производство пищевых продуктов":
                    textBoxOper.Text = "00001";
                    break;
                case "Производство напитков":
                    textBoxOper.Text = "00002";
                    break;
                case "Производство табачных изделий":
                    textBoxOper.Text = "00003";
                    break;
                default:
                    textBoxOper.Text = "";
                    break;
            }
        }
    }
}
