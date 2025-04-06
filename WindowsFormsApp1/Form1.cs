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
    }
}
