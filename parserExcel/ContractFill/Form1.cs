using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ContractFill
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var helper = new wordHelper("Договор.docx");
            var items = new Dictionary<string, string>
            {
                {"{org}", textBox1.Text  },
                {"{dolg-rod}", textBox2.Text  },
                {"{fio-rod}", textBox3.Text  },
                {"{na-osnovanii}", textBox4.Text  },
                {"{INN}", textBox5.Text  },
                {"{KPP}", textBox6.Text  },
                {"{OGRN}", textBox7.Text  },
                {"{Adress}", textBox8.Text  },
                {"{Bank}", textBox9.Text  },
                {"{Bik}", textBox10.Text  },
                {"{DATE}", dateTimePicker1.Value.ToString("dd.MM.yyyy")  },
                {"{dolg-im}", textBox11.Text  },
                {"{fio-im}", textBox12.Text  },
                {"{r/s}", textBox13.Text  },
                {"{k/s}", textBox14.Text  },
            };

            helper.Process(items);
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }
    }
}
