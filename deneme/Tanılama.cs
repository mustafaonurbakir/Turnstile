using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace deneme
{
    public partial class Tanılama : Form
    {

        public string ReturnValue1 { get; set; }

        public Tanılama(string gelenveri)
        {
            InitializeComponent();
            label1.Text = gelenveri + " ismindeki sütunu ilgili sütun ile eşleştiremedim." + Environment.NewLine + "Lütfen manuel olarak eşleştiriniz:";
            label2.Text = gelenveri;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
            ReturnValue1 = comboBox1.Text;
        }
    }
}
