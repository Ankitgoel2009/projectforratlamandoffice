using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Datagridview
{
    public partial class textBox : TextBox
    {
        public textBox()
        {
            InitializeComponent();
        }
        int m = 1;
        protected override void OnTextChanged(EventArgs e)
        {
            m++;
            MessageBox.Show(m.ToString());
            base.OnTextChanged(e);
        }

        private void textBox_TextChanged(object sender, EventArgs e)
        { 
            m++;
            MessageBox.Show(m.ToString());
                
        }
        protected override void OnPaint(PaintEventArgs pe)
        {
           
            base.OnPaint(pe);
        }

      
    }
}
