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
    public partial class custombutton : Button
    {
        public custombutton()
        {
            InitializeComponent();
        }
        //https://stackoverflow.com/questions/12816752/override-paint-a-control-in-c-sharp very good example below of this page 
        protected override void OnPaint(PaintEventArgs pe)
        {
           // RenderRainbowText("Accounting Voucher", 'V', this);
            base.OnPaint(pe);
            RenderRainbowText(this.Text, 'V', this);
        }
      
        public void RenderRainbowText(string Text, char keyword, Control Lb)
        {

            System.Drawing.Graphics formGraphics = this.CreateGraphics();
            
                string[] chunks = Text.Split(keyword);
                string word = keyword.ToString();
                // USED FOR DRAWING chunk
                SolidBrush brush = new SolidBrush(Color.Red);
                SolidBrush brush1 = new SolidBrush(Color.Black);

                 //SolidBrush[] brushes = new SolidBrush[] {  new SolidBrush(Color.Black) };
              
                float x = 0;
                for (int i = 0; i < chunks.Length; i++)
                {
                    formGraphics.DrawString(chunks[i], Lb.Font, brush1, x, 0); // brushes[i] has been replaced with brushes[0]
                    x += (formGraphics.MeasureString(chunks[i], Lb.Font)).Width;
                    //CODE TO MEASURE AND DRAW COMMA
                    if (i < (chunks.Length - 1))
                    {
                        formGraphics.DrawString(word, Lb.Font, brush, x, 0);
                        x += (formGraphics.MeasureString(",", Lb.Font)).Width;
                    }
                }
            
        }
    }
}
