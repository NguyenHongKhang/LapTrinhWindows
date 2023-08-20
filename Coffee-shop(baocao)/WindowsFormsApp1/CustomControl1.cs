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
    public partial class CustomControl1 : Control
    {
        private Button _button;
        private Label _label;

        public CustomControl1(Button button, Label label)
        {
            _button = button;
            _label = label;
            label.AutoSize = true;
            button.BackColor = Color.White;
            Height = button.Height + label.Height;
            Width = Math.Max(button.Width, label.Width);
            Controls.Add(_button);
            _button.Location = new Point(0, 0);
            Controls.Add(_label);
            _label.Location = new Point(0, button.Height);
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
        }
    }
}

