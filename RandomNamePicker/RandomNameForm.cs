using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace RandomNamePicker
{
    public partial class RandomNameForm : Form
    {
        public RandomNameForm()
        {
            InitializeComponent();
        }
        public void UpdateName(string name)
        {
            nameLabel.Text = name;
        }
    }
}
