using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FVRs_IG
{
    public partial class AddNewWord : Form
    {
        public AddNewWord()
        {
            InitializeComponent();
            buttonAddNewWord.Enabled = false;
        }

        private void AddNewWord_Load(object sender, EventArgs e)
        {
            
        }

        internal string retrieveNewWord()
        {
            return textBoxNewWord.Text;
        }

        private void buttonAddNewWord_Click(object sender, EventArgs e)
        {
            if (textBoxNewWord.Text != "")
            {
                this.Hide();
            }
        }

        private void textBoxNewWord_TextChanged(object sender, EventArgs e)
        {
            if (textBoxNewWord.Text != "" )
            {
               buttonAddNewWord.Enabled = true;
            }
            else
            {
                buttonAddNewWord.Enabled = false;
            }
        }
    }
}
