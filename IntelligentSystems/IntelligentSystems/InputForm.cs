using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IntelligentSystems
{
    public partial class InputForm : Form
    {
        public InputForm()
        {
            InitializeComponent();
        }


        /// <summary>
        /// Передача входных данных в KnowledgeCheckForm
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            KnowledgeCheckForm form2 = new KnowledgeCheckForm(DesiredPoints.Text, TimeForPreparation.Text);
            this.Hide();
            form2.ShowDialog();
            //this.Show();
            Close();
        }
    }
}
