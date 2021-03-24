using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace IntelligentSystems
{
    public partial class Form2 : Form
    {

        public Form2(string Points, string Time)
        {
            InitializeComponent();
            UserTask.ImageLocation = "../../Resources/1_1.jpg";
            UserTask.Load();

            TimeForPreparation = double.Parse(Time);
            DesiredPoints = double.Parse(Points);

            for (int i = 0; i < 20; i++)
            {
                Answers[i] = new double[3];
                Answers[i][0] = 0;//правильные ответы
                Answers[i][1] = 1;//баллы за задание 
                Answers[i][2] = 0;//время
            }

            sr = new StreamReader(path);
        }

        public double[][] Answers = new double[20][];
        public double TimeForPreparation;
        public double DesiredPoints;
        private int i=2, j=1, c=0;
        public string path = "../../Resources/RightAnswers.txt";
        public StreamReader sr;
        private void button2_Click(object sender, EventArgs e)
        {
            Form3 form3 = new Form3(DesiredPoints,TimeForPreparation,Answers);
            this.Hide();
            form3.ShowDialog();
            //this.Show();
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (i == 21)
            {
                i = 1;
                j++;
                if (j == 3)
                {
                    Form3 form3 = new Form3(DesiredPoints, TimeForPreparation, Answers);
                    this.Hide();
                    form3.ShowDialog();
                    Close();
                }
            }
            if (j != 3)
            {
                string Name = "../../Resources/" + j + "_" + i + ".jpg";
                UserTask.ImageLocation = Name;
                UserTask.Load();
                
                if (Answer.Text==sr.ReadLine())
                {
                    Answers[c][0]++;
                }
                Answer.Text = String.Empty;
                c++;
                if(c==20)
                {
                    c=0;
                }
                
                i++;
            }
        }
    }
}
