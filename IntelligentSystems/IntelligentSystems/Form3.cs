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
    public partial class Form3 : Form
    {
        public Form3(double Points, double Time, double[][] Answers)
        {
            InitializeComponent();

            TimeForPreparation = Time;
            DesiredPoints = Points;

            Label[] text = new Label[20];
            TextBox[] Score = new TextBox[20];
            Label[] FinalTime = new Label[20];
            /*
            //Тест
            for(int i=0;i<20;i++)
            {
                Answers[i] = new double[3];
                Answers[i][1] = 1;//баллы за задание 
                Answers[i][2] = 0;//время
            }
            Answers[0][0] = 0;
            Answers[1][0] = 1;
            Answers[2][0] = 2;
            Answers[3][0] = 2;
            Answers[4][0] = 1;
            Answers[5][0] = 2;
            Answers[6][0] = 2;
            Answers[7][0] = 1;
            Answers[8][0] = 2;
            Answers[9][0] = 0;
            Answers[10][0] = 1;
            Answers[11][0] = 2;
            Answers[12][0] = 2;
            Answers[13][0] = 1;
            Answers[14][0] = 2;
            Answers[15][0] = 2;
            Answers[16][0] = 1;
            Answers[17][0] = 2;
            Answers[18][0] = 0; //Answers[18][1] = 5;
            Answers[19][0] = 1; //Answers[19][1] = 5;
            //Тест*/

            for (int i = 0; i < 20; i++)
            {
                Score[i] = new TextBox();
                Score[i].ReadOnly = true;
                Score[i].TextAlign = HorizontalAlignment.Center;
                FinalTime[i] = new Label();

                text[i] = new Label();
                text[i].Text = "Задание№" + (i + 1)+':';
                text[i].Size = new Size(74, 23);
                FinalTime[i].Size = new Size(65, 23);
                if (i < 10)
                {
                    text[i].Location = new System.Drawing.Point(12, 33 + 30 * i);
                    FinalTime[i].Location = new System.Drawing.Point(222, 30 + 30 * i);
                    Score[i].Location = new System.Drawing.Point(92, 30 + 30 * i);
                    if(Answers[i][0]==0)
                    {
                        Score[i].Text = "Плохо";
                        Score[i].BackColor = Color.Red;
                    }
                    else if(Answers[i][0] == 1)
                    {
                        Score[i].Text = "Хорошо";
                        Score[i].BackColor = Color.Yellow;
                    }
                    else if(Answers[i][0] == 2)
                    {
                        Score[i].Text = "Отлично";
                        Score[i].BackColor = Color.LightGreen;
                    }
                }
                else
                {
                    text[i].Location = new System.Drawing.Point(283, 33 + 30 * (i - 10));
                    FinalTime[i].Location = new System.Drawing.Point(494, 30 + 30 * (i - 10));
                    Score[i].Location = new System.Drawing.Point(364, 30 + 30 * (i-10));
                    if (Answers[i][0] == 0)
                    {
                        Score[i].Text = "Плохо";
                        Score[i].BackColor = Color.Red;
                    }
                    else if (Answers[i][0] == 1)
                    {
                        Score[i].Text = "Хорошо";
                        Score[i].BackColor = Color.Yellow;
                    }
                    else if (Answers[i][0] == 2)
                    {
                        Score[i].Text = "Отлично";
                        Score[i].BackColor = Color.LightGreen;
                    }
                }
                FinalTime[i].Text = "0";
                Controls.Add(text[i]);
                Controls.Add(Score[i]);
                Controls.Add(FinalTime[i]);
            }

            double buffer = 0;
            double points = DesiredPoints;

            for (int i=0;i<20;i++)
            {
                if(Answers[i][0]==2)
                {
                    points -= Answers[i][1];
                    FinalTime[i].Text = "0";
                }
            }
            if (points > 0)
            {
               for(int i=0;i<20;i++)
               {
                    if(Answers[i][0]==1)
                    {
                        FinalTime[i].Text = Convert.ToString(TimeForPreparation * (Answers[i][1] / points));
                        buffer += Answers[i][1];
                        if(buffer>=points)
                        {
                            break;
                        }
                    }
               }
               if(buffer<points)
               {
                    for(int i=0;i<20;i++)
                    {
                        if (Answers[i][0] == 0)
                        {
                            FinalTime[i].Text = Convert.ToString(TimeForPreparation * (Answers[i][1] / points));
                        }
                    }
               }
            }
        }
        public double[][] Answers = new double[20][];
        public double TimeForPreparation;
        public double DesiredPoints;
    }
}
