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
    public partial class ResultForm : Form
    {
        /// <summary>
        /// Инициализация ResultForm с последующей обработкой данных пользователя, вычисления времени на подготовку и отрисовки рекомендаций
        /// </summary>
        /// <param name="Points"></param>
        /// <param name="Time"></param>
        /// <param name="Answers"></param>
        public ResultForm(double Points, double Time, double[][] Answers)
        {
            InitializeComponent();

            TimeForPreparation = Time;
            DesiredPoints = Points;

            Label[] text = new Label[20];//Label "Задание№n"
            TextBox[] Score = new TextBox[20];//Уровень знаний
            Label[] FinalTime = new Label[20];//Время на подготовку к теме задания

            for (int i = 0; i < 20; i++)
            {
                Score[i] = new TextBox();
                Score[i].ReadOnly = true;
                Score[i].TextAlign = HorizontalAlignment.Center;

                FinalTime[i] = new Label();
                FinalTime[i].Size = new Size(65, 23);
                FinalTime[i].Text = "0";

                text[i] = new Label();
                text[i].Text = "Задание№" + (i + 1)+':';
                text[i].Size = new Size(74, 23);
                
                if (i < 10)//Отрисовка объектов
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

                Controls.Add(text[i]);//добавление объектов на форму
                Controls.Add(Score[i]);
                Controls.Add(FinalTime[i]);
            }

            double buffer = 0;
            double points = DesiredPoints;

            for (int i=0;i<20;i++)//Алгоритм вычисления FinalTime(Время на подготовку к теме задания)
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

        public double[][] Answers = new double[20][];////Массив с данными о решении задач

        public double TimeForPreparation;//Время на подготовку

        public double DesiredPoints;//Желаемый результат

    }
}
