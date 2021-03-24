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
    public partial class KnowledgeCheckForm : Form
    {
        /// <summary>
        /// Инициализация KnowledgeCheckForm с последующей отрисовкой первого задания теста
        /// </summary>
        /// <param name="Points"></param>
        /// <param name="Time"></param>
        public KnowledgeCheckForm(string Points, string Time)
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

        public double[][] Answers = new double[20][];//Массив с данными о решении задач
        public double TimeForPreparation;//Время на подготовку
        public double DesiredPoints;//Желаемы результат
        private int i=2, j=1, c=0;
        public string path = "../../Resources/RightAnswers.txt";//путь к текстовому файлу с решениями
        public StreamReader sr;

        /// <summary>
        /// Переключение задания, сохранение ввода пользователя в массив, передача массива и данных из InputForm в ResultForm
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            if (i == 21)
            {
                i = 1;
                j++;
                if (j == 3)//конец цикла
                {
                    ResultForm form3 = new ResultForm(DesiredPoints, TimeForPreparation, Answers);
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
                
                if (Answer.Text==sr.ReadLine())//проверка верности введеного ответа
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
