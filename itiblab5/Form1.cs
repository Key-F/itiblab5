using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace itiblab5
{
    public partial class Form1 : Form
    {
        List<theatre> T1 = new List<theatre>();
        int count = 0;
        public Form1()
        {
            InitializeComponent();
            
        }
      
             
        private void button1_Click(object sender, EventArgs e) // Получение данных
        {
            
            // openFileDialog1.ShowDialog(); // без этого мб ошибка
            //Открываем файл Экселя
            //OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                
                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                richTextBox1.Clear(); //Очищаем от старого текста окно вывода.
                int i = 2;
                int it = 0;
                char[] separators = { ' ', ';', ',' };
                String[] substrings = textBox2.Text.Split(separators, StringSplitOptions.RemoveEmptyEntries);

                do
                {
                    //Выбираем область таблицы. (в нашем случае просто ячейку)
                    try
                    {
                        Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range(textBox1.Text + i.ToString(), textBox1.Text + i.ToString());
                        Microsoft.Office.Interop.Excel.Range range1 = ObjWorkSheet.get_Range(substrings[0] + i.ToString(), substrings[0] + i.ToString());
                        Microsoft.Office.Interop.Excel.Range range2 = ObjWorkSheet.get_Range(substrings[1] + i.ToString(), substrings[1] + i.ToString());
                        if (range.Text.ToString() == "")
                        {
                            //MessageBox.Show("Строки кончились", "Работа закончена", MessageBoxButtons.OK); // Из-за этого пропадает окно
                            ObjExcel.Quit();  //Удаляем приложение (выходим из экселя) - а то будет висеть в процессах!     
                            richTextBox1.ScrollToCaret();
                            //done();
                            return;
                        }
                        else
                        {
                            string text1 = range1.Text, text2 = range2.Text;
                            if (range1.Text == "")
                                text1 = "0";
                            if (range2.Text == "")
                                text2 = "0";

                            theatre t = new theatre(range.Text, Convert.ToInt32(text1) + Convert.ToInt32(text2));
                            int summamest = Convert.ToInt32(text1) + Convert.ToInt32(text2);
                            string summamestString = Convert.ToString(summamest);
                            if (summamest == 0)
                                summamestString = "Нет данных.";
                            T1.Add(t);
                            //Добавляем полученный из ячейки текст.
                            richTextBox1.Text = richTextBox1.Text + range.Text.ToString() + ". Общее количество мест: " + summamestString + "\n";
                            richTextBox1.Text = richTextBox1.Text + "\n";
                            //richTextBox1.ScrollToCaret();
                            //это чтобы форма прорисовывалась (не подвисала)...
                            System.Windows.Forms.Application.DoEvents();
                            i++;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Некорректное значение столбца");
                        ObjExcel.Quit();  //Удаляем приложение (выходим из экселя) - а то будет висеть в процессах!
                        return;
                    } 
                }
                    while (it != 1);                                          
            }
        }





        private void button2_Click(object sender, EventArgs e) // Работа с НС
        {
            int kolepoh = 0;
            char[] separators = { ' ', ';', ',' };
            String[] substrings = textBox3.Text.Split(separators, StringSplitOptions.RemoveEmptyEntries);            
            int kolkl = substrings.Length; // Количество нейронов совпадает с количеством кластеров, которое должна выделить сеть. 
            
            int[] klastersreal = new int[kolkl];
            for (int i = 0; i < kolkl; i++)
            {
                klastersreal[i] = Convert.ToInt32(substrings[i]);
            }

            var random = new Random();
            if (T1.Count < 1) MessageBox.Show("Список не загружен");
            float[,] w = new float[kolkl, T1.Count];

           /* for (int j = 0; j < kolkl; j++)
                for (int i = 0; i < T1.Count; i++)
                {
                    w[j, i] = rasst(klastersreal[j], T1[i].capacity);
                    //w[j, i] = random.Next(1, 10); 
                } */
            
            for (int j = 0; j < T1.Count; j++)
                for (int i = 0; i < kolkl; i++)
                {
                    w[i, j] = (float)random.Next(10 * (-1), 10 * 1) / 10f; // iter * 10f для точности
                }
        m1:    count++;
            for (int i = 0; i < T1.Count; i++)
            {
                
                for (int j = 0; j < kolkl; j++)
                {
                    float[] klasters = new float[kolkl];
                    klasters[j] = summator(w, T1[i], j, T1.Count);
                    funckonkur(klasters);
                    for (int hj = 0; hj < kolkl; hj++)
                    {
                        if (klasters[hj] == 1)
                        
                            T1[i].nomerklastera = hj;                                                   
                    }
                }

            }
       kolepoh++;
            for (int i = 0; i < T1.Count; i++) // пересчет весов
            {
                for (int hj = 0; hj < kolkl; hj++)
                {
                    //if (klasters[hj] == 1)
                        w[hj, i] = w[hj, i] + Math.Abs(w[hj, i] - T1[i].capacity);
                }
            }
            if (count < 200) goto m1;
            //if (Math.Abs(klastersreal[0] - T1[1].capacity) > (Math.Abs(klastersreal[1] - T1[1].capacity))) goto m1;
            //printT(T1);
            for (int i = 0; i < T1.Count; i++)
            {
                richTextBox1.Text = richTextBox1.Text + T1[i].name +"; Количество мест: " + T1[i].capacity
                    + "; Номер кластера: " + T1[i].nomerklastera + "\n";
                richTextBox1.Text = richTextBox1.Text + "\n";
            }
        }
        static int rasst(int mesta_centra_klastera, int mesta_teatra) // Евклидово расстояние для мест в театре
        {
            //return Math.Sqrt(Math.Pow(Math.Abs(mesta_centra_klastera - mesta_teatra), 2));
            return Math.Abs(mesta_centra_klastera - mesta_teatra);
        }

        static void funckonkur(float[] summator)
        {
            float max = summator[0];
            int maxi = 0;
            for (int i = 1; i < summator.Length; i++)
            {
                if (summator[i] > max)
                {
                    max = summator[i];
                    maxi = i;
                }
            }
            for (int i = 0; i < summator.Length; i++)
            {
                if (summator[i] == max)
                    summator[i] = 1;
                else summator[i] = 0;
            }
            //return maxi;
        }



        static float summator(float[,] w, theatre T, int j, int m) // Ленейный взвешенный сумматор, j - номер нейрона, m - количество входных элементов 
        {
            float s = 0;
            for (int i = 0; i < m; i++)
            {
                s += w[j, i] * T.capacity;
            }
            return s;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            char[] separators = { ' ', ';', ',' };
            String[] substrings = textBox3.Text.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            int kolkl = substrings.Length; // Количество нейронов совпадает с количеством кластеров, которое должна выделить сеть. 

            float[] klastersreal = new float[kolkl];
            for (int i = 0; i < kolkl; i++)
            {
                klastersreal[i] = Convert.ToInt32(substrings[i]);
            }

            var random = new Random();
            if (T1.Count < 1) MessageBox.Show("Список не загружен");
            float[,] w = new float[kolkl, T1.Count];
            int numklastera;

            for (int i = 0; i < T1.Count; i++)
            {

                for (int j = 0; j < kolkl - 1; j++)
                {
                    //float[] klasters = new float[kolkl];
                    float min = Math.Abs(klastersreal[j] - T1[i].capacity);
                    if (min < Math.Abs(klastersreal[j + 1] - T1[i].capacity))
                        numklastera = j;
                    else numklastera = j + 1;
                   

                            T1[i].nomerklastera = numklastera;
                    
                }

            }
            for (int i = 0; i < T1.Count; i++)
            {
                richTextBox1.Text = richTextBox1.Text + T1[i].name + "; Количество мест: " + T1[i].capacity
                    + "; Номер кластера: " + T1[i].nomerklastera + "\n";
                richTextBox1.Text = richTextBox1.Text + "\n";
            }
        }
       
    }
}
