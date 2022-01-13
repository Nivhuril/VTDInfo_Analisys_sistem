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

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
        }
        class numbersOfColumns//класс для хранения согласованных номеров столбцов и строк для импорта данных из файла EXCEL
        {
            public int stringNumber1;
            public int stringNumber2;
            public int stringNumber3;
            public int stringNumber4;
            public int stringNumber5;
            public int stringNumber6;
            public int stringNumber7;
            public int stringNumber8;

            public int columnNumber1;
            public int columnNumber2;
            public int columnNumber3;
            public int columnNumber4;
            public int columnNumber5;
            public int columnNumber6;
            public int columnNumber7;
            public int columnNumber8;
            public int columnNumber9;
            public int columnNumber10;
            public int columnNumber11;
            public int columnNumber12;
            public int columnNumber13;

            public int column2Number1;
            public int column2Number2;
            public int column2Number3;
            public int column2Number4;
            public int column2Number5;
            public int column2Number6;
            public int column2Number7;
            public int column2Number8;
            public int column2Number9;
            public int column2Number10;
            public int column2Number11;
            public int column2Number12;
            public int column2Number13;
            public int column2Number14;
            public int column2Number15;
            public int column2Number16;
            public int column2Number17;
            public int column2Number18;


            public int column3Number1;
            public int column3Number2;
            public int column3Number3;
            public int column3Number4;
            public int column3Number5;
            public int column3Number6;
            public int column3Number7;
            public int column3Number8;
            public int column3Number9;
            public int column3Number10;
            public int column3Number11;
            public int column3Number12;
            public int column3Number13;
        }

        public class MGPipe//класс для хранения данных одной трубы
        {
            public string pipeNumber;//номер трубы
            public string odometrDist;//дистанция по одометру
            public string thikness;//толщина трубы
            public string pipeLength;//длина трубы
            public string distanceFromReferencePoints;//расстояние от реперных точек
            public string characterFeatures;// характер особенности
            public string clockOrientation;//Ориент., ч:мин
            public string bendOfPipe;//Изгиб, °
            public string jointAngle;//Угол стыка,°
            public string Latitude;//Широта
            public string Longitude;//Долгота
            public string heightAboveSeaLevel;//H, м
            public string Note;//Примечание

        }

        public class MGVTD//класс для хранения данных одного отчета ВТД
        {
            public string MGName;
            public string MGDate;
            public string MGPressure;
            public List<MGPipe> MGPipeS = new List<MGPipe>();
        }

        public MGVTD mGVTD = new MGVTD();//создаём экзампляр класса для хранения данных обследования ВТД
        numbersOfColumns NumbersOfColumns = new numbersOfColumns();
        private void readFile1(string fileName)
        { }
        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 newForm = new Form2();
            newForm.Show();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        string fileName;
        private void button4_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName=openFileDialog1.FileName;
                button3.Enabled = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();                                                                                                                                                                      
                //Открываем книгу.                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                string WorksheetName = textBox9.Text;//получаем название вкладки из формы импотра
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

                //Очищаем от старого текста окно вывода.
                //richTextBox1.Clear();

            //номера строк для  листа "информация о трубопроводе"
                int stringNumber1 = Convert.ToInt16(textBox10.Text);//получаем номер строки с названием трубопровода
                int stringNumber2 = Convert.ToInt16(textBox11.Text);//получаем номер строки с названием участка
                int stringNumber3 = Convert.ToInt16(textBox12.Text);//получаем номер строки со значением диаметра
                int stringNumber4 = Convert.ToInt16(textBox13.Text);//получаем номер строки с именем принципала
                int stringNumber5 = Convert.ToInt16(textBox14.Text);//получаем номер строки с датой обследования
                int stringNumber6 = Convert.ToInt16(textBox15.Text);//получаем номер строки с проектным давлением
                int stringNumber7 = Convert.ToInt16(textBox16.Text);//получаем номер строки с рабочим давлением
                int stringNumber8 = Convert.ToInt16(textBox17.Text);//получаем номер строки с датой ввода в эксплуатацию
            //numbersOfColumns





                textBox1.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber1, 4].Text);
                textBox2.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber2, 4].Text);
                textBox3.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber3, 4].Text);
                textBox4.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber4, 4].Text);
                textBox5.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber5, 4].Text);
                textBox6.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber6, 4].Text);
                textBox7.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber7, 4].Text);
                textBox8.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber8, 4].Text);


                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
                string WorksheetName2 = textBox42.Text;//получаем название вкладки из формы импотра
                ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];

            //номера столбцов для "трубного журлала"
               int columnNumber1 = Convert.ToInt16(textBox30.Text);
               int columnNumber2 = Convert.ToInt16(textBox31.Text);
               int columnNumber3 = Convert.ToInt16(textBox32.Text);
               int columnNumber4 = Convert.ToInt16(textBox33.Text);
               int columnNumber5 = Convert.ToInt16(textBox34.Text);
               int columnNumber6 = Convert.ToInt16(textBox35.Text);
               int columnNumber7 = Convert.ToInt16(textBox36.Text);
               int columnNumber8 = Convert.ToInt16(textBox37.Text);
               int columnNumber9 = Convert.ToInt16(textBox38.Text);
               int columnNumber10 = Convert.ToInt16(textBox39.Text);
               int columnNumber11 = Convert.ToInt16(textBox40.Text);
               int columnNumber12 = Convert.ToInt16(textBox41.Text);
               int columnNumber13 = Convert.ToInt16(textBox44.Text);


                textBox18.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber1].Text);
                textBox19.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber2].Text);
                textBox20.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber3].Text);
                textBox21.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber4].Text);
                textBox22.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber5].Text);
                textBox23.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber6].Text);
                textBox24.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber7].Text);
                textBox25.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber8].Text);
                textBox26.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber9].Text);
                textBox27.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber10].Text);
                textBox28.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber11].Text);
                textBox29.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber12].Text);
                textBox43.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber13].Text);


                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet3;
                string WorksheetName3 = textBox45.Text;//получаем название вкладки из формы импотра
                ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];
            //номера столбцов для "журлала аномалий"
                int column2Number1 = Convert.ToInt16(textBox67.Text);
                int column2Number2 = Convert.ToInt16(textBox68.Text);
                int column2Number3 = Convert.ToInt16(textBox69.Text);
                int column2Number4 = Convert.ToInt16(textBox70.Text);
                int column2Number5 = Convert.ToInt16(textBox71.Text);
                int column2Number6 = Convert.ToInt16(textBox72.Text);
                int column2Number7 = Convert.ToInt16(textBox73.Text);
                int column2Number8 = Convert.ToInt16(textBox74.Text);
                int column2Number9 = Convert.ToInt16(textBox75.Text);
                int column2Number10 = Convert.ToInt16(textBox76.Text);
                int column2Number11 = Convert.ToInt16(textBox77.Text);
                int column2Number12 = Convert.ToInt16(textBox78.Text);
                int column2Number13 = Convert.ToInt16(textBox79.Text);
                int column2Number14 = Convert.ToInt16(textBox80.Text);
                int column2Number15 = Convert.ToInt16(textBox81.Text);
                int column2Number16 = Convert.ToInt16(textBox82.Text);
                int column2Number17 = Convert.ToInt16(textBox83.Text);
                int column2Number18 = Convert.ToInt16(textBox84.Text);

                textBox46.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number1].Text);
                textBox47.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number2].Text);
                textBox48.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number3].Text);
                textBox49.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number4].Text);
                textBox50.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number5].Text);
                textBox51.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number6].Text);
                textBox52.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number7].Text);
                textBox53.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number8].Text);
                textBox54.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number9].Text);
                textBox55.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number10].Text);
                textBox56.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number11].Text);
                textBox57.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number12].Text);
                textBox58.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number13].Text);
                textBox59.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number14].Text);
                textBox60.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number15].Text);
                textBox61.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number16].Text);
                textBox62.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number17].Text);
                textBox63.Text = Convert.ToString(ObjWorkSheet3.Cells[2, column2Number18].Text);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet4;
            string WorksheetName4 = textBox108.Text;//получаем название вкладки из формы импотра
            ObjWorkSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName4];


            //номера столбцов для "журнала элементов обустройства"
            int column3Number1 = Convert.ToInt16(textBox96.Text);
            int column3Number2 = Convert.ToInt16(textBox97.Text);
            int column3Number3 = Convert.ToInt16(textBox98.Text);
            int column3Number4 = Convert.ToInt16(textBox99.Text);
            int column3Number5 = Convert.ToInt16(textBox100.Text);
            int column3Number6 = Convert.ToInt16(textBox101.Text);
            int column3Number7 = Convert.ToInt16(textBox102.Text);
            int column3Number8 = Convert.ToInt16(textBox103.Text);
            int column3Number9 = Convert.ToInt16(textBox104.Text);
            int column3Number10 = Convert.ToInt16(textBox105.Text);
            int column3Number11 = Convert.ToInt16(textBox106.Text);
            int column3Number12 = Convert.ToInt16(textBox107.Text);
            int column3Number13 = Convert.ToInt16(textBox109.Text);

            textBox64.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number1].Text);
            textBox65.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number2].Text);
            textBox66.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number3].Text);
            textBox85.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number4].Text);
            textBox86.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number5].Text);
            textBox87.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number6].Text);
            textBox88.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number7].Text);
            textBox89.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number8].Text);
            textBox90.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number9].Text);
            textBox91.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number10].Text);
            textBox92.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number11].Text);
            textBox93.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number12].Text);
            textBox94.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number13].Text);




            /*MGPipe mGPipe = new MGPipe();
            mGPipe.pipeName = Convert.ToString(ObjWorkSheet.Cells[i + 1, 1].Text);
            mGPipe.pipeLong = Convert.ToString(ObjWorkSheet.Cells[i + 1, 2].Text);
            mGVTD.MGPipeS.Add(mGPipe);*/
            //номера строк для  листа "информация о трубопроводе"
            

            NumbersOfColumns.stringNumber1 = Convert.ToInt16(textBox10.Text);
            NumbersOfColumns.stringNumber2 = Convert.ToInt16(textBox11.Text);//получаем номер строки с названием участка
            NumbersOfColumns.stringNumber3 = Convert.ToInt16(textBox12.Text);//получаем номер строки со значением диаметра
            NumbersOfColumns.stringNumber4 = Convert.ToInt16(textBox13.Text);//получаем номер строки с именем принципала
            NumbersOfColumns.stringNumber5 = Convert.ToInt16(textBox14.Text);//получаем номер строки с датой обследования
            NumbersOfColumns.stringNumber6 = Convert.ToInt16(textBox15.Text);//получаем номер строки с проектным давлением
            NumbersOfColumns.stringNumber7 = Convert.ToInt16(textBox16.Text);//получаем номер строки с рабочим давлением
            NumbersOfColumns.stringNumber8 = Convert.ToInt16(textBox17.Text);//получаем номер строки с датой ввода в эксплуатацию
                                                                             //numbersOfColumns

            //номера столбцов для "трубного журлала"
            NumbersOfColumns.columnNumber1 = Convert.ToInt16(textBox30.Text);
            NumbersOfColumns.columnNumber2 = Convert.ToInt16(textBox31.Text);
            NumbersOfColumns.columnNumber3 = Convert.ToInt16(textBox32.Text);
            NumbersOfColumns.columnNumber4 = Convert.ToInt16(textBox33.Text);
            NumbersOfColumns.columnNumber5 = Convert.ToInt16(textBox34.Text);
            NumbersOfColumns.columnNumber6 = Convert.ToInt16(textBox35.Text);
            NumbersOfColumns.columnNumber7 = Convert.ToInt16(textBox36.Text);
            NumbersOfColumns.columnNumber8 = Convert.ToInt16(textBox37.Text);
            NumbersOfColumns.columnNumber9 = Convert.ToInt16(textBox38.Text);
            NumbersOfColumns.columnNumber10 = Convert.ToInt16(textBox39.Text);
            NumbersOfColumns.columnNumber11 = Convert.ToInt16(textBox40.Text);
            NumbersOfColumns.columnNumber12 = Convert.ToInt16(textBox41.Text);
            NumbersOfColumns.columnNumber13 = Convert.ToInt16(textBox44.Text);

            //номера столбцов для "журлала аномалий"
            NumbersOfColumns.column2Number1 = Convert.ToInt16(textBox67.Text);
            NumbersOfColumns.column2Number2 = Convert.ToInt16(textBox68.Text);
            NumbersOfColumns.column2Number3 = Convert.ToInt16(textBox69.Text);
            NumbersOfColumns.column2Number4 = Convert.ToInt16(textBox70.Text);
            NumbersOfColumns.column2Number5 = Convert.ToInt16(textBox71.Text);
            NumbersOfColumns.column2Number6 = Convert.ToInt16(textBox72.Text);
            NumbersOfColumns.column2Number7 = Convert.ToInt16(textBox73.Text);
            NumbersOfColumns.column2Number8 = Convert.ToInt16(textBox74.Text);
            NumbersOfColumns.column2Number9 = Convert.ToInt16(textBox75.Text);
            NumbersOfColumns.column2Number10 = Convert.ToInt16(textBox76.Text);
            NumbersOfColumns.column2Number11 = Convert.ToInt16(textBox77.Text);
            NumbersOfColumns.column2Number12 = Convert.ToInt16(textBox78.Text);
            NumbersOfColumns.column2Number13 = Convert.ToInt16(textBox79.Text);
            NumbersOfColumns.column2Number14 = Convert.ToInt16(textBox80.Text);
            NumbersOfColumns.column2Number15 = Convert.ToInt16(textBox81.Text);
            NumbersOfColumns.column2Number16 = Convert.ToInt16(textBox82.Text);
            NumbersOfColumns.column2Number17 = Convert.ToInt16(textBox83.Text);
            NumbersOfColumns.column2Number18 = Convert.ToInt16(textBox84.Text);

            //номера столбцов для "журнала элементов обустройства"
            NumbersOfColumns.column3Number1 = Convert.ToInt16(textBox96.Text);
            NumbersOfColumns.column3Number2 = Convert.ToInt16(textBox97.Text);
            NumbersOfColumns.column3Number3 = Convert.ToInt16(textBox98.Text);
            NumbersOfColumns.column3Number4 = Convert.ToInt16(textBox99.Text);
            NumbersOfColumns.column3Number5 = Convert.ToInt16(textBox100.Text);
            NumbersOfColumns.column3Number6 = Convert.ToInt16(textBox101.Text);
            NumbersOfColumns.column3Number7 = Convert.ToInt16(textBox102.Text);
            NumbersOfColumns.column3Number8 = Convert.ToInt16(textBox103.Text);
            NumbersOfColumns.column3Number9 = Convert.ToInt16(textBox104.Text);
            NumbersOfColumns.column3Number10 = Convert.ToInt16(textBox105.Text);
            NumbersOfColumns.column3Number11 = Convert.ToInt16(textBox106.Text);
            NumbersOfColumns.column3Number12 = Convert.ToInt16(textBox107.Text);
            NumbersOfColumns.column3Number13 = Convert.ToInt16(textBox109.Text);

            ObjExcel.Quit();
            
        }

        private void label32_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            
            
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            string WorksheetName = textBox9.Text;//получаем название вкладки из формы импотра (данные о газопроводе)
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox42.Text;//получаем название вкладки из формы импотра
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet3;
            string WorksheetName3 = textBox45.Text;//получаем название вкладки из формы импотра
            ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet4;
            string WorksheetName4 = textBox108.Text;//получаем название вкладки из формы импотра
            ObjWorkSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName4];



            
            for (int i = 1; i < 319; i++)
            {
                MGPipe mGPipe = new MGPipe();
                mGPipe.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber1].Text);
                mGPipe.odometrDist = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber2].Text);
                mGPipe.thikness = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber3].Text);
                mGPipe.pipeLength = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber4].Text);
                mGPipe.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber5].Text);
                mGPipe.characterFeatures = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber6].Text);
                mGPipe.clockOrientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber7].Text);
                mGPipe.bendOfPipe = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber8].Text);
                mGPipe.jointAngle = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber9].Text);
                mGPipe.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber10].Text);
                mGPipe.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber11].Text);
                mGPipe.heightAboveSeaLevel = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber12].Text);
                mGPipe.Note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber13].Text);
                mGVTD.MGPipeS.Add(mGPipe);
            }
            for (int j = 0; j < 318; j++)
            {
                richTextBox1.AppendText(Environment.NewLine + mGVTD.MGPipeS[j].distanceFromReferencePoints);
            }

        }
    }
}
