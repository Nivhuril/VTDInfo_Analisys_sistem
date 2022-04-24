using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VTDinfo
{
    public class InsulationDefects
    {
        public string defectCoordinates;
        public double defectLength;
    }

    
    public partial class Form1 : Form
    {
        public List<InsulationDefects> insulationDefects = new List<InsulationDefects>();
        public class defectsOfInsulation
        {
            public string pipeNumber;
            public double defectLength;
            public List<int> numbersOfPipes;
        }
        
        public double ConvertDegreeAngleToDouble(string input)
        {
            double degrees = 0;
            double minutes = 0;


            if (input!=null)
            {
                input = input.Replace(".", ",");

                if (input.Length >1)//53°03.69139'
                {
                    try
                    {
                        degrees = Double.Parse(input.Substring(0, input.IndexOf("°")));
                        minutes = Double.Parse(input.Substring(input.IndexOf("°") + 1, input.Length - input.IndexOf("°") - 2));
                    }
                    catch (Exception) { }
                }
            }
            //double result = 0;
            //return minuts;

            //Decimal degrees = 
            //   whole number of degrees, 
            //   plus minutes divided by 60, 
            //   plus seconds divided by 3600

            return degrees + (minutes / 60);
        }
        public double ConvertDegreeAngleToDoubleLat(string input)
        {
            input = input.Replace(".", ","); //N54 10.531 E52 38.013

            input = input.Replace("N", "");

            double degrees = 0;
            double minutes = 0;
            if (input.Length > 15)//53°03.69139'
            {
                try
                {
                    input = input.Substring(0, input.IndexOf("E"));
                    degrees = Double.Parse(input.Substring(0, input.IndexOf(" ")));
                    minutes = Double.Parse(input.Substring(input.IndexOf(" ") + 1, input.Length - input.IndexOf(" ") - 1));
                }
                catch (Exception) { }
            }

            //double result = 0;
            //return minuts;

            //Decimal degrees = 
            //   whole number of degrees, 
            //   plus minutes divided by 60, 
            //   plus seconds divided by 3600

            return degrees + (minutes / 60);
        }
        public double ConvertDegreeAngleToDoubleLon(string input)
        {
            
            
            
            input = input.Replace(".", ","); //N54 10.531 E52 38.013
            try
            {
                input = input.Substring(input.IndexOf("E") + 1, input.Length - input.IndexOf("E") - 1);
            }
            catch (Exception)
            {

                throw;
            }
            input = input.Replace("E", "");

            double degrees = 0;
            double minutes = 0;

            try
            {
                degrees = Double.Parse(input.Substring(0, input.IndexOf(" ")));
                minutes = Double.Parse(input.Substring(input.IndexOf(" ") + 1, input.Length - input.IndexOf(" ") - 2));
            }
            catch (Exception) { }

            //double result = 0;
            //return minuts;

            //Decimal degrees = 
            //   whole number of degrees, 
            //   plus minutes divided by 60, 
            //   plus seconds divided by 3600

            return degrees + (minutes / 60);
            //return input;
        }
        public List<InsulationDefects> GetInsulationDefectsFromFile(string filename)
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = "Лист1";
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];
            richTextBox1.Invoke(new Action(() => richTextBox9.AppendText(Environment.NewLine + "Выполняется чтение журнала дефектов изоляции...")));
            richTextBox1.Invoke(new Action(() => richTextBox9.AppendText(Environment.NewLine + "->*")));
            
            int incrementor = 0;//переменная для прогресс - индикатора
            
            int i = 1;
            bool mark = true;
            while (mark)
            {
                InsulationDefects insDef = new InsulationDefects();
                insDef.defectCoordinates = Convert.ToString(ObjWorkSheet2.Cells[i, 1].Text);
                try
                {
                    insDef.defectLength = Convert.ToDouble(ObjWorkSheet2.Cells[i, 2].Text);
                }
                catch (Exception) {}
                

                if (String.IsNullOrWhiteSpace(insDef.defectCoordinates))
                {
                    mark = false;//дошли до конца трубного журлала
                }
                else
                {
                    insulationDefects.Add(insDef);
                    //richTextBox9.Invoke(new Action(() => richTextBox9.AppendText(Environment.NewLine + insDef.defectCoordinates + "_" + insDef.defectLength)));
                }

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor == 10)
                {
                    richTextBox9.Invoke(new Action(() => richTextBox1.AppendText("*")));
                    incrementor = 0;
                }
                i++;

            }
            
            richTextBox9.Invoke(new Action(() => richTextBox9.AppendText(Environment.NewLine + "Журнала прочитан, количество дефектов: " + insulationDefects.Count)));
            richTextBox9.Invoke(new Action(() => richTextBox9.AppendText(Environment.NewLine + "==========================================")));

            ObjExcel.Quit();
            return insulationDefects;
        }
       
        private List<MGPipe> ShortOperatingReadToClassPipeLogAutoFinExample(string fileName, numbersOfColumns NumbersOfColumns)//с автофинишем/КОРОТКИЙ!!!метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox42.Text;//получаем название вкладки из формы импотра (трубный журнал)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка трубного журнала...")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "->*")));
            //richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка трубного журнала...");
            //richTextBox1.AppendText(Environment.NewLine + "->*");
            //int pipeListCount = Convert.ToInt16(textBox95.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            List<MGPipe> OMGPipeS = new List<MGPipe>();//трубный журнал

            int i = 1;
            bool mark = true;
            while (mark)
            {
                MGPipe mGPipe = new MGPipe();
                mGPipe.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber1].Text);

                String txt;
                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber2].Text);
                try
                {
                    mGPipe.odometrDist = Convert.ToDouble(txt.Replace(".", ","));
                    //richTextBox1.AppendText(Environment.NewLine + "$"+ mGPipe.odometrDist);
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.odometrDist = 0;
                    richTextBox1.Invoke(new Action(() => richTextBox1.AppendText("^")));
                }


                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber3].Text);
                try
                {
                    mGPipe.thikness = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.thikness = 0;
                }

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber4].Text);
                try
                {
                    mGPipe.pipeLength = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.pipeLength = 0;
                }


                mGPipe.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber5].Text);
                mGPipe.characterFeatures = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber6].Text);
                mGPipe.clockOrientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber7].Text);


                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber8].Text);
                try
                {
                    mGPipe.bendOfPipe = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.bendOfPipe = 0;
                }*/

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber9].Text);
                try
                {
                    mGPipe.jointAngle = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.jointAngle = 0;
                }

                mGPipe.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber10].Text);
                mGPipe.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber11].Text);


                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber12].Text);
                try
                {
                    mGPipe.heightAboveSeaLevel = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.heightAboveSeaLevel = 0;
                }*/
                mGPipe.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber13].Text);
                if (String.IsNullOrEmpty(Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber14].Text)) == false)
                {
                    string localCategory = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber14].Text);
                    if (localCategory.Contains("I"))
                    {
                        mGPipe.pipelineSectionCategory = "1";
                    }
                    if (localCategory.Contains("II"))
                    {
                        mGPipe.pipelineSectionCategory = "2";
                    }
                    if (localCategory.Contains("III"))
                    {
                        mGPipe.pipelineSectionCategory = "3";
                    }
                    if (localCategory.Contains("IV"))
                    {
                        mGPipe.pipelineSectionCategory = "4";
                    }
                    else
                    {
                        mGPipe.pipelineSectionCategory = "1";
                    }
                }
                else
                {
                    mGPipe.pipelineSectionCategory = "1";
                }
                if (String.IsNullOrEmpty(Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber15].Text)) == false)
                {
                    try
                    {
                        int localCategory = Convert.ToInt32(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber15].Text);
                        mGPipe.tensileStrength = localCategory;
                    }
                    catch (Exception)
                    {
                        mGPipe.tensileStrength = 550;
                    }
                }
                else
                {
                    mGPipe.tensileStrength = 550;
                }




                if (String.IsNullOrWhiteSpace(mGPipe.pipeNumber))
                {
                    mark = false;//дошли до конца трубного журлала
                }
                else
                {
                    OMGPipeS.Add(mGPipe);
                }

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor == 100)
                {
                    richTextBox1.Invoke(new Action(() => richTextBox1.AppendText("*")));
                    incrementor = 0;
                }
                i++;
            }

            //textBox95.Text = Convert.ToString(i);//записываем в поле количество труб
            textBox95.Invoke(new Action(() => textBox95.Text = Convert.ToString(i)));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Массив данных из трубного журнала прочитан, количество труб: " + OMGPipeS.Count)));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "==========================================")));

            //richTextBox1.AppendText(Environment.NewLine + "Массив данных из трубного журнала прочитан, количество труб: "+ OMGPipeS.Count);
            //richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return OMGPipeS;
        }
    }

        
}
