using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;


namespace VTDinfo
{
    class ListOfTees
    {
        public int pipelineID;//ID трубопровода (заполняется после чтения таблицы)
        public string LPUMG_name;//Наименование ЛПУ (столбец 1)
        public string teeName;//№ тройника по схеме (столбец 2)
        public string technicalCondition;//Состояние (столбец 3)
        public string installationDate;//дата установки (столбец 4)
        public string dateOfDiagnosis;//Дата диагностики (столбец 5)
        public string repareDate;//Дата ремонта (столбец 6)
        public string affiliation;//Принадлежность (столбец 7)
        public string pipelineName;//трубопровод (название) (столбец 8)
        public string pipelineSection;//участок трубопровода (столбец 9)
        public string installationLocation;//Место установки (столбец 10)
        public double lineDiameter;//Диаметр магистрали, мм (столбец 11)
        public double branchDiameter;//Диаметр отвода, мм (столбец 12)
        public string typeOfTee;//тип пройника (столбец 13)
        public string isUnderground;//Подземный или надземный (столбец 14)
        public string note;//Примечание
        //
        public bool isSorted;//истина, если при сортировке данный тройник был отсортирован в какой- то из учетных участков
    }

    public partial class Form1
    {
        private List<ListOfTees> GetListOfNumbersTees (string fileName)
        {
            List<ListOfTees> listOfTees = new List<ListOfTees>();
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = "Ключ для тройников";//получаем название вкладки из формы импотра (трубный журнал) "SonarFormat"
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];
            richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + "Выполняется чтение таблицы принадлежности тройников...")));
            richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + "->*")));
            bool mark = true;
            int inkrementor = 1;
            int i = 1;
            while (mark)
            {
                ListOfTees str = new ListOfTees();

                //str.pipelineID = Convert.ToInt32(ObjWorkSheet2.Cells[i, 1].Text);//ID трубопровода (заполняется после чтения таблицы)
                str.teeName = Convert.ToString(ObjWorkSheet2.Cells[i, 1].Text);//№ тройника по схеме (столбец 2)
                if (String.IsNullOrWhiteSpace(str.teeName) == false)
                {
                    try
                    {
                        str.pipelineID = Convert.ToInt32(ObjWorkSheet2.Cells[i, 2].Text.Trim());
                    }
                    catch (Exception)
                    {
                        str.pipelineID = 999;
                    }

                    listOfTees.Add(str);
                    //richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + listOfTees[listOfTees.Count - 1].teeName)));
                    if (inkrementor > 100)
                    {
                        inkrementor = 0;
                        richTextBox7.Invoke(new Action(() => richTextBox7.AppendText("*")));
                    }
                }
                else
                {
                    mark = false;
                }
                i++;
                inkrementor++;
            }
            richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + "Количество строк журнала: " + listOfTees.Count)));
            return listOfTees;
        }
        private List<ListOfTees> GetListOfTees(string fileName)
        {
            List<ListOfTees> listOfTees = new List<ListOfTees>();
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = "Лист1";//получаем название вкладки из формы импотра (трубный журнал) "SonarFormat"
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];
            richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + "Выполняется чтение данных о тройниках...")));
            richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + "->*")));
            bool mark = true;
            int inkrementor = 1;
            int i = 5;

   
            while (mark)
            {
                ListOfTees str = new ListOfTees();
                string numb = Convert.ToString(ObjWorkSheet2.Cells[i, 1].Text);
                
                //str.pipelineID = Convert.ToInt32(ObjWorkSheet2.Cells[i, 1].Text);//ID трубопровода (заполняется после чтения таблицы)
                str.LPUMG_name = Convert.ToString(ObjWorkSheet2.Cells[i, 1].Text);//Наименование ЛПУ (столбец 1)
                if (String.IsNullOrWhiteSpace(str.LPUMG_name) ==false)
                {
                    str.teeName = Convert.ToString(ObjWorkSheet2.Cells[i, 2].Text);//№ тройника по схеме (столбец 2)
                    str.technicalCondition = Convert.ToString(ObjWorkSheet2.Cells[i, 3].Text);//Состояние (столбец 3)
                    str.installationDate = Convert.ToString(ObjWorkSheet2.Cells[i, 4].Text);//дата установки (столбец 4)
                    str.dateOfDiagnosis = Convert.ToString(ObjWorkSheet2.Cells[i, 5].Text);//Дата диагностики (столбец 5)
                    str.repareDate = Convert.ToString(ObjWorkSheet2.Cells[i, 6].Text);//Дата ремонта (столбец 6)
                    str.affiliation = Convert.ToString(ObjWorkSheet2.Cells[i, 7].Text);//Принадлежность (столбец 7)
                    str.pipelineName = Convert.ToString(ObjWorkSheet2.Cells[i, 8].Text);//трубопровод (название) (столбец 8)
                    str.pipelineSection = Convert.ToString(ObjWorkSheet2.Cells[i, 9].Text);//участок трубопровода (столбец 9)
                    str.installationLocation = Convert.ToString(ObjWorkSheet2.Cells[i, 10].Text);//Место установки (столбец 10)
                    str.lineDiameter = Convert.ToDouble(ObjWorkSheet2.Cells[i, 11].Text.Trim().Replace(" мм", "").Replace("мм", ""));//Диаметр магистрали, мм (столбец 11)
                    try
                    {
                        str.branchDiameter = Convert.ToDouble(ObjWorkSheet2.Cells[i, 12].Text.Trim().Replace(" мм", "").Replace("мм", ""));//Диаметр отвода, мм (столбец 12)
                    }
                    catch (Exception)
                    {
                        str.branchDiameter = str.lineDiameter;//если не указан диаметр отвода, значит он равен диаметру магистрали
                    }
                    
                    str.typeOfTee = Convert.ToString(ObjWorkSheet2.Cells[i, 13].Text);//тип пройника (столбец 13)
                    str.isUnderground = Convert.ToString(ObjWorkSheet2.Cells[i, 14].Text);//Истина, если подземный (столбец 14)
                    str.note = Convert.ToString(ObjWorkSheet2.Cells[i, 15].Text);//Примечание
                    str.isSorted = false;
                    listOfTees.Add(str);
                    //richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + listOfTees[listOfTees.Count - 1].teeName)));
                    if (inkrementor > 100)
                    {
                        inkrementor = 0;
                        richTextBox7.Invoke(new Action(() => richTextBox7.AppendText("*")));
                    }
                }
                else
                {
                    mark = false;
                }

                i++;
                inkrementor++;
            }
            richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + "Количество строк журнала: " + listOfTees.Count)));
            return listOfTees;
        }
        private List<ListOfTees> SetIDtoTees (List<ListOfTees> listOfTees, List<ListOfTees> InfoOfTees)//расставим айдишники участков на тройниках
        {
            int numb = 0;
            
            for (int i = 0; i < listOfTees.Count; i++)
            {
                for (int j = 0; j < InfoOfTees.Count; j++)
                {
                    if (String.Equals(listOfTees[i].teeName, InfoOfTees[j].teeName))
                    {
                        listOfTees[i].pipelineID = InfoOfTees[j].pipelineID;
                        numb++;
                    }
                }
            }
            richTextBox7.AppendText(Convert.ToString(numb));
            return listOfTees;
        }
        private List<ListOfTees> GetListForCurrentObject(List<ListOfTees> listOfTees, int ID)//создаём список тройников, относящихся к заданному участку.
        {
            int numb = 0;
            List<ListOfTees> result = new List<ListOfTees>();
            for (int i = 0; i < listOfTees.Count; i++)
            {
                numb++;
                if (ID == listOfTees[i].pipelineID)
                {
                    result.Add(listOfTees[i]);
                }
            }
            //richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + result.Count)));
            return result;
        }
        private List<ListOfTees> getInfoAbouteTees(List<ListOfTees> listOfTees, List<pipeSectionLog> pipeSectionS)
        {
            
            for (int i = 0; i < pipeSectionS.Count; i++)
            {
                    int numberOfTees = 0;
                    int numberOfDiagnosisTees = 0;
                    int numberOfBrookenTees = 0;
                    int numberOfReparedTees = 0;

                    string Tees = "";
                    string listOfDiagnosisTees = "";
                    string listOfBrookenTees = "";
                    string listOfReparedTees = "";

                    List<ListOfTees> currentObjectTees = GetListForCurrentObject(listOfTees, pipeSectionS[i].pipelineID);
                    numberOfTees = currentObjectTees.Count;

                for (int f = 0; f < currentObjectTees.Count; f++)//пометим в исходной таблице тройников учтенные тройники как учтенные.
                {
                    for (int k = 0; k < listOfTees.Count; k++)
                    {
                        if (currentObjectTees[f].pipelineID == listOfTees[k].pipelineID)
                        {
                            listOfTees[k].isSorted = true;
                        }
                    }
                }

                if (currentObjectTees.Count>0)
                {
                    for (int j = 0; j < currentObjectTees.Count; j++)
                    {
                        Tees = String.Concat(Tees, ", ", currentObjectTees[j].teeName);

                        if (String.IsNullOrWhiteSpace(currentObjectTees[j].dateOfDiagnosis) == false)
                        {
                            numberOfDiagnosisTees++;
                            listOfDiagnosisTees = String.Concat(listOfDiagnosisTees, ", ", currentObjectTees[j].teeName);
                        }
                        if (String.IsNullOrWhiteSpace(currentObjectTees[j].repareDate) == false)
                        {
                            numberOfReparedTees++;
                            listOfReparedTees = String.Concat(listOfReparedTees, ", ", currentObjectTees[j].teeName);
                        }
                        if (currentObjectTees[j].technicalCondition.Contains("требуется ремонт"))
                        {
                            numberOfBrookenTees++;
                            listOfBrookenTees = String.Concat(listOfBrookenTees, ", ", currentObjectTees[j].teeName);
                        }
                    }

                    if (Tees.Length>2)
                    {
                        Tees = Tees.Remove(0, 2);
                    }
                    if (listOfDiagnosisTees.Length>2)
                    {
                        listOfDiagnosisTees = listOfDiagnosisTees.Remove(0, 2);
                    }
                    if (listOfBrookenTees.Length>2)
                    {
                        listOfBrookenTees = listOfBrookenTees.Remove(0, 2);
                    }
                    if (listOfReparedTees.Length>2)
                    {
                        listOfReparedTees = listOfReparedTees.Remove(0, 2);
                    }
                    
                    richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + pipeSectionS[i].pipelineID + ";" + numberOfTees + ";" + Tees + ";" +
                            numberOfDiagnosisTees + ";" + listOfDiagnosisTees + ";" + numberOfBrookenTees + ";" + listOfBrookenTees + ";" +
                            numberOfReparedTees + ";" + listOfReparedTees)));
                }
                else
                {
                    richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + pipeSectionS[i].pipelineID + ";" + 0 + ";" + "" + ";" +
                        0 + ";" + "" + ";" + 0 + ";" + "" + ";" + 0 + ";" + "")));
                }
            }
            return listOfTees;
        }
        private void printTees(List<ListOfTees> listOfTees)
        {
            for (int i = 0; i < listOfTees.Count; i++)
            {
                richTextBox7.AppendText(Environment.NewLine + listOfTees[i].teeName);
            }
        }

    }
}
