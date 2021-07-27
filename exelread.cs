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

namespace ExcelRead
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public class MGPipe//класс для хранения данных одной трубы
        {
            public string pipeName;
            public string pipeLong;
        }
        
        public class MGVTD//класс для хранения данных одного отчета ВТД
        {
            public string MGName;
            public string MGDate;
            public string MGPressure;
            public List<MGPipe> MGPipeS = new List<MGPipe>();
        }

        public MGVTD mGVTD=new MGVTD();//создаём экзампляр класса для хранения данных обследования ВТД
        

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                
                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //ObjExcel.Visible = true;

                //Открываем книгу.                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets["Лист1"];

                //Очищаем от старого текста окно вывода.
                richTextBox1.Clear();

                for (int i = 0; i < 10; i++)//пробегаем строки
                {
                  
                        MGPipe mGPipe = new MGPipe();
                        mGPipe.pipeName = Convert.ToString(ObjWorkSheet.Cells[i + 1, 1].Text);
                        mGPipe.pipeLong = Convert.ToString(ObjWorkSheet.Cells[i + 1, 2].Text);

                    mGVTD.MGPipeS.Add(mGPipe);
                  
                }

                //String range1 = Convert.ToString(ObjWorkSheet.Cells[1,1].Text);
                //richTextBox1.Text = richTextBox1.Text + range1;

                for (int j = 0; j < 10; j++)
                {
                    richTextBox1.Text = richTextBox1.Text + Convert.ToString(mGVTD.MGPipeS[j].pipeName)+"***";
                    richTextBox1.Text = richTextBox1.Text + Convert.ToString(mGVTD.MGPipeS[j].pipeLong) + "***";
                }


                
                /* for (int i = 1; i < 101; i++)
                 {
                     //Выбираем область таблицы. (в нашем случае просто ячейку)
                     Microsoft.Office.Interop.Excel.Range range1 = ObjWorkSheet.get_Range(textBox1.Text + i.ToString(), textBox1.Text + i.ToString());
                     //Добавляем полученный из ячейки текст.
                     richTextBox1.Text = richTextBox1.Text + range1.Text.ToString() + "\n";
                     //это чтобы форма прорисовывалась (не подвисала)...
                     //Application.DoEvents();
                 }*/

                //Удаляем приложение (выходим из экселя) - ато будет висеть в процессах!
                ObjExcel.Quit();
            }
        }
    }
}
