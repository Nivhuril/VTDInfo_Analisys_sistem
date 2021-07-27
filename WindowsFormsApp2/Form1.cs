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
using System.Data.SqlClient;
//antipov-db1 OByM6oBaKutjA2xv
namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        public class numbersOfColumns//класс для хранения согласованных номеров столбцов и строк для импорта данных из файла EXCEL
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
            public int columnNumber14;//для категории в трубном журнале
            public int columnNumber15;//для предела прочности в трубном журнале

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
            public int column2Number19;//для даты устранения дефекта
            public int column2Number20;//для номера трубы

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


            //ссылки на ячейки с характеристиками труб и категориями трубопроводов
            public int string4Number1;//первая строка таблицы с типами труб
            public int string4Number2;//последняя строка таблицы с типами труб
            public int column4Number3;
            public int column4Number4;
            public int column4Number5;
            public int column4Number6;
            public int column4Number7;
            public int string4Number8;//первая строка таблицы с категориями трубопровода
            public int string4Number9;//последняя строка таблицы с категориями трубопровода
            public int column4Number10;
            public int column4Number11;
            public int column4Number12;
            public int column4Number13;
        }

        public class MGPipe//класс для хранения данных одной трубы
        {

            public int pipeID;//
            public string pipeNumber;//номер трубы
            public double odometrDist;//дистанция по одометру
            public double thikness;//толщина трубы
            public double pipeLength;//длина трубы
            public string distanceFromReferencePoints;//расстояние от реперных точек
            public string characterFeatures;// характер особенности
            public string clockOrientation;//Ориент., ч:мин
            public double bendOfPipe;//Изгиб, °
            public double jointAngle;//Угол стыка,°
            public string Latitude;//Широта
            public string Longitude;//Долгота
            public double heightAboveSeaLevel;//H, м
            public string note;//Примечание

            //Следующие поля заполняются после обработки отчета
            public string pipelineSectionCategory;//!!!категория участка трубопровода - заполняется при обработке массива
            public string steelGrade;//!!!марка стали - заполняется при обработке массива
            public double yieldPoint;//!!!предел текучести - заполняется при обработке массива
            public double tensileStrength;//!!!предел прочности - заполняется при обработке массива
            public bool itIsTee = false;//истина, если екция является тройником
            //это заполняется путём расчета повреждённости различного типа
            public List<double> corossionDamageList = new List<double>();//поврежденность трубы от коррозии
            public List<double> DentDamageList = new List<double>();//поврежденность трубы от вмятин
            public List<double> JoinDamageList = new List<double>();//поврежденность трубы от дефектов КСС
            
        }

        public class pipelineInfo//класс для хранения информации о трубопроводе (информации об отчете ВТД)
        {
            //в БД добавлен столбец с идентификатором!!!!!!!
            public string pipelineName;//трубопровод (название)
            public string pipelineSection;//участок трубопровода
            public double pipeDiameter;//диаметр трубы
            public string principal;//принципал (хозяин трубы)
            public string examinationDate;//дата обследования
            public double designPressure;// проектное давление
            public double operatingPressure;// рабочее давление
            public string comissioningYear;//год ввода в экспуатацию

        }

        public class anomalyLogLine//класс для хранения строки журнала выявленных аномалий
        {
            //в БД добавлен столбец с идентификатором!!!!!!!
            public string pipeNumber;//номер трубы
            public double odometrDist;//дистанция по одометру
            public double thikness;//толщина трубы
            public string distanceFromTransverseWeld;//расстояние от поперечного шва, м
            public string distanceFromReferencePoints;//расстояние от реперных точек
            public string featuresCharacter;//характер особенности
            public string classOfSize;//класс размера
            public string featuresOreientation;//ориентация
            public double length;//длина
            public double widht;//ширина
            public double depthInProcent;//глубина дефекта в процентах
            public double depthInMm;//глубина дефекта в миллиметрах
            public string extOrInt;//характер локаизации(внутри или снаружи)
            public string KBD;//КБД
            public string defectAssessment;//оценка дефекта
            public string Latitude;//Широта
            public string Longitude;//Долгота
            public double heightAboveSeaLevel;//H, м
            public string note;//Примечание
            public string defectVanishDate;//дата устранения дефекта
            public bool isLostMetal = false;
        }
  
        public class furnishingsLog//класс для хранения строки журнала элементов обустройства
        {
            public string itemNumber;//номер пункта
            public string pipeNumber;//номер трубы
            public double odometrDist;//дистанция по одометру 
            public double pipeLength;//длина трубы
            public string distanceFromTransverseWeld;//расстояние от поперечного шва, м
            public string characterFeatures;// характер особенности
            public string designations;//обозначение
            public string marker;//маркер
            public double distanceToNextFeature;//расстояние до седующей особенности
            public string Latitude;//Широта
            public string Longitude;//Долгота
            public double heightAboveSeaLevel;//H, м
            public string note;//Примечание
        }
        public class pipeCharacteristics//класс для хранения сведений о характеристиках труб
        {
            public double thikness;//толщина трубы
            public string pipeType;//тип трубы
            public string steelGrade;//марка стали
            public double yieldPoint;//предел текучести
            public double tensileStrength;//предел прочности
        }

        public class pipelineSectionCategoryLog//класс для хранения сведений о категориях участков
        {
            public string pipeNumber;//номер трубы
            public double odometrDist;//дистанция по одометру
            public double sectionLength;//длина участка
            public string pipelineSectionCategory;//категория участка трубопровода
        }
        public class MGVTD//класс для хранения данных одного отчета ВТД
        {
            public pipelineInfo pipelineInfo = new pipelineInfo();//информация о газопроводе
            public List<MGPipe> MGPipeS = new List<MGPipe>();//трубный журнал
            public List<anomalyLogLine> anomalyLogLineS = new List<anomalyLogLine>();//журнал выявленных аномалий
            public List<furnishingsLog> furnishingsLogS = new List<furnishingsLog>();//журнал элементов обустройства
            public List<pipeCharacteristics> pipeCharacteristicsLog = new List<pipeCharacteristics>();//характеристики труб
            public List<pipelineSectionCategoryLog> pipelineSectionCategoryLogs = new List<pipelineSectionCategoryLog>();//Категории участков трубопровода
        }
        
        public MGVTD mGVTD = new MGVTD();//создаём экзампляр класса для хранения данных обследования ВТД

        numbersOfColumns NumbersOfColumns = new numbersOfColumns();//создаём экземпляр класса ссылок на столбцы и строки отчета ВТД
        public int summCorr2 = 0;
        public int allPipeCount = 0;//сумма всех труб участка++
        public int allPipeWhithСorrosion = 0;//сумма труб с коррозией++
        public double summCorrosionDamag = 0;//суммаорная поврежденность от коррозии++
        //поврежденность участка от коррозии dk++
        //поврежденность участка от трещин (0)++
        //поврежденность участка от овализации(0)++
        public int allPipeWhithDent=0;//количество труб с вмятинами++
        public double summDentDamag = 0;//суммаорная поврежденность от вмятин++
        //поврежденность участка от вмятин dr++
        //public double allconnectingPartsWhithDefects;//поврежденность тройников++
        public double technicalConditionIndicatorOfPipesAndSDT=0;//показатель технического состояния труб и СДТ++
        public int allPipeWhithJointDefects = 0;//количество труб с дефектами КСС++
        public double summJointDefectsDamag = 0;//суммаорная поврежденность КСС
        //показатель технического состояния кольцевых швов по результатам ВТД (pш)
        //показатель технического состояния по результатам шурфовок (pш*0,85)
        //поврежденность участка от переменных нагрузок (0)
        public double allDefectsWhithСorrosionPlus = 0;
        public double summCorrosionDamagPlus = 0;//суммарная поврежденность от коррозии++
        public int allPipeWhithСorrosionPlus = 0;//сумма труб с коррозией  выше заданного значения
        public string fileName;//переменная для хранения пути к файлу с отчетом
        private void button4_Click(object sender, EventArgs e)//открыть файл отчета ВТД (получить путь к файлу) и прочитать тестовые строки для проверки правильности ссылок
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                button3.Enabled = true;
                findStart();//поиск начала и конца журналов свойств труб и категорий
                tableArdesTest();
            }

        }
        private void tableArdesTest()//метод для проверки правильности адресации ячеек и заполнения экземпляра класса numbersOfColumns()
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            string WorksheetName = textBox9.Text;//получаем название вкладки из формы импотра
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

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

            int string4Number1 = Convert.ToInt16(numericUpDown1.Text);//первая строка таблицы с типами труб
            int string4Number2 = Convert.ToInt16(numericUpDown2.Text);//последняя строка таблицы с типами труб
            int string4Number8 = Convert.ToInt16(numericUpDown3.Text);//первая строка таблицы с категориями трубопровода
            int string4Number9 = Convert.ToInt16(numericUpDown4.Text);//последняя строка таблицы с категориями трубопровода
            
            
            //столбцы с типом трубы
            int column4Number3 = Convert.ToInt16(textBox124.Text);
            int column4Number4 = Convert.ToInt16(textBox125.Text);
            int column4Number5 = Convert.ToInt16(textBox126.Text);
            int column4Number6 = Convert.ToInt16(textBox127.Text);
            int column4Number7 = Convert.ToInt16(textBox128.Text);
            //стоолбцы с категориями участков
            int column4Number10 = Convert.ToInt16(textBox132.Text);
            int column4Number11 = Convert.ToInt16(textBox133.Text);
            int column4Number12 = Convert.ToInt16(textBox134.Text);
            int column4Number13 = Convert.ToInt16(textBox135.Text);



            textBox1.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber1, 4].Text);//выводим прочитанные из тестовой строки таблицы данные в соответствующие поля формы
            textBox2.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber2, 4].Text);
            textBox3.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber3, 4].Text);
            textBox4.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber4, 4].Text);
            textBox5.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber5, 4].Text);
            textBox6.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber6, 4].Text);
            textBox7.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber7, 4].Text);
            textBox8.Text = Convert.ToString(ObjWorkSheet.Cells[stringNumber8, 4].Text);

            textBox119.Text = Convert.ToString(ObjWorkSheet.Cells[string4Number1, column4Number3].Text);
            textBox120.Text = Convert.ToString(ObjWorkSheet.Cells[string4Number1, column4Number4].Text);
            textBox121.Text = Convert.ToString(ObjWorkSheet.Cells[string4Number1, column4Number5].Text);
            textBox122.Text = Convert.ToString(ObjWorkSheet.Cells[string4Number1, column4Number6].Text);
            textBox123.Text = Convert.ToString(ObjWorkSheet.Cells[string4Number1, column4Number7].Text);

            textBox117.Text = Convert.ToString(ObjWorkSheet.Cells[string4Number8, column4Number10].Text);
            textBox118.Text = Convert.ToString(ObjWorkSheet.Cells[string4Number8, column4Number11].Text);
            textBox129.Text = Convert.ToString(ObjWorkSheet.Cells[string4Number8, column4Number12].Text);
            textBox130.Text = Convert.ToString(ObjWorkSheet.Cells[string4Number8, column4Number13].Text);



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

            int columnNumber14 = Convert.ToInt16(textBox141.Text);
            int columnNumber15 = Convert.ToInt16(textBox143.Text);

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

            textBox140.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber14].Text);
            textBox142.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber15].Text);

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
            int column2Number19 = Convert.ToInt16(textBox114.Text);
            int column2Number20 = Convert.ToInt16(textBox116.Text);//для номера трубы

            textBox46.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number1].Text);
            textBox47.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number2].Text);
            textBox48.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number3].Text);
            textBox49.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number4].Text);
            textBox50.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number5].Text);
            textBox51.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number6].Text);
            textBox52.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number7].Text);
            textBox53.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number8].Text);
            textBox54.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number9].Text);
            textBox55.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number10].Text);
            textBox56.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number11].Text);
            textBox57.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number12].Text);
            textBox58.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number13].Text);
            textBox59.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number14].Text);
            textBox60.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number15].Text);
            textBox61.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number16].Text);
            textBox62.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number17].Text);
            textBox63.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number18].Text);
            textBox113.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number19].Text);
            textBox115.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number20].Text);
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

            NumbersOfColumns.columnNumber14 = Convert.ToInt16(textBox141.Text);//категория в трубном журнале
            NumbersOfColumns.columnNumber15 = Convert.ToInt16(textBox143.Text);//предел прочности в трубном журнале

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
            NumbersOfColumns.column2Number19 = Convert.ToInt16(textBox114.Text);
            NumbersOfColumns.column2Number20 = Convert.ToInt16(textBox116.Text);//для номера трубы

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




            NumbersOfColumns.string4Number1 = Convert.ToInt16(numericUpDown1.Text);//первая строка таблицы с типами труб
            NumbersOfColumns.string4Number2 = Convert.ToInt16(numericUpDown2.Text);//последняя строка таблицы с типами труб
            NumbersOfColumns.column4Number3 = Convert.ToInt16(textBox124.Text);
            NumbersOfColumns.column4Number4 = Convert.ToInt16(textBox125.Text);
            NumbersOfColumns.column4Number5 = Convert.ToInt16(textBox126.Text);
            NumbersOfColumns.column4Number6 = Convert.ToInt16(textBox127.Text);
            NumbersOfColumns.column4Number7 = Convert.ToInt16(textBox128.Text);
            NumbersOfColumns.string4Number8 = Convert.ToInt16(numericUpDown4.Text);//первая строка таблицы с категориями трубопровода
            NumbersOfColumns.string4Number9 = Convert.ToInt16(numericUpDown3.Text);//последняя строка таблицы с категориями трубопровода
            NumbersOfColumns.column4Number10 = Convert.ToInt16(textBox132.Text);
            NumbersOfColumns.column4Number11 = Convert.ToInt16(textBox133.Text);
            NumbersOfColumns.column4Number12 = Convert.ToInt16(textBox134.Text);
            NumbersOfColumns.column4Number13 = Convert.ToInt16(textBox135.Text);

            ObjExcel.Quit();


        }
        private MGVTD itIsTee(MGVTD mGVTD)
        {
            MGVTD result = new MGVTD();
            result = mGVTD;
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                if (mGVTD.MGPipeS[i].note.Contains("ройн"))
                {
                    result.MGPipeS[i].itIsTee = true;
                }
                else if (mGVTD.MGPipeS[i].characterFeatures.Contains("ройн"))
                {
                    result.MGPipeS[i].itIsTee = true;
                }
            }
            for (int i = 0; i < mGVTD.furnishingsLogS.Count; i++)
            {
                if (mGVTD.furnishingsLogS[i].note.Contains("ройн"))
                {
                    for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                    {
                        if (String.Equals(mGVTD.furnishingsLogS[i].pipeNumber, mGVTD.MGPipeS[j].pipeNumber))
                        {
                            result.MGPipeS[j].itIsTee = true;
                        }
                    }
                    
                }
                else if (mGVTD.furnishingsLogS[i].characterFeatures.Contains("ройн"))
                {
                    for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                    {
                        if (String.Equals(mGVTD.furnishingsLogS[i].pipeNumber, mGVTD.MGPipeS[j].pipeNumber))
                        {
                            result.MGPipeS[j].itIsTee = true;
                        }
                    }
                }
            }
            return result;
        }
        private void findStart()//ищем начало и конец журнала категорий и журнала типов труб на первой странице
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).

            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet33;
            string WorksheetName33 = textBox9.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet33 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName33];
            //Найдём начало и конец журнала типов труб и категорий участков в отчете ВТД
            int startPosition = 0;//начало журнала характеристик труб
            int finishPosition = 0;// конец журнала характеристик труб
            int startPosition2 = 0;//начало журнала характеристик труб
            int finishPosition2 = 0;// конец журнала характеристик труб

            int column4Number3 = Convert.ToInt16(textBox124.Text);
            int column4Number4 = Convert.ToInt16(textBox125.Text);
            int column4Number5 = Convert.ToInt16(textBox126.Text);
            int column4Number6 = Convert.ToInt16(textBox127.Text);
            int column4Number7 = Convert.ToInt16(textBox128.Text);//предел прочности
            //стоолбцы с категориями участков
            int column4Number10 = Convert.ToInt16(textBox132.Text);
            int column4Number11 = Convert.ToInt16(textBox133.Text);
            int column4Number12 = Convert.ToInt16(textBox134.Text);
            int column4Number13 = Convert.ToInt16(textBox135.Text);
            //NumbersOfColumns.string4Number1 = Convert.ToInt16(numericUpDown1.Text);//первая строка таблицы с типами труб
            //NumbersOfColumns.string4Number2 = Convert.ToInt16(numericUpDown2.Text);//последняя строка таблицы с типами труб
            NumbersOfColumns.column4Number3 = Convert.ToInt16(textBox124.Text);
            NumbersOfColumns.column4Number4 = Convert.ToInt16(textBox125.Text);
            NumbersOfColumns.column4Number5 = Convert.ToInt16(textBox126.Text);
            NumbersOfColumns.column4Number6 = Convert.ToInt16(textBox127.Text);
            NumbersOfColumns.column4Number7 = Convert.ToInt16(textBox128.Text);
            //NumbersOfColumns.string4Number8 = Convert.ToInt16(numericUpDown4.Text);//первая строка таблицы с категориями трубопровода
            //NumbersOfColumns.string4Number9 = Convert.ToInt16(numericUpDown3.Text);//последняя строка таблицы с категориями трубопровода
            NumbersOfColumns.column4Number10 = Convert.ToInt16(textBox132.Text);
            NumbersOfColumns.column4Number11 = Convert.ToInt16(textBox133.Text);
            NumbersOfColumns.column4Number12 = Convert.ToInt16(textBox134.Text);
            NumbersOfColumns.column4Number13 = Convert.ToInt16(textBox135.Text);

            
            
                int d = 1;
                bool looking = true;
                bool firstIsFind = false;//маркер, что нашли первый участок
                while (looking)//ищем начало и коней первого журнала
                {
                    string pipeName = Convert.ToString(ObjWorkSheet33.Cells[d, NumbersOfColumns.column4Number3].Text);
                    if (pipeName.Contains("олщ"))
                    {
                        startPosition = d + 1;
                        firstIsFind = true;
                    }
                    if (firstIsFind)
                    {
                        if (String.IsNullOrWhiteSpace(pipeName))
                        {
                            finishPosition = d - 1;
                            looking = false;
                        }
                    }
                    d++;
                }

                d = finishPosition;
                looking = true;
                firstIsFind = false;//маркер, что нашли первый участок
                while (looking)//ищем начало и коней второго журнала
                {
                    string pipeName = Convert.ToString(ObjWorkSheet33.Cells[d, NumbersOfColumns.column4Number3].Text);
                    if (pipeName.Contains("руб"))
                    {
                        startPosition2 = d + 1;
                        firstIsFind = true;
                    }
                    if (firstIsFind)
                    {
                        if (String.IsNullOrWhiteSpace(pipeName))
                        {
                            finishPosition2 = d - 1;
                            looking = false;
                        }
                    }
                    d++;
                }
                NumbersOfColumns.string4Number1 = startPosition;
                NumbersOfColumns.string4Number2 = finishPosition;
                NumbersOfColumns.string4Number8 = startPosition2;
                NumbersOfColumns.string4Number9 = finishPosition2;
                numericUpDown1.Value = Convert.ToInt32(startPosition);
                numericUpDown2.Value = Convert.ToInt32(finishPosition);
                numericUpDown4.Value = Convert.ToInt32(startPosition2);
                numericUpDown3.Value = Convert.ToInt32(finishPosition2);
                ObjExcel.Quit();
            

        }

        //************методы для чтения различных разделов отчета ВТД*********************
        private void readToClassPipeInfo()//метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            string WorksheetName = textBox9.Text;//получаем название вкладки из формы импотра (данные о газопроводе)
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            richTextBox1.AppendText(Environment.NewLine + "Выполняется чтение данных о газопроводе...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            //int pipeListCount = Convert.ToInt16(textBox95.Text);//получаем длину журнала из формы
            pipelineInfo PipelineInfo = new pipelineInfo();

            PipelineInfo.pipelineName = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber1, 4].Text);//трубопровод (название)
            PipelineInfo.pipelineSection = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber2, 4].Text);//участок трубопровода
            PipelineInfo.pipeDiameter = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber3, 4].Text);//диаметр трубы
            PipelineInfo.principal = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber4, 4].Text);//принципал (хозяин трубы)
            PipelineInfo.examinationDate = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber5, 4].Text);//дата обследования
            PipelineInfo.designPressure = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber6, 4].Text);// проектное давление
            PipelineInfo.operatingPressure = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber7, 4].Text);// рабочее давление
            PipelineInfo.comissioningYear = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber8, 4].Text);//год ввода в экспуатацию

            mGVTD.pipelineInfo = PipelineInfo;

            richTextBox1.AppendText(Environment.NewLine + "Сведения о газопроводе получены и записаны в экземпляр класса");
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
        }

        private void readToClassPipeLog()//метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox42.Text;//получаем название вкладки из формы импотра (трубный журнал)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];

            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка трубного журнала...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            int pipeListCount = Convert.ToInt16(textBox95.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            for (int i = 1; i < pipeListCount; i++)//чтение трубного журнала
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
                mGPipe.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber13].Text);
                mGVTD.MGPipeS.Add(mGPipe);


                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor == 10)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }


            }

            richTextBox1.AppendText(Environment.NewLine + "Массив данных из трубного журнала прочитан и записан в экземпляр класса");
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
        }

        private void readToClassAnomalyLogLine()//метод для чтения из файла отчета строк журнала аномалий
        {

            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox45.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала выявленных аномалий...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            int pipeListCount = Convert.ToInt16(textBox110.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            for (int i = 1; i < pipeListCount; i++)//чтение трубного журнала
            {
                anomalyLogLine AnomalyLogLine = new anomalyLogLine();//создаём экземпляр класса строки журнала аномалий

                AnomalyLogLine.odometrDist = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number1].Text);//дистанция по одометру
                AnomalyLogLine.thikness = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number2].Text);//толщина трубы
                AnomalyLogLine.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number3].Text);//расстояние от поперечного шва, м
                AnomalyLogLine.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number4].Text);//расстояние от реперных точек
                AnomalyLogLine.featuresCharacter = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number5].Text);//характер особенности
                AnomalyLogLine.classOfSize = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number6].Text);//класс размера
                AnomalyLogLine.featuresOreientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number7].Text);//ориентация
                AnomalyLogLine.length = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number8].Text);//длина
                AnomalyLogLine.widht = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number9].Text);//ширина
                AnomalyLogLine.depthInProcent = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number10].Text);//глубина дефекта в процентах
                AnomalyLogLine.depthInMm = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number11].Text);//глубина дефекта в миллиметрах
                AnomalyLogLine.extOrInt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number12].Text);//характер локаизации(внутри или снаружи)
                AnomalyLogLine.KBD = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number13].Text);//КБД
                AnomalyLogLine.defectAssessment = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number14].Text);//оценка дефекта
                AnomalyLogLine.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number15].Text);//Широта
                AnomalyLogLine.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number16].Text);//Долгота
                AnomalyLogLine.heightAboveSeaLevel = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number17].Text);//H, м
                AnomalyLogLine.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number18].Text);//Примечание
                AnomalyLogLine.defectVanishDate = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number19].Text);//Примечание
                mGVTD.anomalyLogLineS.Add(AnomalyLogLine);//добавляем заполненный экземпляр класса к списку

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor == 10)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }

            }

            richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала выявленных аномалий прочитан и записан в экземпляр класса");
            richTextBox1.AppendText(Environment.NewLine + "==========================================");

        }

        private void readToClassFurnishingsLogLine()//метод для чтения из файла отчета строк журнала элементов обустройства
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox108.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала элементов обустройства...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            int pipeListCount = Convert.ToInt16(textBox111.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора

            for (int i = 0; i < pipeListCount; i++)
            {
                furnishingsLog FurnishingsLog = new furnishingsLog();//создаём экземпляр класса строки журнала аномалий

                FurnishingsLog.itemNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number1].Text);//номер пункта
                FurnishingsLog.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number2].Text);//номер трубы
                FurnishingsLog.odometrDist = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number3].Text);//дистанция по одометру 
                FurnishingsLog.pipeLength = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number4].Text);//длина трубы
                FurnishingsLog.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number5].Text);//расстояние от поперечного шва, м
                FurnishingsLog.characterFeatures = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number6].Text);// характер особенности
                FurnishingsLog.designations = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number7].Text);//обозначение
                FurnishingsLog.marker = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number8].Text);//маркер
                FurnishingsLog.distanceToNextFeature = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number9].Text);//расстояние до седующей особенности
                FurnishingsLog.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number10].Text);//Широта
                FurnishingsLog.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number11].Text);//Долгота
                FurnishingsLog.heightAboveSeaLevel = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number12].Text);//H, м
                FurnishingsLog.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number13].Text);//Примечание

                mGVTD.furnishingsLogS.Add(FurnishingsLog);//добавляем заполненный экземпляр класса к списку

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor > 4)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }

            }
            richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала элементов обустройства прочитан и записан в экземпляр класса");
            richTextBox1.AppendText(Environment.NewLine + "==========================================");


        }

        //*********это модификации методов для более универсального применения***********
        private pipelineInfo operatingReadToClassPipeInfo(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            string WorksheetName = textBox9.Text;//получаем название вкладки из формы импотра (данные о газопроводе)
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            richTextBox1.AppendText(Environment.NewLine + "Выполняется чтение данных о газопроводе...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            //int pipeListCount = Convert.ToInt16(textBox95.Text);//получаем длину журнала из формы
            pipelineInfo PipelineInfo = new pipelineInfo();

            PipelineInfo.pipelineName = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber1, 4].Text);//трубопровод (название)
            PipelineInfo.pipelineSection = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber2, 4].Text);//участок трубопровода
            String txt = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber3, 4].Text);
            PipelineInfo.pipeDiameter = Convert.ToDouble(txt.Replace(".", ","));//диаметр трубы
            PipelineInfo.principal = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber4, 4].Text);//принципал (хозяин трубы)
            PipelineInfo.examinationDate = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber5, 4].Text);//дата обследования
            txt = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber6, 4].Text);
            PipelineInfo.designPressure = Convert.ToDouble(txt.Replace(".", ","));// проектное давление
            txt = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber7, 4].Text);
            PipelineInfo.operatingPressure = Convert.ToDouble(txt.Replace(".", ","));// рабочее давление
            PipelineInfo.comissioningYear = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber8, 4].Text);//год ввода в экспуатацию

            //mGVTD.pipelineInfo = PipelineInfo;

            richTextBox1.AppendText(Environment.NewLine + "Сведения о газопроводе получены и записаны в экземпляр класса");
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return PipelineInfo;
        }

        private List<MGPipe> operatingReadToClassPipeLog(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета ВТД трубного журнала
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox42.Text;//получаем название вкладки из формы импотра (трубный журнал)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];

            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка трубного журнала...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            int pipeListCount = Convert.ToInt16(textBox95.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            List<MGPipe> OMGPipeS = new List<MGPipe>();//трубный журнал

            for (int i = 1; i < pipeListCount + 1; i++)//чтение трубного журнала
            {
                MGPipe mGPipe = new MGPipe();
                mGPipe.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber1].Text);


                String txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number8].Text);
                try
                {
                    mGPipe.odometrDist = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.odometrDist = 0;
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


                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber8].Text);
                try
                {
                    mGPipe.bendOfPipe = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.bendOfPipe = 0;
                }

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


                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber12].Text);
                try
                {
                    mGPipe.heightAboveSeaLevel = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.heightAboveSeaLevel = 0;
                }

                mGPipe.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber13].Text);
                OMGPipeS.Add(mGPipe);


                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor == 10)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }


            }

            richTextBox1.AppendText(Environment.NewLine + "Массив данных из трубного журнала прочитан и записан в экземпляр класса");
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return OMGPipeS;
        }

        private List<anomalyLogLine> operatingReadToClassAnomalyLog(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета строк журнала аномалий
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox45.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала выявленных аномалий...");
            richTextBox1.AppendText(Environment.NewLine + "->*");

            List<anomalyLogLine> anomalyLogLineS = new List<anomalyLogLine>();
            int pipeListCount = Convert.ToInt16(textBox110.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            for (int i = 1; i < pipeListCount; i++)//чтение трубного журнала
            {
                anomalyLogLine AnomalyLogLine = new anomalyLogLine();//создаём экземпляр класса строки журнала аномалий
                AnomalyLogLine.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number20].Text);//расстояние от поперечного шва, м
                //String txt= Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber2].Text);
                //mGPipe.odometrDist = Convert.ToDouble(txt.Replace(".",","));
                String txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number1].Text);
                AnomalyLogLine.odometrDist = Convert.ToDouble(txt.Replace(".", ","));//дистанция по одометру
                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number2].Text);
                AnomalyLogLine.thikness = Convert.ToDouble(txt.Replace(".", ","));//толщина трубы
                AnomalyLogLine.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number3].Text);//расстояние от поперечного шва, м
                AnomalyLogLine.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number4].Text);//расстояние от реперных точек
                AnomalyLogLine.featuresCharacter = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number5].Text);//характер особенности
                AnomalyLogLine.classOfSize = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number6].Text);//класс размера
                AnomalyLogLine.featuresOreientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number7].Text);//ориентация


                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number8].Text);
                try
                {
                    AnomalyLogLine.length = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.length = 0;
                }

                //txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number9].Text);
                //AnomalyLogLine.widht = Convert.ToDouble(txt.Replace(".", ","));//ширина
                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number9].Text);
                try
                {
                    AnomalyLogLine.widht = Convert.ToDouble(txt.Replace(".", ","));//ширина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.widht = 0;
                }



                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number10].Text);
                try
                {
                    AnomalyLogLine.depthInProcent = Convert.ToDouble(txt.Replace(".", ","));//глубина дефекта в процентах
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.depthInProcent = 0;
                }
                //AnomalyLogLine.depthInProcent = Convert.ToDouble(txt.Replace(".", ","));//глубина дефекта в процентах



                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number11].Text);
                try
                {
                    AnomalyLogLine.depthInMm = Convert.ToDouble(txt.Replace(".", ","));//глубина дефекта в миллиметрах
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.depthInMm = 0;
                }
                //AnomalyLogLine.depthInMm = Convert.ToDouble(txt.Replace(".", ","));//глубина дефекта в миллиметрах
                AnomalyLogLine.extOrInt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number12].Text);//характер локаизации(внутри или снаружи)
                AnomalyLogLine.KBD = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number13].Text);//КБД
                AnomalyLogLine.defectAssessment = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number14].Text);//оценка дефекта
                AnomalyLogLine.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number15].Text);//Широта
                AnomalyLogLine.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number16].Text);//Долгота
                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number17].Text);
                try
                {
                    AnomalyLogLine.heightAboveSeaLevel = Convert.ToDouble(txt.Replace(".", ","));//H, м
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.heightAboveSeaLevel = 0;
                }
                //AnomalyLogLine.heightAboveSeaLevel = Convert.ToDouble(txt.Replace(".", ","));//H, м
                AnomalyLogLine.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number18].Text);//Примечание
                AnomalyLogLine.defectVanishDate = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number19].Text);//Примечание
                anomalyLogLineS.Add(AnomalyLogLine);//добавляем заполненный экземпляр класса к списку

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor == 10)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }

            }
            richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала выявленных аномалий прочитан и записан в экземпляр класса");
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return anomalyLogLineS;
        }
        private List<furnishingsLog> operatingReadToClassFurnishingsLog(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета строк журнала элементов обустройства
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox108.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            List<furnishingsLog> furnishingsLogS = new List<furnishingsLog>();
            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала элементов обустройства...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            int pipeListCount = Convert.ToInt16(textBox111.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора

            for (int i = 0; i < pipeListCount; i++)
            {
                furnishingsLog FurnishingsLog = new furnishingsLog();//создаём экземпляр класса строки журнала аномалий

                FurnishingsLog.itemNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number1].Text);//номер пункта
                FurnishingsLog.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number2].Text);//номер трубы

                String txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number3].Text);
                try
                {
                    FurnishingsLog.odometrDist = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    FurnishingsLog.odometrDist = 0;
                }

                //FurnishingsLog.odometrDist = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number3].Text);//дистанция по одометру 

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number4].Text);
                try
                {
                    FurnishingsLog.pipeLength = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    FurnishingsLog.pipeLength = 0;
                }
                //FurnishingsLog.pipeLength = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number4].Text);//длина трубы
                FurnishingsLog.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number5].Text);//расстояние от поперечного шва, м
                FurnishingsLog.characterFeatures = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number6].Text);// характер особенности
                FurnishingsLog.designations = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number7].Text);//обозначение
                FurnishingsLog.marker = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number8].Text);//маркер

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number9].Text);
                try
                {
                    FurnishingsLog.distanceToNextFeature = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    FurnishingsLog.distanceToNextFeature = 0;
                }
                //FurnishingsLog.distanceToNextFeature = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number9].Text);//расстояние до седующей особенности
                FurnishingsLog.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number10].Text);//Широта
                FurnishingsLog.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number11].Text);//Долгота

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number12].Text);
                try
                {
                    FurnishingsLog.heightAboveSeaLevel = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    FurnishingsLog.heightAboveSeaLevel = 0;
                }
                //FurnishingsLog.heightAboveSeaLevel = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number12].Text);//H, м
                FurnishingsLog.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number13].Text);//Примечание

                furnishingsLogS.Add(FurnishingsLog);//добавляем заполненный экземпляр класса к списку

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor > 4)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }

            }
            richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала элементов обустройства прочитан и записан в экземпляр класса");
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return furnishingsLogS;

        }
        private List<furnishingsLog> operatingReadToClassFurnishingsLogAutoFin(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета строк журнала элементов обустройства
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox108.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            List<furnishingsLog> furnishingsLogS = new List<furnishingsLog>();
            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала элементов обустройства...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            //int pipeListCount = Convert.ToInt16(textBox111.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора

            int i = 1;
            bool mark = true;
            while (mark)
            {
                furnishingsLog FurnishingsLog = new furnishingsLog();//создаём экземпляр класса строки журнала аномалий

                FurnishingsLog.itemNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number1].Text);//номер пункта
                FurnishingsLog.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number2].Text);//номер трубы

                String txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number3].Text);
                try
                {
                    FurnishingsLog.odometrDist = Convert.ToDouble(txt.Replace(".", ","));//длина
                    
                }
                catch (Exception)
                {
                    FurnishingsLog.odometrDist = 0;
                }

                //FurnishingsLog.odometrDist = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number3].Text);//дистанция по одометру 

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number4].Text);
                try
                {
                    FurnishingsLog.pipeLength = Convert.ToDouble(txt.Replace(".", ","));//длина
                    
                }
                catch (Exception)
                {
                    FurnishingsLog.pipeLength = 0;
                }
                //FurnishingsLog.pipeLength = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number4].Text);//длина трубы
                FurnishingsLog.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number5].Text);//расстояние от поперечного шва, м
                FurnishingsLog.characterFeatures = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number6].Text);// характер особенности
                FurnishingsLog.designations = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number7].Text);//обозначение
                FurnishingsLog.marker = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number8].Text);//маркер

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number9].Text);
                try
                {
                    FurnishingsLog.distanceToNextFeature = Convert.ToDouble(txt.Replace(".", ","));//длина
                    
                }
                catch (Exception)
                {
                    FurnishingsLog.distanceToNextFeature = 0;
                }
                //FurnishingsLog.distanceToNextFeature = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number9].Text);//расстояние до седующей особенности
                FurnishingsLog.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number10].Text);//Широта
                FurnishingsLog.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number11].Text);//Долгота

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number12].Text);
                try
                {
                    FurnishingsLog.heightAboveSeaLevel = Convert.ToDouble(txt.Replace(".", ","));//длина
                    
                }
                catch (Exception)
                {
                    FurnishingsLog.heightAboveSeaLevel = 0;
                }
                //FurnishingsLog.heightAboveSeaLevel = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number12].Text);//H, м
                FurnishingsLog.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number13].Text);//Примечание

                
                if (String.IsNullOrWhiteSpace(FurnishingsLog.characterFeatures))
                {
                    mark = false;//дошли до конца трубного журлала
                }
                else
                {
                    furnishingsLogS.Add(FurnishingsLog);//добавляем заполненный экземпляр класса к списку
                    //richTextBox1.AppendText(FurnishingsLog.characterFeatures);
                }


                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor > 9)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }
                i++;
            }
            textBox111.Text = Convert.ToString(i);//записываем в поле номер последней строки
            richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала элементов обустройства прочитан, количество строк:"+ furnishingsLogS.Count);
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return furnishingsLogS;

        }
        private List<pipeCharacteristics> operatingReadToClassPipeCharacteristics(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета строк журнала элементов обустройства
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox137.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            List<pipeCharacteristics> pipeCharacteristicseS = new List<pipeCharacteristics>();
            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала характеристик труб...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            //int pipeListCount = Convert.ToInt16(textBox111.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора

            

            for (int i = NumbersOfColumns.string4Number1; i < NumbersOfColumns.string4Number2 + 1; i++)
            {

                pipeCharacteristics PipeCharacteristics = new pipeCharacteristics();//создаём экземпляр класса строки

                String txt = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.column4Number3].Text);
                try
                {
                    PipeCharacteristics.thikness = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    PipeCharacteristics.thikness = 0;
                }

                PipeCharacteristics.pipeType = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.column4Number4].Text);
                PipeCharacteristics.steelGrade = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.column4Number5].Text);

                txt = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.column4Number6].Text);
                try
                {
                    PipeCharacteristics.yieldPoint = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    PipeCharacteristics.yieldPoint = 0;
                }

                txt = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.column4Number7].Text);
                try
                {
                    PipeCharacteristics.tensileStrength = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    PipeCharacteristics.tensileStrength = 0;
                }
                pipeCharacteristicseS.Add(PipeCharacteristics);//добавляем заполненный экземпляр класса к списку

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor > 4)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }

            }
            richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала характеристик труб прочитан. Количество строк:"+ pipeCharacteristicseS.Count);
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return pipeCharacteristicseS;
        }
        private List<pipelineSectionCategoryLog> operatingReadToClassPipelineSectionCategoryLog(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета строк журнала элементов обустройства
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox138.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            List<pipelineSectionCategoryLog> pipelineSectionCategoryLogS = new List<pipelineSectionCategoryLog>();
            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала категорий участков...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            //int pipeListCount = Convert.ToInt16(textBox111.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора

            for (int i = NumbersOfColumns.string4Number8; i < NumbersOfColumns.string4Number9 + 1; i++)
            {

                pipelineSectionCategoryLog PipelineSectionCategoryLog = new pipelineSectionCategoryLog();//создаём экземпляр класса строки журнала категорий

                PipelineSectionCategoryLog.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.column4Number10].Text);

                String txt = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.column4Number11].Text);
                try
                {
                    PipelineSectionCategoryLog.odometrDist = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    PipelineSectionCategoryLog.odometrDist = 0;
                }

                txt = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.column4Number12].Text);
                try
                {
                    PipelineSectionCategoryLog.sectionLength = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    PipelineSectionCategoryLog.sectionLength = 0;
                }

                PipelineSectionCategoryLog.pipelineSectionCategory = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.column4Number13].Text);

                pipelineSectionCategoryLogS.Add(PipelineSectionCategoryLog);//добавляем заполненный экземпляр класса к списку

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor > 4)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }

            }
            richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала категорий учавстков прочитан. Количество строк:"+ pipelineSectionCategoryLogS.Count);
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return pipelineSectionCategoryLogS;

        }
        //********************************************************************************
        private List<MGPipe> shortOperatingReadToClassPipeLog(string fileName, numbersOfColumns NumbersOfColumns)//КОРОТКИЙ!!!метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox42.Text;//получаем название вкладки из формы импотра (трубный журнал)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];

            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка трубного журнала...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
            int pipeListCount = Convert.ToInt16(textBox95.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            List<MGPipe> OMGPipeS = new List<MGPipe>();//трубный журнал

            for (int i = 1; i < pipeListCount + 1; i++)//чтение трубного журнала
            {
                MGPipe mGPipe = new MGPipe();
                mGPipe.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber1].Text);

                String txt;
                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber2].Text);
                try
                {
                    mGPipe.odometrDist = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.odometrDist = 0;
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

                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber4].Text);
                try
                {
                    mGPipe.pipeLength = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.pipeLength = 0;
                }*/


                //mGPipe.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber5].Text);
                mGPipe.characterFeatures = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber6].Text);
                //mGPipe.clockOrientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber7].Text);


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

                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber9].Text);
                try
                {
                    mGPipe.jointAngle = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.jointAngle = 0;
                }*/

                //mGPipe.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber10].Text);
                //mGPipe.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber11].Text);


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
                OMGPipeS.Add(mGPipe);


                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor == 100)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }
            }

            richTextBox1.AppendText(Environment.NewLine + "Массив данных из трубного журнала прочитан и записан в экземпляр класса");
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            return OMGPipeS;
        }
        private List<MGPipe> shortOperatingReadToClassPipeLogAutoFin(string fileName, numbersOfColumns NumbersOfColumns)//с автофинишем/КОРОТКИЙ!!!метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox42.Text;//получаем название вкладки из формы импотра (трубный журнал)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];

            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка трубного журнала...");
            richTextBox1.AppendText(Environment.NewLine + "->*");
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
                    richTextBox1.AppendText(Environment.NewLine + "^");
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

                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber4].Text);
                try
                {
                    mGPipe.pipeLength = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.pipeLength = 0;
                }*/


                //mGPipe.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber5].Text);
                mGPipe.characterFeatures = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber6].Text);
                //mGPipe.clockOrientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber7].Text);


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

                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber9].Text);
                try
                {
                    mGPipe.jointAngle = Convert.ToDouble(txt.Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.jointAngle = 0;
                }*/

                //mGPipe.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber10].Text);
                //mGPipe.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber11].Text);


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
                if (String.IsNullOrEmpty(Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber14].Text))==false)
                {
                    string localCategory= Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber14].Text);
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
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }
                i++;   
            }

            textBox95.Text = Convert.ToString(i);//записываем в поле количество труб
            richTextBox1.AppendText(Environment.NewLine + "Массив данных из трубного журнала прочитан, количество труб: "+ OMGPipeS.Count);
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return OMGPipeS;
        }
        private List<anomalyLogLine> shortOperatingReadToClassAnomalyLog(string fileName, numbersOfColumns NumbersOfColumns)//КОРОТКИЙ!!!метод для чтения из файла отчета строк журнала аномалий
        {

            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox45.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала выявленных аномалий...");
            richTextBox1.AppendText(Environment.NewLine + "->*");

            List<anomalyLogLine> anomalyLogLineS = new List<anomalyLogLine>();
            int pipeListCount = Convert.ToInt16(textBox110.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            for (int i = 1; i < pipeListCount; i++)//чтение трубного журнала
            {
                anomalyLogLine AnomalyLogLine = new anomalyLogLine();//создаём экземпляр класса строки журнала аномалий
                AnomalyLogLine.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number20].Text);//расстояние от поперечного шва, м

                String txt;
                /*String txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number1].Text);
                try
                {
                    AnomalyLogLine.odometrDist = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.odometrDist = 0;
                }*/


                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number2].Text);
                try
                {
                    AnomalyLogLine.thikness = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.thikness = 0;
                }

                AnomalyLogLine.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number3].Text);//расстояние от поперечного шва, м
                //AnomalyLogLine.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number4].Text);//расстояние от реперных точек
                AnomalyLogLine.featuresCharacter = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number5].Text);//характер особенности
                //AnomalyLogLine.classOfSize = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number6].Text);//класс размера
                //AnomalyLogLine.featuresOreientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number7].Text);//ориентация


                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number8].Text);
                try
                {
                    AnomalyLogLine.length = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.length = 0;
                }

                //txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number9].Text);
                //AnomalyLogLine.widht = Convert.ToDouble(txt.Replace(".", ","));//ширина
                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number9].Text);
                try
                {
                    AnomalyLogLine.widht = Convert.ToDouble(txt.Replace(".", ","));//ширина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.widht = 0;
                }



                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number10].Text);
                try
                {
                    AnomalyLogLine.depthInProcent = Convert.ToDouble(txt.Replace(".", ","));//глубина дефекта в процентах
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.depthInProcent = 0;
                }*/


                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number11].Text);
                try
                {
                    AnomalyLogLine.depthInMm = Convert.ToDouble(txt.Replace(".", ","));//глубина дефекта в миллиметрах
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.depthInMm = 0;
                }
                //AnomalyLogLine.depthInMm = Convert.ToDouble(txt.Replace(".", ","));//глубина дефекта в миллиметрах
                //AnomalyLogLine.extOrInt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number12].Text);//характер локаизации(внутри или снаружи)
                //AnomalyLogLine.KBD = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number13].Text);//КБД
                AnomalyLogLine.defectAssessment = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number14].Text);//оценка дефекта
                //AnomalyLogLine.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number15].Text);//Широта
                //AnomalyLogLine.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number16].Text);//Долгота
                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number17].Text);
                try
                {
                    AnomalyLogLine.heightAboveSeaLevel = Convert.ToDouble(txt.Replace(".", ","));//H, м
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.heightAboveSeaLevel = 0;
                }*/
                //AnomalyLogLine.heightAboveSeaLevel = Convert.ToDouble(txt.Replace(".", ","));//H, м
                AnomalyLogLine.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number18].Text);//Примечание
                AnomalyLogLine.defectVanishDate = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number19].Text);//Примечание
                anomalyLogLineS.Add(AnomalyLogLine);//добавляем заполненный экземпляр класса к списку

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor == 100)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }

            }

            richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала выявленных аномалий прочитан и записан в экземпляр класса");
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            return anomalyLogLineS;
        }
        private List<anomalyLogLine> shortOperatingReadToClassAnomalyLogAutoFin(string fileName, numbersOfColumns NumbersOfColumns)//с автофинишем/КОРОТКИЙ!!!метод для чтения из файла отчета строк журнала аномалий
        {

            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox45.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала выявленных аномалий...");
            richTextBox1.AppendText(Environment.NewLine + "->*");

            List<anomalyLogLine> anomalyLogLineS = new List<anomalyLogLine>();
            int pipeListCount = Convert.ToInt16(textBox110.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            int i = 1;
            bool mark = true;
            while (mark)
            {
                anomalyLogLine AnomalyLogLine = new anomalyLogLine();//создаём экземпляр класса строки журнала аномалий
                AnomalyLogLine.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number20].Text);//расстояние от поперечного шва, м

                String txt;
                /*String txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number1].Text);
                try
                {
                    AnomalyLogLine.odometrDist = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.odometrDist = 0;
                }*/


                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number2].Text);
                try
                {
                    AnomalyLogLine.thikness = Convert.ToDouble(txt.Replace(".", ","));//длина
                    
                }
                catch (Exception)
                {
                    AnomalyLogLine.thikness = 0;
                }

                AnomalyLogLine.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number3].Text);//расстояние от поперечного шва, м
                //AnomalyLogLine.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number4].Text);//расстояние от реперных точек
                AnomalyLogLine.featuresCharacter = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number5].Text);//характер особенности
                //AnomalyLogLine.classOfSize = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number6].Text);//класс размера
                //AnomalyLogLine.featuresOreientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number7].Text);//ориентация


                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number8].Text);
                try
                {
                    AnomalyLogLine.length = Convert.ToDouble(txt.Replace(".", ","));//длина
                    
                }
                catch (Exception)
                {
                    AnomalyLogLine.length = 0;
                }

                //txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number9].Text);
                //AnomalyLogLine.widht = Convert.ToDouble(txt.Replace(".", ","));//ширина
                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number9].Text);
                try
                {
                    AnomalyLogLine.widht = Convert.ToDouble(txt.Replace(".", ","));//ширина
                    
                }
                catch (Exception)
                {
                    AnomalyLogLine.widht = 0;
                }



                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number10].Text);
                try
                {
                    AnomalyLogLine.depthInProcent = Convert.ToDouble(txt.Replace(".", ","));//глубина дефекта в процентах                  
                }
                catch (Exception)
                {
                    AnomalyLogLine.depthInProcent = 0;
                }


                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number11].Text);
                try
                {
                    AnomalyLogLine.depthInMm = Convert.ToDouble(txt.Replace(".", ","));//глубина дефекта в миллиметрах                    
                }
                catch (Exception)
                {
                    AnomalyLogLine.depthInMm = 0;
                }
                //AnomalyLogLine.depthInMm = Convert.ToDouble(txt.Replace(".", ","));//глубина дефекта в миллиметрах
                //AnomalyLogLine.extOrInt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number12].Text);//характер локаизации(внутри или снаружи)
                //AnomalyLogLine.KBD = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number13].Text);//КБД
                AnomalyLogLine.defectAssessment = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number14].Text);//оценка дефекта
                //AnomalyLogLine.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number15].Text);//Широта
                //AnomalyLogLine.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number16].Text);//Долгота
                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number17].Text);
                try
                {
                    AnomalyLogLine.heightAboveSeaLevel = Convert.ToDouble(txt.Replace(".", ","));//H, м
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.heightAboveSeaLevel = 0;
                }*/
                //AnomalyLogLine.heightAboveSeaLevel = Convert.ToDouble(txt.Replace(".", ","));//H, м
                AnomalyLogLine.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number18].Text);//Примечание
                AnomalyLogLine.defectVanishDate = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number19].Text);//Примечание

                if (String.IsNullOrWhiteSpace(AnomalyLogLine.pipeNumber))
                {
                    mark = false;//дошли до конца трубного журлала
                }
                else
                {
                    anomalyLogLineS.Add(AnomalyLogLine);//добавляем заполненный экземпляр класса к списку
                }
                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                i++;
                if (incrementor == 100)
                {
                    richTextBox1.AppendText("*");
                    incrementor = 0;
                }


            }
            textBox110.Text = Convert.ToString(i);//записываем в поле количество труб
            richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала выявленных аномалий прочитан, количество дефектов:"+ anomalyLogLineS.Count);
            richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return anomalyLogLineS;
        }
        //********************************************************************************        
        private MGVTD PipeLogWithCategory(MGVTD mGVTD)//расстановка категорий и характеристик труб
        {
            MGVTD MgvtdNew = new MGVTD();
            MgvtdNew = mGVTD;
            richTextBox1.AppendText(Environment.NewLine + "Выполняется расстановка характеристик труб в трубном журнале...(" + mGVTD.MGPipeS.Count + " труб)");
            for (int i = 0; i < mGVTD.pipeCharacteristicsLog.Count; i++)
            {
                for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                {
                    if (String.IsNullOrWhiteSpace(MgvtdNew.MGPipeS[j].steelGrade))
                    {
                        if (mGVTD.MGPipeS[j].thikness == mGVTD.pipeCharacteristicsLog[i].thikness)
                        {
                            mGVTD.MGPipeS[j].tensileStrength = mGVTD.pipeCharacteristicsLog[i].tensileStrength;
                            mGVTD.MGPipeS[j].steelGrade = mGVTD.pipeCharacteristicsLog[i].steelGrade;
                            mGVTD.MGPipeS[j].yieldPoint = mGVTD.pipeCharacteristicsLog[i].yieldPoint;
                            //richTextBox2.AppendText(Environment.NewLine + "Труба "+ mGVTD.MGPipeS[j].pipeNumber + " марка стали " + mGVTD.MGPipeS[j].steelGrade + " предел текучести " + mGVTD.MGPipeS[j].yieldPoint + " предел прочности " + mGVTD.MGPipeS[j].tensileStrength);
                        }
                    }
                    /*if (j % 100 == 0)
                    {
                        richTextBox2.AppendText("*");
                    }*/
                }
            }
            richTextBox1.AppendText(Environment.NewLine + "Расстановка характеристик труб выполнена");
            richTextBox1.AppendText(Environment.NewLine + "========================================");
            richTextBox1.AppendText(Environment.NewLine + "Выполняется расстановка категорий участков в трубном журнале...(" + mGVTD.MGPipeS.Count + " труб)");
            for (int i = 0; i < mGVTD.pipelineSectionCategoryLogs.Count; i++)//расставляем категории
            {
                string pipeTwo;
                string pipeOne;
                int numStart = 0;
                int numFinish = 0;
                pipeOne = mGVTD.pipelineSectionCategoryLogs[i].pipeNumber;//запоминаем начальную трубу участка с определённой категорией

                try
                {
                    pipeTwo = mGVTD.pipelineSectionCategoryLogs[i + 1].pipeNumber;
                }
                catch (Exception)
                {
                    pipeTwo = mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].pipeNumber;
                }

                for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                {
                    if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, pipeOne))
                    {
                        numStart = j;
                    }
                    if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, pipeTwo))
                    {
                        numFinish = j;
                    }
                }
                for (int k = numStart; k < numFinish; k++)
                {
                    if (String.IsNullOrWhiteSpace(MgvtdNew.MGPipeS[k].pipelineSectionCategory))
                    {
                        mGVTD.MGPipeS[k].pipelineSectionCategory = mGVTD.pipelineSectionCategoryLogs[i].pipelineSectionCategory;
                        //richTextBox2.AppendText(Environment.NewLine + "Труба: " + mGVTD.MGPipeS[k].pipeNumber + ", категория: " + mGVTD.MGPipeS[k].pipelineSectionCategory);
                    }
                }
            }
            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)//если подрядчики не расставили толщину трубы в в журнале аномалий, расставим сами
            {
                if (mGVTD.anomalyLogLineS[i].thikness<1)
                {
                    for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                    {
                        if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                        {
                            mGVTD.anomalyLogLineS[i].thikness = mGVTD.MGPipeS[j].thikness;
                            //richTextBox2.AppendText("+");
                        }
                    }
                }
                else
                {
                    //richTextBox2.AppendText("-");
                }
            }
            richTextBox1.AppendText(Environment.NewLine + "Расстановка категорий участков выполнена");
            richTextBox1.AppendText(Environment.NewLine + "========================================");
            return mGVTD;
        }
        private void tableExcelReadToClass()//метод для чтения файла в класс
        {
            mGVTD.pipelineInfo = operatingReadToClassPipeInfo(fileName, NumbersOfColumns);//данные о трубе
            mGVTD.MGPipeS = operatingReadToClassPipeLog(fileName, NumbersOfColumns);//трубный журнал
            mGVTD.anomalyLogLineS = operatingReadToClassAnomalyLog(fileName, NumbersOfColumns);//журнал аномалий
            mGVTD.furnishingsLogS = operatingReadToClassFurnishingsLog(fileName, NumbersOfColumns);//элементы обустройства
            mGVTD.pipeCharacteristicsLog = operatingReadToClassPipeCharacteristics(fileName, NumbersOfColumns);//Характеристики труб
            mGVTD.pipelineSectionCategoryLogs = operatingReadToClassPipelineSectionCategoryLog(fileName, NumbersOfColumns);//категории участков трубопровода
            mGVTD = PipeLogWithCategory(mGVTD);//расставим в трубном журнале характеристики труб и категории участков
        }
        private void shortTableExcelReadToClass()//КОРОТКИЙ!!! метод для чтения файла в класс
        {
            
            mGVTD.pipelineInfo = operatingReadToClassPipeInfo(fileName, NumbersOfColumns);//данные о трубе
            mGVTD.MGPipeS = shortOperatingReadToClassPipeLogAutoFin(fileName, NumbersOfColumns);//трубный журнал
            mGVTD.anomalyLogLineS = shortOperatingReadToClassAnomalyLogAutoFin(fileName, NumbersOfColumns);//журнал аномалий
            mGVTD.furnishingsLogS = operatingReadToClassFurnishingsLogAutoFin(fileName, NumbersOfColumns);//элементы обустройства
            mGVTD.pipeCharacteristicsLog = operatingReadToClassPipeCharacteristics(fileName, NumbersOfColumns);//Характеристики труб
            mGVTD.pipelineSectionCategoryLogs = operatingReadToClassPipelineSectionCategoryLog(fileName, NumbersOfColumns);//категории участков трубопровода

            if (isHaveCategory.Checked==false)
            {
                mGVTD = PipeLogWithCategory(mGVTD);//расставим в трубном журнале характеристики труб и категории участков
            }
            
        }
        private void button3_Click(object sender, EventArgs e)//проверка правильности адресации ячеек
        {
            if (fileName != null)
            {
                if (isHaveCategory.Checked == false)
                {
                    findStart();
                }
                tableArdesTest();
            }

        }
        private double damagFromСorrosion(MGVTD mGVTD)//вычисляем повреждённость локального участка от коррозии (ф. 5.3 СТО 292)
        {
            double result = 0;//переменная для хранения искомой величины
            double Summdkt = 0;//(числитель ф. 5.3 СТО 292)
            int corrosionPipesCount = 0;//создаём переменную для хранения количества труб с коррозией
            double localThikness = 0;



            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)
            {
                
                if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectVanishDate))
                {
                    if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("орроз"))//для всех труб с коррозией вычисляем ранг опасности и складываем, как того требует п. 6.1.2 СТО 292
                    {
                        double tensileStrength = 510;//ищем по трубному журналу предел прочности
                        for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                        {
                            if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                            {
                                tensileStrength = mGVTD.MGPipeS[j].tensileStrength;//вот он предел прочности, нашли.
                            }
                            else
                            {
                                tensileStrength = 500;
                            }
                        }

                        corrosionPipesCount++;//инкрементируем счетчик труб с коррозией

                        localThikness = mGVTD.anomalyLogLineS[i].thikness;
                        double Q = Math.Sqrt(1 + 0.31 * Math.Pow((mGVTD.anomalyLogLineS[i].length / (Math.Sqrt(mGVTD.pipelineInfo.pipeDiameter * localThikness))), 2));//коэффициент, учитывающий длину дефекта потери металла (ф. 6.3 СТО 292)

                        double a = (mGVTD.pipelineInfo.operatingPressure * (mGVTD.pipelineInfo.pipeDiameter - localThikness)) / (2 * localThikness * tensileStrength);//(ф. 6.4 СТО 292)
                        //richTextBox2.AppendText(Environment.NewLine + "localThikness=" + localThikness + "Q=" + Q + ",  a=" + a);
                        double ksiP = ((a - 1) * Q) / (a - Q); //(ф. 6.2 СТО 292)
                        double ksi = mGVTD.anomalyLogLineS[i].depthInMm / localThikness;
                        double Rk = ksi / ksiP;//(ф. 6.1 СТО 292)
                        Summdkt = Summdkt + Rk;//расчет суммы рангов опасности для всех дефектов данного типа
                    }
                }
            }
            result = Summdkt / mGVTD.MGPipeS.Count;
            return result;
        }

        private double damagFromСorrosionProcent(MGVTD mGVTD, plotBoundaries PlotBoundaries, double procentOfCorrosion)//26/07/2021/вычисляем повреждённость локального участка от коррозии (ф. 5.3 СТО 292)
        {
            double result = 0;//переменная для хранения искомой величины
            double Summdkt = 0;//(числитель ф. 5.3 СТО 292)
            int corrosionPipesCount = 0;//создаём переменную для хранения количества труб с коррозией
            double localThikness = 0;


            for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo; i++)
            {

                if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectVanishDate))
                {
                    if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("орроз"))//для всех труб с коррозией вычисляем ранг опасности и складываем, как того требует п. 6.1.2 СТО 292
                    {
                        double tensileStrength = 510;//ищем по трубному журналу предел прочности
                        for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                        {
                            if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                            {
                                tensileStrength = mGVTD.MGPipeS[j].tensileStrength;//вот он предел прочности, нашли.
                            }
                            else
                            {
                                tensileStrength = 500;
                            }
                        }

                        corrosionPipesCount++;//инкрементируем счетчик труб с коррозией
                        localThikness = mGVTD.anomalyLogLineS[i].thikness;
                        if (tensileStrength > 0)
                        {
                            if (localThikness>0)
                            {
                                if (mGVTD.anomalyLogLineS[i].depthInProcent>= procentOfCorrosion)
                                {
                                    double Q = Math.Sqrt(1 + 0.31 * Math.Pow((mGVTD.anomalyLogLineS[i].length / (Math.Sqrt(mGVTD.pipelineInfo.pipeDiameter * localThikness))), 2));//коэффициент, учитывающий длину дефекта потери металла (ф. 6.3 СТО 292)

                                    double a = (mGVTD.pipelineInfo.operatingPressure * (mGVTD.pipelineInfo.pipeDiameter - localThikness)) / (2 * localThikness * tensileStrength);//(ф. 6.4 СТО 292)
                                                                                                                                                                                  //richTextBox2.AppendText(Environment.NewLine + "localThikness=" + localThikness + "Q=" + Q + ",  a=" + a);
                                    double ksiP = ((a - 1) * Q) / (a - Q); //(ф. 6.2 СТО 292)
                                    double ksi = mGVTD.anomalyLogLineS[i].depthInMm / localThikness;
                                    double Rk = ksi / ksiP;//(ф. 6.1 СТО 292)
                                    Summdkt = Summdkt + Rk;//расчет суммы рангов опасности для всех дефектов данного типа
                                }                                
                            }
                        }                        
                    }
                }
            }
            result = Summdkt;
            return result;
        }


        private MGVTD isLostMetal(MGVTD mGVTD)
        {
            
            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)
            {
                if (String.IsNullOrEmpty(mGVTD.anomalyLogLineS[i].defectVanishDate))
                {
                    if (mGVTD.anomalyLogLineS[i].depthInMm > 0)
                    {
                        if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("орроз"))
                        {
                            for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                            {
                                if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                                {
                                    mGVTD.anomalyLogLineS[i].isLostMetal = true;
                                }
                            }
                        }
                        else if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("ехноло"))
                        {
                            for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                            {
                                if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                                {
                                    mGVTD.anomalyLogLineS[i].isLostMetal = true;
                                }
                            }
                        }
                        else if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("аводс"))
                        {
                            for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                            {
                                if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                                {
                                    mGVTD.anomalyLogLineS[i].isLostMetal = true;
                                }
                            }
                        }
                        else if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("оврежд"))
                        {
                            for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                            {
                                if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                                {
                                    mGVTD.anomalyLogLineS[i].isLostMetal = true;
                                }
                            }
                        }
                    }

                }
            }
            return mGVTD;
        }
        private MGVTD damagFromСorrosion(MGVTD mGVTD, plotBoundaries PlotBoundaries)//!!!В ЗАДАННЫХ ГРАНИЦАХ УЧАСТКА!!вычисляем повреждённость локального участка от коррозии (ф. 5.3 СТО 292)
        {
            //double result = 0;//переменная для хранения искомой величины
            double Summdkt = 0;//(числитель ф. 5.3 СТО 292)
            //int corrosionPipesCount = 0;//создаём переменную для хранения количества труб с коррозией
            double localThikness = 0;
            if (String.IsNullOrEmpty(PlotBoundaries.pipeNumberOne))
            {
                if (String.IsNullOrWhiteSpace(PlotBoundaries.pipeNumberTwo))//если есть первая и последняя труба с дефектами в границах заданного участка
                {
                    //если оба значения адресов труб пусты
                    Summdkt = 0;//если дефектов нет, то и считать ничего не надо.
                    allPipeWhithСorrosion = 0;
                }
                else
                {
                    //если первое значение пустое а второе не пустое
                    //такого не может быть, логика работы программы этого не допустит
                   
                }
            }
            else
            {
                if (String.IsNullOrWhiteSpace(PlotBoundaries.pipeNumberTwo))//если есть первая и последняя труба с дефектами в границах заданного участка
                {
                    //если первое значение не пустое а второе пустое
                    //тогда расчет проводим для одной единственной трубы
                    
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].defectVanishDate))
                        {
                            if (mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].isLostMetal)//для всех труб с коррозией вычисляем ранг опасности и складываем, как того требует п. 6.1.2 СТО 292
                            {
                                double tensileStrength = 510;//ищем по трубному журналу предел прочности
                                for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                                {
                                    if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].pipeNumber))
                                    {
                                        tensileStrength = mGVTD.MGPipeS[j].tensileStrength;//вот он предел прочности, нашли.
                                    }
                                    else
                                    {
                                        tensileStrength = 500;
                                    }
                                }

                                allPipeWhithСorrosion++;//инкрементируем счетчик труб с коррозией

                                localThikness = mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].thikness;
                                double Q = Math.Sqrt(1 + 0.31 * Math.Pow((mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].length / (Math.Sqrt(mGVTD.pipelineInfo.pipeDiameter * localThikness))), 2));//коэффициент, учитывающий длину дефекта потери металла (ф. 6.3 СТО 292)

                                double a = (mGVTD.pipelineInfo.operatingPressure * (mGVTD.pipelineInfo.pipeDiameter - localThikness)) / (2 * localThikness * tensileStrength);//(ф. 6.4 СТО 292)
                                                                                                                                                                              //richTextBox2.AppendText(Environment.NewLine + "localThikness=" + localThikness + "Q=" + Q + ",  a=" + a);
                                double ksiP = ((a - 1) * Q) / (a - Q); //(ф. 6.2 СТО 292)
                                double ksi = mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].depthInMm / localThikness;
                                double Rk = ksi / ksiP;//(ф. 6.1 СТО 292)
                            

                            mGVTD.MGPipeS[PlotBoundaries.pipeIdNumberOne].corossionDamageList.Add(Rk);


                            //Summdkt = Rk;//расчет суммы рангов опасности для всех дефектов данного типа
                        }
                        }
                    //

                }
                else
                {

                    bool mark;
                    List<string> defectpipes = new List<string>();//это просто список учтенных труб
                    for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo; i++)
                    {
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectVanishDate))
                        {
                            if (mGVTD.anomalyLogLineS[i].isLostMetal)//для всех труб с коррозией вычисляем ранг опасности и складываем, как того требует п. 6.1.2 СТО 292
                            {

                                mark = true;
                                for (int q = 0; q < defectpipes.Count; q++)//проверяем, не учли ли мы уже эту трубу
                                {
                                    if (String.Equals(defectpipes[q], mGVTD.anomalyLogLineS[i].pipeNumber))
                                    {
                                        mark = false;
                                    }
                                    
                                }
                                if (mark)
                                {
                                    defectpipes.Add(mGVTD.anomalyLogLineS[i].pipeNumber);
                                    allPipeWhithСorrosion++;//инкрементируем счетчик труб с коррозией
                                }


                                double tensileStrength = 510;//ищем по трубному журналу предел прочности
                                for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                                {
                                    if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                                    {
                                        tensileStrength = mGVTD.MGPipeS[j].tensileStrength;//вот он предел прочности, нашли.
                                    }
                                }

                                try
                                {
                                    localThikness = 1 * mGVTD.anomalyLogLineS[i].thikness;
                                    if (localThikness > 0)
                                    {
                                        if (tensileStrength > 0)
                                        {
                                            //richTextBox2.AppendText(Environment.NewLine + "localThikness " + localThikness);
                                            double Q = Math.Sqrt(1 + 0.31 * Math.Pow((mGVTD.anomalyLogLineS[i].length / (Math.Sqrt(mGVTD.pipelineInfo.pipeDiameter * localThikness))), 2));//коэффициент, учитывающий длину дефекта потери металла (ф. 6.3 СТО 292)
                                            //richTextBox2.AppendText(Environment.NewLine + "Q " + Q);
                                            double a = (mGVTD.pipelineInfo.operatingPressure * (mGVTD.pipelineInfo.pipeDiameter - localThikness)) / (2 * localThikness * tensileStrength);//(ф. 6.4 СТО 292)
                                            //richTextBox2.AppendText(Environment.NewLine + "a " + a);                                                                                                                                //richTextBox2.AppendText(Environment.NewLine + "localThikness=" + localThikness + "Q=" + Q + ",  a=" + a);
                                            double ksiP = ((a - 1) * Q) / (a - Q); //(ф. 6.2 СТО 292)
                                            //richTextBox2.AppendText(Environment.NewLine + "ksiP " + ksiP);
                                            double ksi = mGVTD.anomalyLogLineS[i].depthInMm / localThikness;
                                            //richTextBox2.AppendText(Environment.NewLine + "ksi " + ksi);
                                            double Rk = ksi / ksiP;//(ф. 6.1 СТО 292)
                                            //richTextBox2.AppendText(Environment.NewLine + "Rk " + Rk);
                                            //Summdkt = Summdkt + Rk;//расчет суммы рангов опасности для всех дефектов данного типа
                                            for (int f = 0; f < mGVTD.MGPipeS.Count; f++)
                                            {
                                                if (String.Equals(mGVTD.MGPipeS[f].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                                                {
                                                    mGVTD.MGPipeS[f].corossionDamageList.Add(Rk);
                                                }
                                            }
                                            //richTextBox2.AppendText(Environment.NewLine + "Summdkt " + Summdkt);
                                            //richTextBox2.AppendText(Environment.NewLine + "==========================================================");
                                        }
                                        else
                                        {
                                            tensileStrength = 500;
                                            //richTextBox2.AppendText(Environment.NewLine + "localThikness " + localThikness);
                                            double Q = Math.Sqrt(1 + 0.31 * Math.Pow((mGVTD.anomalyLogLineS[i].length / (Math.Sqrt(mGVTD.pipelineInfo.pipeDiameter * localThikness))), 2));//коэффициент, учитывающий длину дефекта потери металла (ф. 6.3 СТО 292)
                                            //richTextBox2.AppendText(Environment.NewLine + "Q " + Q);
                                            double a = (mGVTD.pipelineInfo.operatingPressure * (mGVTD.pipelineInfo.pipeDiameter - localThikness)) / (2 * localThikness * tensileStrength);//(ф. 6.4 СТО 292)
                                            //richTextBox2.AppendText(Environment.NewLine + "a " + a);                                                                                                                                //richTextBox2.AppendText(Environment.NewLine + "localThikness=" + localThikness + "Q=" + Q + ",  a=" + a);
                                            double ksiP = ((a - 1) * Q) / (a - Q); //(ф. 6.2 СТО 292)
                                            //richTextBox2.AppendText(Environment.NewLine + "ksiP " + ksiP);
                                            double ksi = mGVTD.anomalyLogLineS[i].depthInMm / localThikness;
                                            //richTextBox2.AppendText(Environment.NewLine + "ksi " + ksi);
                                            double Rk = ksi / ksiP;//(ф. 6.1 СТО 292)
                                            //richTextBox2.AppendText(Environment.NewLine + "Rk " + Rk);
                                            //Summdkt = Summdkt + Rk;//расчет суммы рангов опасности для всех дефектов данного типа
                                            for (int f = 0; f < mGVTD.MGPipeS.Count; f++)
                                            {
                                                if (String.Equals(mGVTD.MGPipeS[f].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                                                {
                                                    mGVTD.MGPipeS[f].corossionDamageList.Add(Rk);
                                                }
                                            }
                                            //richTextBox2.AppendText(Environment.NewLine + "Summdkt " + Summdkt);
                                            //richTextBox2.AppendText(Environment.NewLine + "==========================================================");
                                        }
                                    }
                                    else
                                    {
                                        //richTextBox2.AppendText(Environment.NewLine + "Номер трубы " + mGVTD.anomalyLogLineS[i].pipeNumber + "localThikness<=0");
                                    }
                                }
                                catch (Exception)
                                {
                                                                        
                                }


                                
                            }
                        }
                    }

                }

                
            }
            summCorrosionDamag = Summdkt;//запоминаем суммарную поврежденность от коррозии
            //result = Summdkt / (PlotBoundaries.pipeIdNumberTwoPipeLog - PlotBoundaries.pipeIdNumberOnePipeLog);
                    return mGVTD;

        }
        

        //возвращает суммарную поврежденность по 1 дефекту на трубу
        private double damagFromСorrosion(MGVTD mGVTD, plotBoundaries PlotBoundaries, double procentOfCorrosion)//по одному дефекту на трубу//не менее заданного уровня в процентах!!!В ЗАДАННЫХ ГРАНИЦАХ УЧАСТКА!!вычисляем повреждённость локального участка от коррозии (ф. 5.3 СТО 292)
        {
            double result = 0;//переменная для хранения искомой величины            
            double Summdkt = 0;//(числитель ф. 5.3 СТО 292)
            //int corrosionPipesCount = 0;//создаём переменную для хранения количества труб с коррозией
            double localThikness = 0;
            if (String.IsNullOrEmpty(PlotBoundaries.pipeNumberOne))
            {
                if (String.IsNullOrWhiteSpace(PlotBoundaries.pipeNumberTwo))//если есть первая и последняя труба с дефектами в границах заданного участка
                {
                    //если оба значения адресов труб пусты
                    Summdkt = 0;//если дефектов нет, то и считать ничего не надо.
                    allPipeWhithСorrosionPlus = 0;
                }
                else
                {
                    //если первое значение пустое а второе не пустое
                    //такого не может быть, логика работы программы этого не допустит

                }
            }
            else
            {
                if (String.IsNullOrWhiteSpace(PlotBoundaries.pipeNumberTwo))//если есть первая и последняя труба с дефектами в границах заданного участка
                {
                    //если первое значение не пустое а второе пустое
                    //тогда расчет проводим для одной единственной трубы
                    bool mark;
                    List<string> defectpipes = new List<string>();//это просто список учтенных труб
                    if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].defectVanishDate))//проверяем, что нет пометки об устранении дефекта
                    {
                        if (mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].depthInProcent>= procentOfCorrosion)//проверяем, что дефект глубже заданного уровня
                        {
                            if (mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].isLostMetal)//для всех труб с коррозией вычисляем ранг опасности и складываем, как того требует п. 6.1.2 СТО 292
                            {
                                double tensileStrength=500;//ищем по трубному журналу предел прочности
                                for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                                {
                                    if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].pipeNumber))
                                    {
                                        tensileStrength = mGVTD.MGPipeS[j].tensileStrength;//вот он предел прочности, нашли.
                                    }  
                                }

                                mark = true;
                                for (int q = 0; q < defectpipes.Count; q++)//проверяем, не учли ли мы уже эту трубу
                                {
                                    if (String.Equals(defectpipes[q], mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].pipeNumber))
                                    {
                                        mark = false;
                                    }
                                }
                                if (mark)
                                {
                                    defectpipes.Add(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].pipeNumber);
                                    allPipeWhithСorrosionPlus++;//инкрементируем счетчик труб с коррозией
                                }
                                localThikness = mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].thikness;
                                double Q = Math.Sqrt(1 + 0.31 * Math.Pow((mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].length / (Math.Sqrt(mGVTD.pipelineInfo.pipeDiameter * localThikness))), 2));//коэффициент, учитывающий длину дефекта потери металла (ф. 6.3 СТО 292)

                                double a = (mGVTD.pipelineInfo.operatingPressure * (mGVTD.pipelineInfo.pipeDiameter - localThikness)) / (2 * localThikness * tensileStrength);//(ф. 6.4 СТО 292)
                                                                                                                                                                              //richTextBox2.AppendText(Environment.NewLine + "localThikness=" + localThikness + "Q=" + Q + ",  a=" + a);
                                double ksiP = ((a - 1) * Q) / (a - Q); //(ф. 6.2 СТО 292)
                                double ksi = mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].depthInMm / localThikness;
                                double Rk = ksi / ksiP;//(ф. 6.1 СТО 292)
                                Summdkt = Rk;//расчет суммы рангов опасности для всех дефектов данного типа
                            }
                        }
                    } 
                }
                else
                {
                    bool mark;
                    List<string> defectpipes = new List<string>();//это просто список учтенных труб
                    for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo; i++)
                    {
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectVanishDate))
                        {
                            if (mGVTD.anomalyLogLineS[i].depthInProcent >= procentOfCorrosion)
                            {
                                if (mGVTD.anomalyLogLineS[i].isLostMetal)//для всех труб с коррозией вычисляем ранг опасности и складываем, как того требует п. 6.1.2 СТО 292
                                {
                                    mark = true;
                                    for (int q = 0; q < defectpipes.Count; q++)//проверяем, не учли ли мы уже эту трубу
                                    {
                                        if (String.Equals(defectpipes[q], mGVTD.anomalyLogLineS[i].pipeNumber))
                                        {
                                            mark = false;
                                        }
                                    }
                                    if (mark)
                                    {
                                        defectpipes.Add(mGVTD.anomalyLogLineS[i].pipeNumber);
                                        allPipeWhithСorrosionPlus++;//инкрементируем счетчик труб с коррозией
                                    }
                                    double tensileStrength = 510;//ищем по трубному журналу предел прочности
                                    for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                                    {
                                        if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                                        {
                                            tensileStrength = mGVTD.MGPipeS[j].tensileStrength;//вот он предел прочности, нашли.
                                        }

                                    }

                                    try
                                    {
                                        localThikness = mGVTD.anomalyLogLineS[i].thikness;
                                        if (localThikness > 0)
                                        {
                                            if (tensileStrength > 0)
                                            {
                                                //richTextBox2.AppendText(Environment.NewLine + "localThikness " + localThikness);
                                                double Q = Math.Sqrt(1 + 0.31 * Math.Pow((mGVTD.anomalyLogLineS[i].length / (Math.Sqrt(mGVTD.pipelineInfo.pipeDiameter * localThikness))), 2));//коэффициент, учитывающий длину дефекта потери металла (ф. 6.3 СТО 292)
                                                double a = (mGVTD.pipelineInfo.operatingPressure * (mGVTD.pipelineInfo.pipeDiameter - localThikness)) / (2 * localThikness * tensileStrength);//(ф. 6.4 СТО 292)                                                                                                                             //richTextBox2.AppendText(Environment.NewLine + "localThikness=" + localThikness + "Q=" + Q + ",  a=" + a);
                                                double ksiP = ((a - 1) * Q) / (a - Q); //(ф. 6.2 СТО 292)
                                                double ksi = mGVTD.anomalyLogLineS[i].depthInMm / localThikness;
                                                double Rk = ksi / ksiP;//(ф. 6.1 СТО 292)                                              
                                                Summdkt = Summdkt + Rk;//расчет суммы рангов опасности для всех дефектов данного типа                                                                      
                                            }
                                            else
                                            {
                                                tensileStrength = 500;
                                                double Q = Math.Sqrt(1 + 0.31 * Math.Pow((mGVTD.anomalyLogLineS[i].length / (Math.Sqrt(mGVTD.pipelineInfo.pipeDiameter * localThikness))), 2));//коэффициент, учитывающий длину дефекта потери металла (ф. 6.3 СТО 292)
                                                double a = (mGVTD.pipelineInfo.operatingPressure * (mGVTD.pipelineInfo.pipeDiameter - localThikness)) / (2 * localThikness * tensileStrength);//(ф. 6.4 СТО 292)                                                                                                                             //richTextBox2.AppendText(Environment.NewLine + "localThikness=" + localThikness + "Q=" + Q + ",  a=" + a);
                                                double ksiP = ((a - 1) * Q) / (a - Q); //(ф. 6.2 СТО 292)
                                                double ksi = mGVTD.anomalyLogLineS[i].depthInMm / localThikness;
                                                double Rk = ksi / ksiP;//(ф. 6.1 СТО 292)                                              
                                                Summdkt = Summdkt + Rk;//расчет суммы рангов опасности для всех дефектов данного типа        
                                                //richTextBox2.AppendText(Environment.NewLine + "Номер трубы " + mGVTD.anomalyLogLineS[i].pipeNumber + " tensileStrength<=0");
                                            }
                                        }
                                        else
                                        {
                                            //richTextBox2.AppendText(Environment.NewLine + "Номер трубы " + mGVTD.anomalyLogLineS[i].pipeNumber + "localThikness<=0");
                                        }
                                    }
                                    catch (Exception)
                                    {

                                    }
                                }
                            }
                        }
                    }
                }
            }
            summCorrosionDamag = Summdkt;//запоминаем суммарную поврежденность от коррозии
            result = Summdkt / (PlotBoundaries.pipeIdNumberTwoPipeLog - PlotBoundaries.pipeIdNumberOnePipeLog);
            return Summdkt;
        }
        
        
        //Возвращает количество всех коррозионных дефектов больше заданного на данном участке
        private int damagFromСorrosionAllDefects(MGVTD mGVTD, plotBoundaries PlotBoundaries, double procentOfCorrosion)//все коррозионные дефекты//не менее заданного уровня в процентах!!!В ЗАДАННЫХ ГРАНИЦАХ УЧАСТКА!!вычисляем повреждённость локального участка от коррозии (ф. 5.3 СТО 292)
        {
            //double result = 0;//переменная для хранения искомой величины
            int SummOfDefects = 0;
            //int LocalsummCorrosionDamag = 0;
            double Summdkt = 0;//(числитель ф. 5.3 СТО 292)
            //int corrosionPipesCount = 0;//создаём переменную для хранения количества труб с коррозией
            double localThikness = 0;
            if (String.IsNullOrEmpty(PlotBoundaries.pipeNumberOne))
            {
                if (String.IsNullOrWhiteSpace(PlotBoundaries.pipeNumberTwo))//если есть первая и последняя труба с дефектами в границах заданного участка
                {
                    //если оба значения адресов труб пусты
                    Summdkt = 0;//если дефектов нет, то и считать ничего не надо.
                    allPipeWhithСorrosionPlus = 0;
                }
                else
                {
                    //если первое значение пустое а второе не пустое
                    //такого не может быть, логика работы программы этого не допустит

                }
            }
            else
            {
                if (String.IsNullOrWhiteSpace(PlotBoundaries.pipeNumberTwo))//если есть первая и последняя труба с дефектами в границах заданного участка
                {
                    //если первое значение не пустое а второе пустое
                    //тогда расчет проводим для одной единственной трубы

                    if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].defectVanishDate))
                    {
                        if (mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].depthInProcent >= procentOfCorrosion)
                        {
                            if (mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].featuresCharacter.Contains("орроз"))//для всех труб с коррозией вычисляем ранг опасности и складываем, как того требует п. 6.1.2 СТО 292
                            {
                                double tensileStrength = 510;//ищем по трубному журналу предел прочности
                                for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                                {
                                    if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].pipeNumber))
                                    {
                                        tensileStrength = mGVTD.MGPipeS[j].tensileStrength;//вот он предел прочности, нашли.
                                    }
                                }
                                SummOfDefects++;
                                allPipeWhithСorrosionPlus++;//инкрементируем счетчик труб с коррозией

                                localThikness = mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].thikness;
                                double Q = Math.Sqrt(1 + 0.31 * Math.Pow((mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].length / (Math.Sqrt(mGVTD.pipelineInfo.pipeDiameter * localThikness))), 2));//коэффициент, учитывающий длину дефекта потери металла (ф. 6.3 СТО 292)

                                double a = (mGVTD.pipelineInfo.operatingPressure * (mGVTD.pipelineInfo.pipeDiameter - localThikness)) / (2 * localThikness * tensileStrength);//(ф. 6.4 СТО 292)
                                                                                                                                                                              //richTextBox2.AppendText(Environment.NewLine + "localThikness=" + localThikness + "Q=" + Q + ",  a=" + a);
                                double ksiP = ((a - 1) * Q) / (a - Q); //(ф. 6.2 СТО 292)
                                double ksi = mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].depthInMm / localThikness;
                                double Rk = ksi / ksiP;//(ф. 6.1 СТО 292)
                                Summdkt = Rk;//расчет суммы рангов опасности для всех дефектов данного типа
                            }
                        }
                    }
                }
                else
                {
                    //bool mark;
                    //List<string> defectpipes = new List<string>();//это просто список учтенных труб
                    for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo; i++)
                    {
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectVanishDate))
                        {
                            if (mGVTD.anomalyLogLineS[i].depthInProcent >= procentOfCorrosion)
                            {
                                if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("орроз"))//для всех труб с коррозией вычисляем ранг опасности и складываем, как того требует п. 6.1.2 СТО 292
                                {

                                    SummOfDefects++;
                                    allDefectsWhithСorrosionPlus++;//инкрементируем счетчик коррозионных дефектов
                                    double tensileStrength = 510;//ищем по трубному журналу предел прочности
                                    for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                                    {
                                        if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                                        {
                                            tensileStrength = mGVTD.MGPipeS[j].tensileStrength;//вот он предел прочности, нашли.
                                        }
                                    }
                                    localThikness = mGVTD.anomalyLogLineS[i].thikness;
                                    double Q = Math.Sqrt(1 + 0.31 * Math.Pow((mGVTD.anomalyLogLineS[i].length / (Math.Sqrt(mGVTD.pipelineInfo.pipeDiameter * localThikness))), 2));//коэффициент, учитывающий длину дефекта потери металла (ф. 6.3 СТО 292)

                                    double a = (mGVTD.pipelineInfo.operatingPressure * (mGVTD.pipelineInfo.pipeDiameter - localThikness)) / (2 * localThikness * tensileStrength);//(ф. 6.4 СТО 292)
                                                                                                                                                                                  //richTextBox2.AppendText(Environment.NewLine + "localThikness=" + localThikness + "Q=" + Q + ",  a=" + a);
                                    double ksiP = ((a - 1) * Q) / (a - Q); //(ф. 6.2 СТО 292)
                                    double ksi = mGVTD.anomalyLogLineS[i].depthInMm / localThikness;
                                    double Rk = ksi / ksiP;//(ф. 6.1 СТО 292)
                                    Summdkt = Summdkt + Rk;//расчет суммы рангов опасности для всех дефектов данного типа
                                }
                            }
                        }
                    }
                }
            }
            //LocalsummCorrosionDamag = Summdkt;//запоминаем суммарную поврежденность от коррозии
            //result = Summdkt / (PlotBoundaries.pipeIdNumberTwoPipeLog - PlotBoundaries.pipeIdNumberOnePipeLog);
            return SummOfDefects;//возвращаем количество дефектов
        }

        //Возвращает поврежденность всех коррозионных дефектов больше заданного на данном участке
        

        private double damagOfconnectingParts(MGVTD mGVTD)//вычисляем повреждённость соединительных деталей трубопровода (ф. 5.9 СТО 292)
        {
            double result = 0;//переменная для хранения искомой величины            
            int PipesCount = 0;//создаём переменную для хранения количества СДТ с дефектами            

            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)
            {
                if (mGVTD.anomalyLogLineS[i].featuresCharacter != "")//ищем строку с каким-нибудь дефектом
                {
                    if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectVanishDate))
                    {
                        for (int k = 0; k < mGVTD.MGPipeS.Count; k++)
                        {
                            if (String.Equals(mGVTD.MGPipeS[k].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))//ищем дефектную трубу в трубном журнале
                            {
                                if (mGVTD.MGPipeS[k].note.Contains("ройник"))//проверяем является ли дефектная труба тройником
                                {
                                    bool mark;
                                    List<string> defectpipes = new List<string>();//это просто список учтенных труб
                                    mark = true;
                                    for (int q = 0; q < defectpipes.Count; q++)//проверяем, не учли ли мы уже эту трубу
                                    {
                                        if (String.Equals(defectpipes[q], mGVTD.anomalyLogLineS[i].pipeNumber))
                                        {
                                            mark = false;
                                        }
                                    }
                                    if (mark)
                                    {
                                        defectpipes.Add(mGVTD.anomalyLogLineS[i].pipeNumber);
                                        PipesCount++;//инкрементируем счетчик тройников с дефектами
                                    }

                                }
                            }
                        }
                    }
                }
            }

            try
            {
                result = PipesCount / mGVTD.MGPipeS.Count;
            }
            catch (Exception)
            {
                result = 0;
            }


            return result;
        }
        private double damagOfconnectingParts(MGVTD mGVTD, plotBoundaries PlotBoundaries)//!!!В ЗАДАННЫХ ГРАНИЦАХ УЧАСТКА!!вычисляем повреждённость соединительных деталей трубопровода (ф. 5.9 СТО 292)
        {
            double result = 0;//переменная для хранения искомой величины            
            int PipesCount = 0;//создаём переменную для хранения количества СДТ с дефектами            

            if (String.IsNullOrEmpty(PlotBoundaries.pipeNumberOne))
            {
                if (String.IsNullOrWhiteSpace(PlotBoundaries.pipeNumberTwo))//если есть первая и последняя труба с дефектами в границах заданного участка
                {
                    //если оба значения адресов труб пусты
                    PipesCount = 0;//если дефектов нет, то и считать ничего не надо.
                }
                else
                {
                    //если первое значение пустое а второе не пустое
                    //такого не может быть, логика работы программы этого не допустит

                }
            }
            else
            {
                if (String.IsNullOrWhiteSpace(PlotBoundaries.pipeNumberTwo))//если есть первая и последняя труба с дефектами в границах заданного участка
                {
                    //если первое значение не пустое а второе пустое
                    //тогда расчет проводим для одной единственной трубы
                   
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].featuresCharacter)==false)//ищем строку с каким-нибудь дефектом
                        {
                            if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].defectVanishDate))//проверяем, что поле со сведениями об устранении дефекта пустое
                            {
                                for (int k = 0; k < mGVTD.MGPipeS.Count; k++)
                                {
                                    if (String.Equals(mGVTD.MGPipeS[k].pipeNumber, mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].pipeNumber))//ищем дефектную трубу в трубном журнале
                                    {
                                        if (mGVTD.MGPipeS[k].itIsTee)//проверяем является ли дефектная труба тройником
                                        {

                                        PipesCount = 1;

                                        }
                                    }
                                }
                            }
                        }
                    

                    try
                    {
                        result = PipesCount;
                    }
                    catch (Exception)
                    {
                        result = 0;
                    }
                }
                else
                {
                    bool mark;
                    List<string> defectpipes = new List<string>();//это просто список учтенных труб
                    for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo-1; i++)
                    {
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].featuresCharacter)==false)//ищем строку с каким-нибудь дефектом
                        {
                            if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectVanishDate))
                            {
                                for (int k = 0; k < mGVTD.MGPipeS.Count; k++)
                                {
                                    if (String.Equals(mGVTD.MGPipeS[k].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))//ищем дефектную трубу в трубном журнале
                                    {
                                        if (mGVTD.MGPipeS[k].itIsTee)//проверяем является ли дефектная труба тройником
                                        {

                                            mark = true;
                                            for (int q = 0; q < defectpipes.Count; q++)//проверяем, не учли ли мы уже эту трубу
                                            {
                                                if (String.Equals(defectpipes[q], mGVTD.anomalyLogLineS[i].pipeNumber))
                                                {
                                                    mark = false;
                                                }
                                            }
                                            if (mark)
                                            {
                                                defectpipes.Add(mGVTD.anomalyLogLineS[i].pipeNumber);
                                                PipesCount++;//инкрементируем счетчик тройников с дефектами
                                            }

                                        }
                                    }
                                }
                            }
                        }
                    }

                    try
                    {
                        result = PipesCount / (PlotBoundaries.pipeIdNumberTwoPipeLog - PlotBoundaries.pipeIdNumberOnePipeLog);
                    }
                    catch (Exception)
                    {
                        result = 0;
                    }
                }
            }


                    //*****************
           
            //********************


            return result;
        }       
        private double damageFromDent(MGVTD mGVTD)//вычисляем повреждённость участка МГ от вмятин и гофр (ф. 5.8 СТО 292)
        {
            double result;
            double summOfRangs = 0;
            int defectPipesCount = 0;//создаём переменную для хранения количества труб с вмятинами
            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)
            {
                if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("мятин"))
                {
                    if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectVanishDate))
                    {
                        double tensileStrength = 510;//ищем по трубному журналу предел прочности
                        int pipelineSectionCategory = 1;//ищем по трубному журналу предел прочности
                        for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                        {
                            if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                            {
                                tensileStrength = mGVTD.MGPipeS[j].tensileStrength;//вот он предел прочности, нашли.
                                try
                                {
                                    pipelineSectionCategory = Convert.ToInt32(mGVTD.MGPipeS[j].pipelineSectionCategory);//и категория участка тоже нашлась
                                }
                                catch (Exception)
                                {
                                    pipelineSectionCategory = 1;
                                }
                            }
                        }
                        defectPipesCount++;//инкриментируем счетчик труб с вмятинами
                                           //тут будут проходить вычисления
                        double r = (mGVTD.pipelineInfo.pipeDiameter - mGVTD.anomalyLogLineS[i].thikness) / 2;//радиус средней линии сечения (ф 7.14 Рекомендаций по оценке прочности...)
                        double a = mGVTD.anomalyLogLineS[i].length / 2;//полоовина протяженности вмятины в осевом направлении
                        double b = mGVTD.anomalyLogLineS[i].widht / 2;//половина протяженности дефекта в направлении кривой сечения
                        double u = (r * Math.PI) / (2 * b);//коэффициент (ф 7.16 Рекомендаций по оценке прочности...)
                        double t = (r * Math.PI) / (2 * a);//коэффициент (ф 7.15 Рекомендаций по оценке прочности...)
                                                           // В этом расчете пока будем считать, что при проведении ВТД глубина вмятин измерялась без избыточного давление в газопроводе, и пересчет на трубу без бавления будет опущен.
                        double e2 = 0.5 * (mGVTD.anomalyLogLineS[i].thikness / r) * (mGVTD.anomalyLogLineS[i].depthInMm / r) * (3 * u * u - 1);//остаточные окружные деформации (ф 7.12 Рекомендаций по оценке прочности...)
                        double e1 = 0.5 * (mGVTD.anomalyLogLineS[i].thikness / r) * (mGVTD.anomalyLogLineS[i].depthInMm / r) * (3 * t * t - 1);//остаточные продольные деформации (ф 7.13 Рекомендаций по оценке прочности...)
                        double w000 = mGVTD.anomalyLogLineS[i].depthInMm / mGVTD.pipelineInfo.pipeDiameter;//относиттельная глубина дефекта (ф 7.21 Рекомендаций по оценке прочности...)
                        double kr = 24;//определяем кооэффициент, зависящий от категории участка трубопровода
                        if (pipelineSectionCategory < 3)
                        {
                            kr = 24;
                        }
                        else
                        {
                            kr = 20;
                        }
                        double Rr = kr * Math.Max(e2, Math.Max(w000, e1));//это ранг опасности вмятины
                        summOfRangs = summOfRangs + Rr;//считаем сумму рангов опасности дефектов на трубопроводе
                    }
                }
            }
            result = summOfRangs / mGVTD.MGPipeS.Count;//повреждённость линейного участка МГ от вмятин (ф. 5.8 СТО 292)
            richTextBox1.AppendText(Environment.NewLine + "Выполнен расчет повреждённости участка МГ от вмятин и гофр (ф. 5.8 СТО 292) d=" + result);
            return result;
        }
        private MGVTD damageFromDent(MGVTD mGVTD, plotBoundaries PlotBoundaries)//!!!В ЗАДАННЫХ ГРАНИЦАХ УЧАСТКА!!вычисляем повреждённость участка МГ от вмятин и гофр (ф. 5.8 СТО 292)
        {
            double result=0;
            double summOfRangs = 0;
            int defectPipesCount = 0;//создаём переменную для хранения количества труб с вмятинами

            if (String.IsNullOrEmpty(PlotBoundaries.pipeNumberOne))
            {
                if (String.IsNullOrWhiteSpace(PlotBoundaries.pipeNumberTwo))//если есть первая и последняя труба с дефектами в границах заданного участка
                {
                    //если оба значения адресов труб пусты
                    summOfRangs = 0;//если дефектов нет, то и считать ничего не надо.
                }
                else
                {
                    //если первое значение пустое а второе не пустое
                    //такого не может быть, логика работы программы этого не допустит

                }
            }
            else
            {
                if (String.IsNullOrWhiteSpace(PlotBoundaries.pipeNumberTwo))//если есть первая и последняя труба с дефектами в границах заданного участка
                {
                    //если первое значение не пустое а второе пустое
                    //тогда расчет проводим для одной единственной трубы
                   
                        if (mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].featuresCharacter.Contains("мятин"))
                        {
                            if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].defectVanishDate))
                            {
                                double tensileStrength = 510;//ищем по трубному журналу предел прочности
                                int pipelineSectionCategory = 1;//ищем по трубному журналу предел прочности
                                for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                                {
                                    if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].pipeNumber))
                                    {
                                        tensileStrength = mGVTD.MGPipeS[j].tensileStrength;//вот он предел прочности, нашли.
                                        try
                                        {
                                            pipelineSectionCategory = Convert.ToInt32(mGVTD.MGPipeS[j].pipelineSectionCategory);//и категория участка тоже нашлась
                                        }
                                        catch (Exception)
                                        {
                                            pipelineSectionCategory = 1;
                                        }
                                    }
                                }
                                defectPipesCount++;//инкриментируем счетчик труб с вмятинами
                                                   //тут будут проходить вычисления
                                double r = (mGVTD.pipelineInfo.pipeDiameter - mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].thikness) / 2;//радиус средней линии сечения (ф 7.14 Рекомендаций по оценке прочности...)
                                double a = mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].length / 2;//полоовина протяженности вмятины в осевом направлении
                                double b = mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].widht / 2;//половина протяженности дефекта в направлении кривой сечения
                                double u = (r * Math.PI) / (2 * b);//коэффициент (ф 7.16 Рекомендаций по оценке прочности...)
                                double t = (r * Math.PI) / (2 * a);//коэффициент (ф 7.15 Рекомендаций по оценке прочности...)
                            double pf = (1 - 0.3 * 0.3) * Math.Pow((r / mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].thikness), 3) * (mGVTD.pipelineInfo.operatingPressure / 206000);//поправка глубины вмятины на давление
                            double Uz = 225 * Math.Pow(t, 3) + 27 * 0.3 * Math.Pow(t, 2) * (9 * u * u - 5) + 25 * (Math.Pow(u, 3) + 1);//поправка глубины вмятины на давление
                            double H = 30 * (9 * u * u - 5) * pf / (Uz + 150 * (4 * u * u - 1) * pf);//поправка глубины вмятины на давление
                            double woo = Math.Pow(1 - H, -1) * mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].depthInMm;//поправка глубины вмятины на давление
                            double e2 = 0.5 * (mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].thikness / r) * (woo / r) * (3 * u * u - 1);//остаточные окружные деформации (ф 7.12 Рекомендаций по оценке прочности...)
                                double e1 = 0.5 * (mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].thikness / r) * (woo / r) * (3 * t * t - 1);//остаточные продольные деформации (ф 7.13 Рекомендаций по оценке прочности...)
                                double w000 = mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].depthInMm / mGVTD.pipelineInfo.pipeDiameter;//относиттельная глубина дефекта (ф 7.21 Рекомендаций по оценке прочности...)
                                double kr = 24;//определяем кооэффициент, зависящий от категории участка трубопровода
                                if (pipelineSectionCategory < 3)
                                {
                                    kr = 24;
                                }
                                else
                                {
                                    kr = 20;
                                }
                                double Rr = kr * Math.Max(e2, Math.Max(w000, e1));//это ранг опасности вмятины
                            mGVTD.MGPipeS[PlotBoundaries.pipeIdNumberOne].DentDamageList.Add(Rr);
                                //summOfRangs = Rr;//считаем сумму рангов опасности дефектов на трубопроводе
                            }
                        }
                    
                    result = summOfRangs;//повреждённость линейного участка МГ от вмятин (ф. 5.8 СТО 292)
                }
                else
                {
                    bool mark;
                    List<string> defectpipes = new List<string>();//это просто список учтенных труб
                    for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo; i++)
                    {
                        if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("мятин"))
                        {
                            if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectVanishDate))
                            {
                                double tensileStrength = 510;//ищем по трубному журналу предел прочности
                                int pipelineSectionCategory = 1;//ищем по трубному журналу предел прочности
                                for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                                {
                                    if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                                    {
                                        tensileStrength = mGVTD.MGPipeS[j].tensileStrength;//вот он предел прочности, нашли.
                                        try
                                        {
                                            pipelineSectionCategory = Convert.ToInt32(mGVTD.MGPipeS[j].pipelineSectionCategory);//и категория участка тоже нашлась
                                        }
                                        catch (Exception)
                                        {
                                            pipelineSectionCategory = 1;
                                        }
                                    }
                                }
                                

                                mark = true;
                                for (int q = 0; q < defectpipes.Count; q++)//проверяем, не учли ли мы уже эту трубу
                                {
                                    if (String.Equals(defectpipes[q], mGVTD.anomalyLogLineS[i].pipeNumber))
                                    {
                                        mark = false;
                                    }
                                }
                                if (mark)
                                {
                                    defectpipes.Add(mGVTD.anomalyLogLineS[i].pipeNumber);
                                    defectPipesCount++;//инкриментируем счетчик труб с вмятинами
                                }
                                
                                                   //тут будут проходить вычисления-20
                                double r = (mGVTD.pipelineInfo.pipeDiameter - mGVTD.anomalyLogLineS[i].thikness) / 2;//радиус средней линии сечения (ф 7.14 Рекомендаций по оценке прочности...)
                                double a = mGVTD.anomalyLogLineS[i].length / 2;//полоовина протяженности вмятины в осевом направлении
                                double b = mGVTD.anomalyLogLineS[i].widht / 2;//половина протяженности дефекта в направлении кривой сечения

                                double u = (r * Math.PI) / (2 * b);//коэффициент (ф 7.16 Рекомендаций по оценке прочности...)
                                //double u = (r / b) * (Math.PI / 2);//коэффициент (ф 7.16 Рекомендаций по оценке прочности...)
                                //double t = (r / a) * (Math.PI / 2);//коэффициент (ф 7.15 Рекомендаций по оценке прочности...)
                                double t = (r * Math.PI) / (2 * a);//коэффициент (ф 7.15 Рекомендаций по оценке прочности...
                                double pf = (1 - 0.3 * 0.3) * Math.Pow((r / mGVTD.anomalyLogLineS[i].thikness), 3) * (mGVTD.pipelineInfo.operatingPressure / 206000);//поправка глубины вмятины на давление
                                double Uz = 225 * Math.Pow(t, 3) + 27 * 0.3 * Math.Pow(t, 2) * (9 * u * u - 5) + 25 * (Math.Pow(u, 3) + 1);//поправка глубины вмятины на давление
                                double H = 30 * (9 * u * u - 5) * pf / (Uz + 150 * (4 * u * u - 1) * pf);//поправка глубины вмятины на давление
                                double woo = Math.Pow(1 - H, -1) * mGVTD.anomalyLogLineS[i].depthInMm;//поправка глубины вмятины на давление
                                double e2 = 0.5 * (mGVTD.anomalyLogLineS[i].thikness / r) * (woo / r) * (3 * u * u - 1);//остаточные окружные деформации (ф 7.12 Рекомендаций по оценке прочности...)
                                double e1 = 0.5 * (mGVTD.anomalyLogLineS[i].thikness / r) * (woo / r) * (3 * t * t - 1);//остаточные продольные деформации (ф 7.13 Рекомендаций по оценке прочности...)
                                //double e2 = 0.5 * (mGVTD.anomalyLogLineS[i].thikness / r) * (mGVTD.anomalyLogLineS[i].depthInMm / r) * (3 * u * u - 1);//остаточные окружные деформации (ф 7.12 Рекомендаций по оценке прочности...)
                                //double e1 = 0.5 * (mGVTD.anomalyLogLineS[i].thikness / r) * (mGVTD.anomalyLogLineS[i].depthInMm / r) * (3 * t * t - 1);//остаточные продольные деформации (ф 7.13 Рекомендаций по оценке прочности...)
                                double w000 = mGVTD.anomalyLogLineS[i].depthInMm / mGVTD.pipelineInfo.pipeDiameter;//относиттельная глубина дефекта (ф 7.21 Рекомендаций по оценке прочности...)
                                double kr = 24;//определяем кооэффициент, зависящий от категории участка трубопровода
                                if (pipelineSectionCategory < 3)
                                {
                                    kr = 24;
                                }
                                else
                                {
                                    kr = 20;
                                }
                                double Rr = kr * Math.Max(e2, Math.Max(w000, e1));//это ранг опасности вмятины
                                if (Rr>1)//условие, что значение искомой величины по определению не больше единицы
                                {
                                    Rr = 1;
                                }
                                for (int f = 0; f < mGVTD.MGPipeS.Count; f++)
                                {
                                    if (String.Equals(mGVTD.MGPipeS[f].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                                    {
                                        mGVTD.MGPipeS[f].DentDamageList.Add(Rr);
                                        //richTextBox2.AppendText(Environment.NewLine + "Вмятина. Труба № " + mGVTD.MGPipeS[f].pipeNumber+" Поврежденность: "+ Rr);//Для отладки. Потом закомментировать.
                                    }
                                }
                                //summOfRangs = summOfRangs + Rr;//считаем сумму рангов опасности дефектов на трубопроводе
                            }
                        }
                    }
                    result = summOfRangs / (PlotBoundaries.pipeIdNumberTwoPipeLog - PlotBoundaries.pipeIdNumberOnePipeLog);//повреждённость линейного участка МГ от вмятин (ф. 5.8 СТО 292)
                }
                
            }
            //summDentDamag = summOfRangs;//запоминаем суммарную поврежденность от вмятин
            allPipeWhithDent = defectPipesCount;//запоминаем количество труб с вмятинами
            //richTextBox1.AppendText(Environment.NewLine + "Выполнен расчет повреждённости участка МГ от вмятин и гофр (ф. 5.8 СТО 292) d=" + result);
            return mGVTD;
        }
        private double damagOfCoilJoin(MGVTD mGVTD)//вычисляем повреждённость участка МГ от дефектов КСС (ф. 5.8 СТО 292)
        {
            double result = 0;
            double summOfRangs = 0;
            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)
            {
                if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("кольцев"))
                {
                    if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectVanishDate))
                    {
                        double Rsh = 3.82 * (mGVTD.anomalyLogLineS[i].widht / mGVTD.pipelineInfo.pipeDiameter);//вычисляем ранг опасности дефектов КСС (ф. 6.10 СТО 292)
                        summOfRangs = summOfRangs + Rsh;//Сумма рангов
                    }
                }
            }
            result = summOfRangs / mGVTD.MGPipeS.Count;//вычисляем повреждённость участка МГ от дефектов КСС (ф. 5.10 СТО 292)
            richTextBox1.AppendText(Environment.NewLine + "Выполнен расчет повреждённости от дефектов КСС (ф. 5.8 СТО 292) d=" + result);
            return result;
        }
        private MGVTD damagOfCoilJoin(MGVTD mGVTD, plotBoundaries PlotBoundaries)//!!!В ЗАДАННЫХ ГРАНИЦАХ УЧАСТКА!!вычисляем повреждённость участка МГ от дефектов КСС (ф. 5.8 СТО 292)
        {
            double result = 0;
            double summOfRangs = 0;


            if (String.IsNullOrEmpty(PlotBoundaries.pipeNumberOne))
            {
                if (String.IsNullOrWhiteSpace(PlotBoundaries.pipeNumberTwo))//если есть первая и последняя труба с дефектами в границах заданного участка
                {
                    //если оба значения адресов труб пусты
                    summOfRangs = 0;//если дефектов нет, то и считать ничего не надо.
                    allPipeWhithJointDefects = 0;
                }
                else
                {
                    //если первое значение пустое а второе не пустое
                    //такого не может быть, логика работы программы этого не допустит

                }
            }
            else
            {
                if (String.IsNullOrWhiteSpace(PlotBoundaries.pipeNumberTwo))//если есть первая и последняя труба с дефектами в границах заданного участка
                {
                    //если первое значение не пустое а второе пустое
                    //тогда расчет проводим для одной единственной трубы

                    if (mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].featuresCharacter.Contains("кольцев"))
                    {
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].defectVanishDate))
                        {
                            double Rsh = 3.82 * (mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].widht / mGVTD.pipelineInfo.pipeDiameter);//вычисляем ранг опасности дефектов КСС (ф. 6.10 СТО 292)
                            if (Rsh > 1)//последнее условие п. 6.5 СТО 292.
                            {
                                Rsh = 1;
                            }
                            //summOfRangs = Rsh;//Сумма рангов
                            mGVTD.MGPipeS[PlotBoundaries.pipeIdNumberOne].JoinDamageList.Add(Rsh);
                            allPipeWhithJointDefects = 1;
                        }
                    }

                    result = summOfRangs;//вычисляем повреждённость участка МГ от дефектов КСС (ф. 5.10 СТО 292)
                }

                else
                {
                    bool mark;
                    List<string> defectpipes = new List<string>();//это просто список учтенных труб
                    for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo; i++)
                    {
                        if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("кольцев"))
                        {
                            if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectVanishDate))
                            {
                                mark = true;
                                for (int q = 0; q < defectpipes.Count; q++)//проверяем, не учли ли мы уже эту трубу
                                {
                                    if (String.Equals(defectpipes[q], mGVTD.anomalyLogLineS[i].pipeNumber))
                                    {
                                        mark = false;
                                    }
                                }
                                if (mark)
                                {
                                    defectpipes.Add(mGVTD.anomalyLogLineS[i].pipeNumber);
                                    allPipeWhithJointDefects++;//инкрементироуем количество дефектов
                                }

                                double Rsh = 3.82 * (mGVTD.anomalyLogLineS[i].widht / mGVTD.pipelineInfo.pipeDiameter);//вычисляем ранг опасности дефектов КСС (ф. 6.10 СТО 292)
                                if (Rsh > 1)//последнее условие п. 6.5 СТО 292.
                                {
                                    Rsh = 1;
                                }
                                //summOfRangs = summOfRangs + Rsh;//Сумма рангов

                                for (int f = 0; f < mGVTD.MGPipeS.Count; f++)
                                {
                                    if (String.Equals(mGVTD.MGPipeS[f].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                                    {
                                        mGVTD.MGPipeS[f].JoinDamageList.Add(Rsh);
                                        Rsh = 0;
                                    }
                                }
                            }
                        }
                    }
                }
                 result = summOfRangs / (PlotBoundaries.pipeIdNumberTwoPipeLog - PlotBoundaries.pipeIdNumberOnePipeLog);//вычисляем повреждённость участка МГ от дефектов КСС (ф. 5.10 СТО 292)                }
            }

            //summJointDefectsDamag = summOfRangs;//запоминаем суммарную поврежденность сварных соединений
            //richTextBox1.AppendText(Environment.NewLine + "Выполнен расчет повреждённости от дефектов КСС (ф. 5.8 СТО 292) d=" + result);
            return mGVTD;
        }
        public class plotBoundaries//класс для хранения имён труб на границах расчетного участка
            {
                public string pipeNumberOne;//имя первой трубы
                public string pipeNumberTwo;//имя второй трубы
                public int pipeIdNumberOne;//номер первой трубы
                public int pipeIdNumberTwo;//номер второй трубы
            public int pipeIdNumberOnePipeLog;//номер первой трубы участка в трубном журнале
            public int pipeIdNumberTwoPipeLog;//номер второй трубы участка в трубном журнале
        }
        private plotBoundaries lookingOfPlotBoundaries(MGVTD mGVTD, string pipeNumberOne, string pipeNumberTwo)//ищем имена и порядковые номера первой и последней труб в ведомости аномалий, попавших в заданный интервал трубного журнала. 
        {
            plotBoundaries result = new plotBoundaries();

            int firstPipeID=0;
            int secondPipeID= mGVTD.MGPipeS.Count;
            int marker = 0;//маркер для определения, что первая труба диапазона уже найдена
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                if (String.Equals(pipeNumberOne, mGVTD.MGPipeS[i].pipeNumber))
                {
                    firstPipeID = i;
                    result.pipeIdNumberOnePipeLog = firstPipeID;
                }
                if (String.Equals(pipeNumberTwo, mGVTD.MGPipeS[i].pipeNumber))
                {
                    secondPipeID = i;
                    result.pipeIdNumberTwoPipeLog = secondPipeID;
                }
                
            }
            for (int j = firstPipeID; j < secondPipeID; j++)
            {
                //richTextBox2.AppendText(Environment.NewLine + "j: " + j);
                for (int k = 0; k < mGVTD.anomalyLogLineS.Count; k++)
                {
                    if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[k].pipeNumber))
                    {
                        if (marker == 0)//если первая дефектная труба участка ещё не найдена
                        {
                            result.pipeIdNumberOne = k;
                            result.pipeNumberOne = mGVTD.anomalyLogLineS[k].pipeNumber;
                            marker = 1;
                        }
                        if (marker == 1)//если первая дефектная труба участка уже найдена
                        {
                            if (String.Equals(mGVTD.anomalyLogLineS[k].pipeNumber, result.pipeIdNumberOne))//это чтобы если участок будет содержать только одну дефектную трубу, поле номера второй трубы (которогй нет) осталось пустым
                            {

                            }
                            else
                            {
                                result.pipeIdNumberTwo = k;
                                result.pipeNumberTwo = mGVTD.anomalyLogLineS[k].pipeNumber;
                                marker = 1;
                            }

                        }
                    }
                }
            }
            return result;
        }
        private void button5_Click(object sender, EventArgs e)//чтение данных из таблицы ексель в экземпляр класса
        {
            //tableExcelReadToClass();//чтение данных из файла в экземпляр класса
            shortTableExcelReadToClass();//чтение данных из файла в экземпляр класса (избирательный метод)
            mGVTD = itIsTee(mGVTD);//помечаем соответствующие поля у секций, являющихся тройниками
            textBox131.Text = mGVTD.MGPipeS[0].pipeNumber;
            textBox136.Text= mGVTD.MGPipeS[mGVTD.MGPipeS.Count-1].pipeNumber;
        }
        private void button6_Click(object sender, EventArgs e)//выполнение расчета
        {
            
        }
        /*private void button1_Click(object sender, EventArgs e)
        {
        }*/
        /*private void button2_Click(object sender, EventArgs e)
        {
            //Form2 newForm = new Form2();
             //newForm.Show();
        }*/
        /*private void label91_Click(object sender, EventArgs e)
        {
        }*/
        
        private double damagFromСorrosionReal(MGVTD mGVTD, plotBoundaries PlotBoundaries)//сбор поврежденности от коррозии
        {
            double result = 0;
            for (int i = PlotBoundaries.pipeIdNumberOnePipeLog; i < PlotBoundaries.pipeIdNumberTwoPipeLog; i++)
            {
                double maxCorrDamagOfHhisPipe = 0;
                double maxDentDamagOfHhisPipe = 0;

                for (int j = 0; j < mGVTD.MGPipeS[i].corossionDamageList.Count; j++)
                {
                    if (maxCorrDamagOfHhisPipe < mGVTD.MGPipeS[i].corossionDamageList[j])
                    {
                        maxCorrDamagOfHhisPipe = mGVTD.MGPipeS[i].corossionDamageList[j];
                    }
                }
                for (int j = 0; j < mGVTD.MGPipeS[i].DentDamageList.Count; j++)
                {
                    if (maxDentDamagOfHhisPipe < mGVTD.MGPipeS[i].DentDamageList[j])
                    {
                        maxDentDamagOfHhisPipe = mGVTD.MGPipeS[i].DentDamageList[j];
                    }

                }
                if (maxCorrDamagOfHhisPipe > maxDentDamagOfHhisPipe)
                {
                    result = result + maxCorrDamagOfHhisPipe;
                   // richTextBox2.AppendText(Environment.NewLine + "_"+ mGVTD.MGPipeS[i].pipeNumber);
                }
            } 
            return result;
        }
        private double damagFromDentReal(MGVTD mGVTD, plotBoundaries PlotBoundaries)//сбор поврежденности от вмятин
        {
            double result = 0;
            for (int i = PlotBoundaries.pipeIdNumberOnePipeLog; i < PlotBoundaries.pipeIdNumberTwoPipeLog; i++)
            {
                double maxCorrDamagOfHhisPipe = 0;
                double maxDentDamagOfHhisPipe = 0;

                for (int j = 0; j < mGVTD.MGPipeS[i].corossionDamageList.Count; j++)
                {
                    if (maxCorrDamagOfHhisPipe < mGVTD.MGPipeS[i].corossionDamageList[j])
                    {
                        maxCorrDamagOfHhisPipe = mGVTD.MGPipeS[i].corossionDamageList[j];
                    }
                }
                for (int j = 0; j < mGVTD.MGPipeS[i].DentDamageList.Count; j++)
                {
                    if (maxDentDamagOfHhisPipe < mGVTD.MGPipeS[i].DentDamageList[j])
                    {
                        maxDentDamagOfHhisPipe = mGVTD.MGPipeS[i].DentDamageList[j];
                    }

                }
                if (maxCorrDamagOfHhisPipe < maxDentDamagOfHhisPipe)
                {
                    result = result+maxDentDamagOfHhisPipe;
                }

            }


            return result;
        }
        private double damagFromJoinReal(MGVTD mGVTD, plotBoundaries PlotBoundaries)//сбор поврежденности сварных соединений
        {
            double result = 0;
            for (int i = PlotBoundaries.pipeIdNumberOnePipeLog; i < PlotBoundaries.pipeIdNumberTwoPipeLog; i++)
            {
                
                double maxJoinOfHhisPipe = 0;
                for (int j = 0; j < mGVTD.MGPipeS[i].JoinDamageList.Count; j++)
                {
                   
                    if (maxJoinOfHhisPipe < mGVTD.MGPipeS[i].JoinDamageList[j])
                    {
                        maxJoinOfHhisPipe = mGVTD.MGPipeS[i].JoinDamageList[j];
                    }
                    
                }
                result = result + maxJoinOfHhisPipe;

            }


            return result;
        }
        private void button1_Click_1(object sender, EventArgs e)//проведение вычислений Pvtd !!!В ЗАДАННЫХ ГРАНИЦАХ УЧАСТКА!!
        {
        
        }
        private int MaxCorrDefectNumber (MGVTD mGVTD, plotBoundaries PlotBoundaries)//ищем номер строки с максимальным дефектом потери металла
        {
            double maxLostMetalProcent = 0;
            int numberPipeWithMaxDefect = 0;
            for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo; i++)//ищем максимальную глубину дефекта
            {
                if (mGVTD.anomalyLogLineS[i].isLostMetal)
                {
                    if (String.IsNullOrEmpty(mGVTD.anomalyLogLineS[i].defectVanishDate))
                    {
                        if (maxLostMetalProcent < mGVTD.anomalyLogLineS[i].depthInProcent)
                        {
                            maxLostMetalProcent = mGVTD.anomalyLogLineS[i].depthInProcent;
                            numberPipeWithMaxDefect = i;
                        }
                    }
                }                               
            }
            return numberPipeWithMaxDefect;

        }
        private int numberOfTriples(MGVTD mGVTD, plotBoundaries PlotBoundaries)//считаем количество тройников на участке
        {
            int result=0;
            for (int i = PlotBoundaries.pipeIdNumberOnePipeLog; i < PlotBoundaries.pipeIdNumberTwoPipeLog; i++)
            {
                if (mGVTD.MGPipeS[i].itIsTee)
                {
                    result++;
                }

            }


            return result;
        }
        private int numberOfDefectTriples(MGVTD mGVTD, plotBoundaries PlotBoundaries)//считаем количество дефектных тройников
        {
            int result=0;
            for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo; i++)
            {
                if (String.IsNullOrEmpty(mGVTD.anomalyLogLineS[i].defectVanishDate))
                {
                    if (String.IsNullOrEmpty(mGVTD.anomalyLogLineS[i].distanceFromTransverseWeld))
                    {

                    }
                    else
                    {
                        for (int j = PlotBoundaries.pipeIdNumberOnePipeLog; j < PlotBoundaries.pipeIdNumberTwoPipeLog; j++)
                        {
                            if (String.Equals(mGVTD.anomalyLogLineS[i].pipeNumber, mGVTD.MGPipeS[j].pipeNumber))
                            {
                                if (mGVTD.MGPipeS[j].itIsTee)
                                {
                                    result++;
                                }          
                            }
                        }
                    }
                }
            }
            return result;
        }
        private int numberOfDefectUnderRoads(MGVTD mGVTD, plotBoundaries PlotBoundaries)//НЕ РАБОТАЕТ!!!считаем количество дефектных стыков в кожухах
        {
            int result=0;
            int startPipe = 0;
            int finishPipe = 0;
            bool marker = false;
            bool mark = true;
            List<string> defectpipes = new List<string>();//это просто список учтенных труб
            for (int i = 0; i < mGVTD.furnishingsLogS.Count; i++)//пролистываем журнал элементов обустройства и ищем начало патрона
            {
                if (marker == false)//состояние поиска начала патрона
                {
                    if (mGVTD.furnishingsLogS[i].characterFeatures.Contains("рон нач"))//если нашли нгачало патрона, запоминаем номер строки
                    {
                        for (int f = 0; f < mGVTD.MGPipeS.Count; f++)
                        {
                            if (String.Equals(mGVTD.furnishingsLogS[i].pipeNumber, mGVTD.MGPipeS[f].pipeNumber))
                            {
                                marker = true;//состояние поиска конца патрона
                                startPipe = f;//номер первой трубы патрона в трубном журнале
                            }
                        }
                        
                       
                        if (marker == true)
                        {
                            for (int j = i; j < mGVTD.furnishingsLogS.Count; j++)//и ищем конец этого патрона
                            {
                                if (mGVTD.furnishingsLogS[j].characterFeatures.Contains("рон кон"))
                                {

                                    for (int f = startPipe; f < mGVTD.MGPipeS.Count; f++)
                                    {
                                        if (String.Equals(mGVTD.furnishingsLogS[i].pipeNumber, mGVTD.MGPipeS[f].pipeNumber))
                                        {
                                            finishPipe = f;
                                            marker = false;
                                        }
                                    }
                                    
                                    if (marker == false)
                                    {
                                        for (int k = startPipe; k < finishPipe; k++)
                                        {                                            
                                            if (mGVTD.MGPipeS[k].JoinDamageList.Count > 0)
                                            {
                                                mark = true;
                                                for (int q = 0; q < defectpipes.Count; q++)//проверяем, не учли ли мы уже эту трубу
                                                {
                                                    if (String.Equals(defectpipes[q], mGVTD.anomalyLogLineS[i].pipeNumber))
                                                    {
                                                        mark = false;
                                                    }
                                                }
                                                if (mark)
                                                {
                                                    defectpipes.Add(mGVTD.anomalyLogLineS[i].pipeNumber);
                                                    result++;
                                                    richTextBox2.AppendText(Environment.NewLine + "Дефектный стык № " + k + ", в кожухе от" + startPipe + " до " + finishPipe + " трубы");
                                                }                                                
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }
        private int numberOfDefectCoilUnderRoads(MGVTD mGVTD, plotBoundaries PlotBoundaries)//считаем количество дефектных стыков в кожухах
        {
            int result = 0;
            int startPipeID = 0;
            int finishPipeID = 0;
            List<string> defectpipes = new List<string>();//это просто список учтенных труб
            List<string> pipesUnderRoad = new List<string>();//это список труб, находящихся внутри кожуха
            bool isStartRoad = false;//признак того, что мы нашли начало кожуха и ещё не достигли его конца
            int startPosition = 0;
            bool mark = true;
            for (int i = startPosition; i < mGVTD.furnishingsLogS.Count; i++)//ищем начало кожуха
            {
                if (isStartRoad == false)//если еще не нашли начало кожуха
                {
                    if (mGVTD.furnishingsLogS[i].characterFeatures.Contains("рон нач"))
                    {
                        isStartRoad = true;
                        for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                        {
                            if (String.Equals(mGVTD.furnishingsLogS[i].pipeNumber, mGVTD.MGPipeS[j].pipeNumber))//нашли первую трубу кожуха в трубном журнале
                            {
                                startPipeID = j;
                                //richTextBox2.AppendText(Environment.NewLine + "Первая труба кожуха " + mGVTD.MGPipeS[startPipeID].pipeNumber);
                                for (int k = i; k < mGVTD.furnishingsLogS.Count; k++)//ищем конец кожуха
                                {
                                    if (isStartRoad == true)//если уже нашли начало кожуха
                                    {
                                        if (mGVTD.furnishingsLogS[k].characterFeatures.Contains("рон кон"))
                                        {
                                            startPosition = k;
                                            isStartRoad = false;
                                            for (int f = 0; f < mGVTD.MGPipeS.Count; f++)
                                            {
                                                if (String.Equals(mGVTD.furnishingsLogS[k].pipeNumber, mGVTD.MGPipeS[f].pipeNumber))//нашли последнюю трубу кожуха в трубном журнале
                                                {
                                                    finishPipeID = f;
                                                    //richTextBox2.AppendText(Environment.NewLine + "Последняя труба кожуха " + mGVTD.MGPipeS[finishPipeID].pipeNumber);
                                                    //richTextBox2.AppendText(Environment.NewLine + "======================================= ");
                                                    for (int q = startPipeID + 1; q < finishPipeID+1; q++)//+1 потому, что первый стык находится не в кожухе и мы его пропускаем
                                                    {
                                                        pipesUnderRoad.Add(mGVTD.MGPipeS[q].pipeNumber);//добавляем трубы, находящиеся в кожухе к специально созданноому списку
                                                        //richTextBox2.AppendText(Environment.NewLine + "Очередная труба кожуха " + mGVTD.MGPipeS[q].pipeNumber);
                                                    }
                                                    //richTextBox2.AppendText(Environment.NewLine + ">===> Добавили трубы кожуха к списку. Продолжаем...");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            for (int i = 0; i < pipesUnderRoad.Count; i++)
            {
                for (int j = PlotBoundaries.pipeIdNumberOne; j < PlotBoundaries.pipeIdNumberTwo; j++)
                {
                    if (String.Equals(pipesUnderRoad[i], mGVTD.anomalyLogLineS[j].pipeNumber))
                    {
                        if (mGVTD.anomalyLogLineS[j].featuresCharacter.Contains("кольцев"))
                        {
                            if (String.IsNullOrEmpty(mGVTD.anomalyLogLineS[j].defectVanishDate))//если дефект не помечен как устраненный
                            {
                                mark = true;
                                for (int q = 0; q < defectpipes.Count; q++)//проверяем, не учли ли мы уже эту трубу
                                {
                                    if (String.Equals(defectpipes[q], mGVTD.anomalyLogLineS[j].pipeNumber))
                                    {
                                        mark = false;
                                    }
                                }
                                if (mark)
                                {
                                    defectpipes.Add(mGVTD.anomalyLogLineS[j].pipeNumber);
                                    result++;
                                    richTextBox2.AppendText(Environment.NewLine + "Дефектный стык № " + mGVTD.anomalyLogLineS[j].pipeNumber + " внутри кожуха");
                                }
                            }
                            
                        }
                    }
                }
            }

            return result;
        }
        private int numberOfDefectLongitudinalWelds(MGVTD mGVTD, plotBoundaries PlotBoundaries)//считаем дефектные продольные швы
        {
            int result = 0;
            bool mark = true;
            List<string> defectpipes = new List<string>();//это просто список учтенных труб
            for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo; i++)
            {
                if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectVanishDate))
                {
                    if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("продольн"))
                    {
                        mark = true;
                        for (int q = 0; q < defectpipes.Count; q++)//проверяем, не учли ли мы уже эту трубу
                        {
                            if (String.Equals(defectpipes[q], mGVTD.anomalyLogLineS[i].pipeNumber))
                            {
                                mark = false;
                            }
                        }
                        if (mark)
                        {
                            defectpipes.Add(mGVTD.anomalyLogLineS[i].pipeNumber);
                            result++;//инкрементироуем количество дефектных труб
                        }
                    }
                }                

            }
            return result;
        }
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
        }

        private void textBox139_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox18_TextChanged(object sender, EventArgs e)
        {

        }

        private void StartCalculations_Click(object sender, EventArgs e)
        {
            allPipeCount = 0;//сумма всех труб участка++
            allPipeWhithСorrosion = 0;//сумма труб с коррозией++
            summCorrosionDamag = 0;//суммаорная поврежденность от коррозии++
            allPipeWhithСorrosionPlus = 0;//сумма труб с коррозией++
            summCorrosionDamagPlus = 0;//суммаорная поврежденность от коррозии++
            allPipeWhithDent = 0;//количество труб с вмятинами++
            summDentDamag = 0;//суммаорная поврежденность от вмятин++        
            technicalConditionIndicatorOfPipesAndSDT = 0;//показатель технического состояния труб и СДТ++
            allPipeWhithJointDefects = 0;//количество труб с дефектами КСС++
            summJointDefectsDamag = 0;//суммаорная поврежденность КСС

            double dCoil = 0;
            allDefectsWhithСorrosionPlus = 0;//сумма всех коррозионных дефектов глубиной больше указанного процента
            plotBoundaries PlotBoundaries = new plotBoundaries();
            mGVTD = isLostMetal(mGVTD);//расставляем метки на дефектах потери металла
            PlotBoundaries = lookingOfPlotBoundaries(mGVTD, textBox131.Text, textBox136.Text);
            allPipeCount = PlotBoundaries.pipeIdNumberTwoPipeLog - PlotBoundaries.pipeIdNumberOnePipeLog + 1;
            richTextBox2.AppendText(Environment.NewLine + "=======================================");
            richTextBox2.AppendText(Environment.NewLine + "Выполняется расчет Pвтд для участка газопровода в заданных границах");

            mGVTD = damagFromСorrosion(damageFromDent(damagOfCoilJoin(mGVTD, PlotBoundaries), PlotBoundaries), PlotBoundaries);

            double dc = 0;//поврежденность от трещиноподобных дефектов
            summCorrosionDamag = damagFromСorrosionReal(mGVTD, PlotBoundaries);
            double dk = summCorrosionDamag / allPipeCount;//поврежд от коррозии
            double Do = 0;//поврежденность при наличии овализации
            summDentDamag = damagFromDentReal(mGVTD, PlotBoundaries);
            double dr = summDentDamag / allPipeCount;//поврежденность связанная с наличием вмятин и гофр
            double dd = damagOfconnectingParts(mGVTD, PlotBoundaries);//поврежденность соединительных деталей
            double dJoin = damagFromJoinReal(mGVTD, PlotBoundaries);
            dCoil = dJoin / allPipeCount;//поврежденность сварных соединений-!!он же Показатель состояния


            double dSigma = 0;//от повышенного уровня напряжений
            double df = 0;//от переменных нагрузок
            double Pt = 1 - (1 - dc) * (1 - dk) * (1 - Do) * (1 - dr) * (1 - dd);//показатель технического состояния труб и соединительных деталей
            technicalConditionIndicatorOfPipesAndSDT = Pt;
            double Pvtd = 1 - (1 - Pt) * (1 - 0.5 * dCoil) * (1 - dSigma) * (1 - df * df);//Показатель технического состояния линейного участка МГ по результатам ВТД
            richTextBox2.AppendText(Environment.NewLine + "Выполнен расчет для участка МГ от трубы № " + textBox131.Text + " до трубы № " + textBox136.Text);

            richTextBox2.AppendText(Environment.NewLine + "Повреждённость соединительных деталей линейного участка (ф. 5.9 СТО 292): " + Math.Round(dd, 3));
            richTextBox2.AppendText(Environment.NewLine + "Повреждённость линейного участка МГ от вмятин и гофр (ф. 5.8 СТО 292): " + Math.Round(dr, 3));
            richTextBox2.AppendText(Environment.NewLine + "Повреждённость линейного участка МГ от от дефектов КСС (ф. 5.10 СТО 292): " + Math.Round(dCoil, 3));
            richTextBox2.AppendText(Environment.NewLine + "Pвтд= " + Math.Round(Pvtd, 3));
            richTextBox2.AppendText(Environment.NewLine + allPipeCount + ";" + allPipeWhithСorrosion + ";" + Math.Round(summCorrosionDamag, 3) + ";" + Math.Round(dk, 3) + ";" + Math.Round(dc, 3) + ";" + Math.Round(Do, 3) + ";" +
            allPipeWhithDent + ";" + Math.Round(summDentDamag, 3) + ";" + Math.Round(dr, 3) + ";" + Math.Round(dd, 3) + ";" + Math.Round(Pt, 3) + ";" + allPipeWhithJointDefects + ";" +
            Math.Round(dJoin, 3) + ";" + Math.Round(dCoil, 3) + ";" + Math.Round(0.85 * dCoil, 3) + ";" + dSigma + ";" + df + ";" + Math.Round(Pvtd, 3));
            double procent = 15;

            summJointDefectsDamag = 0;
            try
            {
                procent = Convert.ToDouble(textBox139.Text);
            }
            catch (Exception)
            {
                procent = 15;
                textBox139.Text = Convert.ToString(15);
            }
            PlotBoundaries = lookingOfPlotBoundaries(mGVTD, textBox131.Text, textBox136.Text);


            richTextBox2.AppendText(Environment.NewLine + "=======================================");
            allPipeWhithСorrosionPlus = 0;
            summCorr2 = 0;
            int summ = damagFromСorrosionAllDefects(mGVTD, PlotBoundaries, procent);
            double damagg = damagFromСorrosionProcent(mGVTD, PlotBoundaries, procent);
            double x0 = procent;//процент коррозии из окна на вкладке "анализ"
            double x1 = Math.Round(damagg, 3);//поврежденность
            richTextBox2.AppendText(Environment.NewLine + "Повреждённость локального участка от коррозии >" + procent + " % (ф. 5.3 СТО 292): " + Math.Round(damagg, 3));
            double x2 = summ;//количество
            richTextBox2.AppendText(Environment.NewLine + "Количество коррозионных дефектов глубиной >" + procent + " % : " + summ);
            richTextBox2.AppendText(Environment.NewLine + "=======================================");

            allPipeWhithСorrosionPlus = 0;
            summCorr2 = 0;
            summ = damagFromСorrosionAllDefects(mGVTD, PlotBoundaries, 30);
            damagg = damagFromСorrosionProcent(mGVTD, PlotBoundaries, 30);
            double lengthMG = mGVTD.MGPipeS[PlotBoundaries.pipeIdNumberTwoPipeLog].odometrDist - mGVTD.MGPipeS[PlotBoundaries.pipeIdNumberOnePipeLog].odometrDist;//протяженность участка
            double x3 = Math.Round(damagg, 3);//поврежденность от коррозии 30%
            double x4 = summ;//количество дефектов >30%
            richTextBox2.AppendText(Environment.NewLine + "Повреждённость локального участка от коррозии >" + 30 + " % (ф. 5.3 СТО 292): " + Math.Round(damagg, 3));
            richTextBox2.AppendText(Environment.NewLine + "Количество коррозионных дефектов глубиной >" + 30 + " % : " + summ);
            richTextBox2.AppendText(Environment.NewLine + "Плотность дефектов > 30%: " + 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3));//
            richTextBox2.AppendText(Environment.NewLine + "=======================================");
            double x5 = 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3);//Плотность дефектов > 30%
            summCorr2 = 0;
            summ = damagFromСorrosionAllDefects(mGVTD, PlotBoundaries, 0);
            damagg = damagFromСorrosionProcent(mGVTD, PlotBoundaries, 0);
            double x6 = Math.Round(damagg, 3);//поврежденность от коррозии
            double x7 = summ;//количество дефектов
            richTextBox2.AppendText(Environment.NewLine + "Повреждённость локального участка от коррозии (все корр. деф.)  (ф. 5.3 СТО 292): " + Math.Round(damagg, 3));
            richTextBox2.AppendText(Environment.NewLine + "Количество коррозионных дефектов : " + summ);
            richTextBox2.AppendText(Environment.NewLine + "Плотность коррозионных дефектов: " + 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3));//
            double x8 = 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3);//Плотность коррозионных дефектов
            richTextBox2.AppendText(Environment.NewLine + "=======================================");
            richTextBox2.AppendText(Environment.NewLine + "Доля труб с дефектами потери металла, %: " + Math.Round(100*Convert.ToDouble(allPipeWhithСorrosion) / allPipeCount, 3));
            richTextBox2.AppendText(Environment.NewLine + "Максимальная глубина дефекта потери металла: " + mGVTD.anomalyLogLineS[MaxCorrDefectNumber(mGVTD, PlotBoundaries)].depthInProcent);
            double x9= Math.Round(100 * Convert.ToDouble(allPipeWhithСorrosion) / allPipeCount, 3);//Доля труб с дефектами потери металла, %
            double x10 = mGVTD.anomalyLogLineS[MaxCorrDefectNumber(mGVTD, PlotBoundaries)].depthInProcent;//Максимальная глубина дефекта потери металла:
            summ = damagFromСorrosionAllDefects(mGVTD, PlotBoundaries, 15);//
            

            richTextBox2.AppendText(Environment.NewLine + "Плотность дефектов > 15%: " + 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3));//
            double x11= 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3);//Плотность дефектов > 15%
            richTextBox2.AppendText(Environment.NewLine + "Доля труб с дефектами геометрии, %: " + Math.Round(100*Convert.ToDouble(allPipeWhithDent) / allPipeCount, 3));
            double x12= Math.Round(100 * Convert.ToDouble(allPipeWhithDent) / allPipeCount, 3);//Доля труб с дефектами геометрии, %
            richTextBox2.AppendText(Environment.NewLine + "Общее количество тройников: " + numberOfTriples(mGVTD, PlotBoundaries));
            double x13= numberOfTriples(mGVTD, PlotBoundaries);//Общее количество тройников
            richTextBox2.AppendText(Environment.NewLine + "Количество дефектных тройников: " + numberOfDefectTriples(mGVTD, PlotBoundaries));
            double x14= numberOfDefectTriples(mGVTD, PlotBoundaries);//Количество дефектных тройников
            richTextBox2.AppendText(Environment.NewLine + "======================================= ");
            richTextBox2.AppendText(Environment.NewLine + "Количество дефектных труб в кожухах: " + numberOfDefectCoilUnderRoads(mGVTD, PlotBoundaries));
            double x15= numberOfDefectCoilUnderRoads(mGVTD, PlotBoundaries);//Количество дефектных труб в кожухах
            richTextBox2.AppendText(Environment.NewLine + "Количество аномальных поперечных швов: " + allPipeWhithJointDefects);
            double x16= allPipeWhithJointDefects;//Количество аномальных поперечных швов
            richTextBox2.AppendText(Environment.NewLine + "Количество аномальных продольных швов: " + numberOfDefectLongitudinalWelds(mGVTD, PlotBoundaries));
            double x17= numberOfDefectLongitudinalWelds(mGVTD, PlotBoundaries);//Количество аномальных продольных швов
            richTextBox2.AppendText(Environment.NewLine + x0 + ";" + x1+ ";" + x2 + ";" + x3 + ";" + x4 + ";" + x5 + ";" + x6 + ";" + x7 + ";" + x8 + ";" + x9 + ";" + x10 + ";" +
                x11 + ";" + x12 + ";" + x13 + ";" + x14 + ";" + x15 + ";" + x16 + ";" + x17);

        }
    }
}
