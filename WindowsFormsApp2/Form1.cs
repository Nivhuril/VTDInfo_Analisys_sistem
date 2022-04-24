using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using GMap.NET.MapProviders;
using GMap.NET;
using System.Globalization;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.Markers;
using GMap.NET.WindowsForms.ToolTips;
using System.Device.Location;
//using System.Windows.Forms.DataVisualization.Charting;
//antipov-db1 OByM6oBaKutjA2xv
namespace VTDinfo
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        public class numbersOfColumns//класс для хранения согласованных номеров столбцов и строк для импорта данных из файла EXCEL
        {

            //Следующие поля заполняются после обработки отчета
            public int featuresNumber_BHTTS;//Номер особенности///1
            public int pipeNumber_BHTTS;//номер трубы///2
            public int odometrDist_BHTTS;//дистанция по одометру///3
            public int distanceFromReferencePoints_BHTTS;//расстояние от реперных точек///4
            public int distanceToNextReferencePoints_BHTTS;//расстояние до следующей реперной точки///5
            public int featuresCharacter_BHTTS;//характер особенности///6
            public int distanceFromTransverseWeld_BHTTS;//расстояние от поперечного шва, м///7
            public int featuresOrientation_BHTTS;//угловая ориентация///8
            public int length_BHTTS;//длина///9
            public int widht_BHTTS;//ширина///10
            public int thikness_BHTTS;//толщина трубы///11
            public int depthInProcent_BHTTS;//глубина дефекта в процентах///12
            public int extOrInt_BHTTS;//характер локаизации(внутри или снаружи)///13
            public int note_BHTTS;//Примечание///14
            public int defectVanishDate;//дата устранения дефекта

            public int pipeID_BHTTS;//            
            public int pipeLength_BHTTS;//длина трубы               
            public int classOfSize_BHTTS;//класс размера
            public int depthInMm_BHTTS;//глубина дефекта в миллиметрах
            public int KBD_BHTTS;//КБД
            public int defectAssessment_BHTTS;//оценка дефекта
            public int Latitude_BHTTS;//Широта
            public int Longitude_BHTTS;//Долгота
            public int heightAboveSeaLevel_BHTTS;//H, м
            public int defectVanishDate_BHTTS;//дата устранения дефекта



            public int pipelineSectionCategory_BHTTS;//!!!категория участка трубопровода - заполняется при обработке массива
            public int steelGrade_BHTTS;//!!!марка стали - заполняется при обработке массива
            public int yieldPoint_BHTTS;//!!!предел текучести - заполняется при обработке массива
            public int tensileStrength_BHTTS;//!!!предел прочности - заполняется при обработке массива



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
            public int columnNumber13;//примечание
            public int columnNumber14;//для категории в трубном журнале
            public int columnNumber15;//для предела прочности в трубном журнале
            public int columnNumber16;//для предела текучести в трубном журнале
            public int columnNumber17;//марки стали в трубном журнале

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
            public bool itIsTee = false;//истина, если секция является тройником
            //это заполняется путём расчета повреждённости различного типа
            public List<double> corossionDamageList = new List<double>();//поврежденность трубы от коррозии
            public List<double> DentDamageList = new List<double>();//поврежденность трубы от вмятин
            public List<double> JoinDamageList = new List<double>();//поврежденность трубы от дефектов КСС

            public double MaximumCorrProcent;//максимальная глубина коррозии в процентах, выявленная на данной трубе
            public double MaximumDamageCorr;//максимальная поврежденность от коррозии выявленная на данной трубе
            public double MaximumDentDepth;//максимальная глубина вмятин, выявленная на данной трубе
            public double MaximumDentDamage;//максимальная поврежденность от вмятин, выявленная на данной трубе
            public double MaximumJoinDamage;//максимальная поврежденность от дефектов швов, выявленная на данной трубе
            public double critikalThikness;//критическая толщина, вычисленная для данной трубы
            public double residualResource;//остаточный ресурс трубы


            public double firstJointAngle;
            public double secondJointAngle;
            public bool isTwoJoint;

            public bool isInsulationDefect;
        }

        MGVTD setTypesForIUSTToFirnishingLog(MGVTD input)
        {
            MGVTD result = input;
            for (int i = 0; i < input.furnishingsLogS.Count; i++)
            {
                if (input.furnishingsLogS[i].characterFeatures.Contains("Кран"))
                {
                    result.furnishingsLogS[i].typeForIUST = "Кран";
                }
                else if (input.furnishingsLogS[i].characterFeatures.Contains("кран"))
                {
                    result.furnishingsLogS[i].typeForIUST = "Кран";
                }
                else if (input.furnishingsLogS[i].characterFeatures.Contains("запус"))
                {
                    result.furnishingsLogS[i].typeForIUST = "Камера запуска";
                }
                else if (input.furnishingsLogS[i].characterFeatures.Contains("приема"))
                {
                    result.furnishingsLogS[i].typeForIUST = "Камера приема";
                }
                else if (input.furnishingsLogS[i].characterFeatures.Contains("аркер"))
                {
                    result.furnishingsLogS[i].typeForIUST = "Маркер";
                }
                else if (input.furnishingsLogS[i].characterFeatures.Contains("ройник"))
                {
                    result.furnishingsLogS[i].typeForIUST = "Тройник";
                }
                else if (input.furnishingsLogS[i].characterFeatures.Contains("твод"))
                {
                    result.furnishingsLogS[i].typeForIUST = "Отвод";
                }
                else
                {
                    result.furnishingsLogS[i].typeForIUST = "Прочее";
                }
            }
            return result;
        }

        MGVTD setJointAnglesToMgPipesGPAS(MGVTD input)//расстановка информации об ориентации продольных швов в трубном журнале
        {
            MGVTD result = input;
            for (int i = 0; i < input.MGPipeS.Count; i++)
            {
                if (String.IsNullOrWhiteSpace(input.MGPipeS[i].clockOrientation)==false)
                {
                    if (input.MGPipeS[i].clockOrientation.Contains("/"))
                    {
                        result.MGPipeS[i].firstJointAngle = GetStartAngle(input.MGPipeS[i].clockOrientation);
                        if (result.MGPipeS[i].firstJointAngle<6)
                        {
                            result.MGPipeS[i].secondJointAngle = result.MGPipeS[i].firstJointAngle + 6;
                        }
                        else
                        {
                            result.MGPipeS[i].secondJointAngle = result.MGPipeS[i].firstJointAngle - 6;
                        }
                        result.MGPipeS[i].isTwoJoint = true;
                    }
                    else
                    {
                        result.MGPipeS[i].firstJointAngle = GetStartAngle(input.MGPipeS[i].clockOrientation);
                        result.MGPipeS[i].isTwoJoint = false;
                    }
                }                
            }
            return result;
        }
        MGVTD GetMaximumValues(MGVTD mGVTD)//поиск максимальной повреждённости на трубе
        {
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {


                double maxCorrDem = 0;//вычисляем максимальную поврежденность от коррозии
                for (int j = 0; j < mGVTD.MGPipeS[i].corossionDamageList.Count; j++)
                {
                    if (mGVTD.MGPipeS[i].corossionDamageList[j] > maxCorrDem)
                    {
                        maxCorrDem = mGVTD.MGPipeS[i].corossionDamageList[j];
                    }
                }
                mGVTD.MGPipeS[i].MaximumDamageCorr = maxCorrDem;

                double maxDentDem = 0;//вычисляем максимальную поврежденность от вмятин
                for (int j = 0; j < mGVTD.MGPipeS[i].DentDamageList.Count; j++)
                {
                    if (mGVTD.MGPipeS[i].DentDamageList[j] > maxCorrDem)
                    {
                        maxDentDem = mGVTD.MGPipeS[i].DentDamageList[j];
                    }
                }
                mGVTD.MGPipeS[i].MaximumDentDamage = maxDentDem;

                double maxJointDem = 0;//вычисляем максимальную поврежденность от дефектов швов
                for (int j = 0; j < mGVTD.MGPipeS[i].JoinDamageList.Count; j++)
                {
                    if (mGVTD.MGPipeS[i].JoinDamageList[j] > maxJointDem)
                    {
                        maxJointDem = mGVTD.MGPipeS[i].JoinDamageList[j];
                    }
                }
                mGVTD.MGPipeS[i].MaximumJoinDamage = maxJointDem;
            }

            return mGVTD;
        }

        MGVTD GetMaximumValuesCorrerionInProcent(MGVTD mGVTD)//поиск максимальной коррозии на трубе
        {
            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)
            {
                if (mGVTD.anomalyLogLineS[i].featuresCharacter.Contains("орроз"))
                {
                    for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                    {
                        if (String.Equals(mGVTD.anomalyLogLineS[i].pipeNumber, mGVTD.MGPipeS[j].pipeNumber))
                        {
                            if (mGVTD.MGPipeS[j].MaximumCorrProcent < mGVTD.anomalyLogLineS[i].depthInProcent)
                            {
                                mGVTD.MGPipeS[j].MaximumCorrProcent = mGVTD.anomalyLogLineS[i].depthInProcent;
                            }
                        }
                    }
                }
            }



            return mGVTD;
        }

        public class pipeSectionLog
        {
            public int pipelineID;//трубопровод (название)
            public string LPUMG_name;//участок трубопровода
            public string pipelineName;//трубопровод (название)
            public string pipelineSection;//участок трубопровода
            public string isVTD;//если есть данные о ВТД, будет не пустым
        }
        List<pipeSectionLog> pipeSectionS = new List<pipeSectionLog>();//создаём экзампляр класса для хранения списка учтённых участков МГ
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
            public string contractor;//подрядчик

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
            public string featuresOrientation;//ориентация
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
            public string defectRepareDate;//дата устранения дефекта
            public bool isLostMetal = false;

            //Для ИУС Т

            public int defectNumber;//номер дефекта по порядку
            //public string pipeNumber;//номер трубы
            public string defectType;//тип дефекта
            public string defectCode;//код дефекта
            public double distanceFromTransverseWeldIUST;//расстояние от первого поперечного шва
            public double distanceFromLongitudinalWeld;//расстояние от продольного шва
            public double start_angle;//начальный угол дефекта
            //public double length;//длина
            //public double widht;//ширина
            //public double depthInMm;//глубина дефекта в миллиметрах
            public string inside_or_outside;//внутренний, наружный, внутристенный
            public string defect_location;//расположение дефекта (основной металл, сварной шов, околошовная зона)
            public string danger_level;//уровень опасности (закритический, критический, допустимый)

            public double clockOrientation;//Ориент., ч:мин
            public double pipeLength;//длина трубы
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

            public string typeForIUST;//тип элемента для ШСЗ ИУСТ
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

        public class BHTTS_pipelog_String//строка трубного журнала БХТТС
        {
            public string featuresNumber_BHTTS;//Номер особенности///1
            public string pipeNumber_BHTTS;//номер трубы///2
            public double odometrDist_BHTTS;//дистанция по одометру///3
            public string distanceFromReferencePoints_BHTTS;//расстояние от реперных точек///4
            public double distanceToNextReferencePoints_BHTTS;//расстояние до следующей реперной точки///5
            public string featuresCharacter_BHTTS;//характер особенности///6
            public string distanceFromTransverseWeld_BHTTS;//расстояние от поперечного шва, м///7
            public string featuresOrientation_BHTTS;//угловая ориентация///8
            public double length_BHTTS;//длина///9
            public double widht_BHTTS;//ширина///10
            public double thikness_BHTTS;//толщина трубы///11
            public double depthInProcent_BHTTS;//глубина дефекта в процентах///12
            public string extOrInt_BHTTS;//характер локаизации(внутри или снаружи)///13
            public string note_BHTTS;//Примечание///14
            public string defectVanishDate;//дата устранения дефекта
        }
        List<BHTTS_pipelog_String> BHTTS_pipelog = new List<BHTTS_pipelog_String>();//создаём список для хранения строк трубного журнала БХТТС
        numbersOfColumns NumbersOfColumns = new numbersOfColumns();//создаём экземпляр класса ссылок на столбцы и строки отчета ВТД
        public int summCorr2 = 0;
        public int allPipeCount = 0;//сумма всех труб участка++
        public int allPipeWhithСorrosion = 0;//сумма труб с коррозией++
        public double summCorrosionDamag = 0;//суммарная поврежденность от коррозии++
        //поврежденность участка от коррозии dk++
        //поврежденность участка от трещин (0)++
        //поврежденность участка от овализации(0)++
        public int allPipeWhithDent = 0;//количество труб с вмятинами++
        public double summDentDamag = 0;//суммарная поврежденность от вмятин++
        //поврежденность участка от вмятин dr++
        //public double allconnectingPartsWhithDefects;//поврежденность тройников++
        public double technicalConditionIndicatorOfPipesAndSDT = 0;//показатель технического состояния труб и СДТ++
        public int allPipeWhithJointDefects = 0;//количество труб с дефектами КСС++
        public double summJointDefectsDamag = 0;//суммарная поврежденность КСС
        //показатель технического состояния кольцевых швов по результатам ВТД (pш)
        //показатель технического состояния по результатам шурфовок (pш*0,85)
        //поврежденность участка от переменных нагрузок (0)
        public double allDefectsWhithСorrosionPlus = 0;
        public double summCorrosionDamagPlus = 0;//суммарная поврежденность от коррозии++
        public int allPipeWhithСorrosionPlus = 0;//сумма труб с коррозией  выше заданного значения
        public string fileName;//переменная для хранения пути к файлу с отчетом
        public double ConvertToDouble(string inputString)//конвертируем строку в число double
        {
            double result = 0;

            if (Double.TryParse(inputString.Trim().Replace(".", ","), out result))
            {
                result = Double.Parse(inputString.Trim().Replace(".", ","));
            }

            return result;
        }
        public int ConvertToInt(string inputString)//конвертируем строку в число int
        {
            int result = 0;
            if (Int32.TryParse(inputString.Trim().Replace(".", ","), out result))
            {
                result = Int32.Parse(inputString.Trim().Replace(".", ","));
            }
            return result;
        }
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
        private void tableArdesTestSODPipeLog()//(для трубника СОД)метод для проверки правильности адресации ячеек и заполнения экземпляра класса numbersOfColumns()
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileNamePipeLog, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox170.Text;//получаем название вкладки из формы импотра
            try
            {
                ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];
            }
            catch (Exception)
            {
                ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2.Replace(".xlsx", "")];
                textBox170.Text = WorksheetName2.Replace(".xlsx", "");
            }



            //получаем номера столбцов для "трубного журлала"
            int columnNumber1 = Convert.ToInt16(textBox169.Text);//номер трубы
            int columnNumber2 = Convert.ToInt16(textBox168.Text);//дист
            int columnNumber3 = Convert.ToInt16(textBox167.Text);//толщина
            int columnNumber4 = Convert.ToInt16(textBox166.Text);//Длина трубы

            int columnNumber6 = Convert.ToInt16(textBox164.Text);//Характер особ.
            int columnNumber7 = Convert.ToInt16(textBox163.Text);//Ориент.
            //int columnNumber9 = Convert.ToInt16(textBox163.Text);//Ориент.!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

            int columnNumber13 = Convert.ToInt16(textBox144.Text);//примечание

            int columnNumber14 = Convert.ToInt16(textBox172.Text);//Категория
            int columnNumber15 = Convert.ToInt16(textBox175.Text);//Предел прочности

            //выводим значения соответствующих ячеек для проверки
            textBox157.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber1].Text);//номер трубы
            textBox156.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber2].Text);//дист
            textBox155.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber3].Text);//толщина
            textBox154.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber4].Text);//Длина трубы

            textBox152.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber6].Text);//Характер особ.
            textBox151.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber7].Text);//Ориент.

            textBox145.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber13].Text);//примечание

            textBox171.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber14].Text);//Категория
            textBox174.Text = Convert.ToString(ObjWorkSheet2.Cells[2, columnNumber15].Text);//Предел прочности


            //заполняем ссылки на номера столбцов для "трубного журлала"
            NumbersOfColumns.columnNumber1 = Convert.ToInt16(textBox169.Text);//номер трубы
            NumbersOfColumns.columnNumber2 = Convert.ToInt16(textBox168.Text);//дист
            NumbersOfColumns.columnNumber3 = Convert.ToInt16(textBox167.Text);//толщина
            NumbersOfColumns.columnNumber4 = Convert.ToInt16(textBox166.Text);//Длина трубы

            NumbersOfColumns.columnNumber6 = Convert.ToInt16(textBox164.Text);//Характер особ.
            NumbersOfColumns.columnNumber7 = Convert.ToInt16(textBox163.Text);//Ориент.
            //NumbersOfColumns.columnNumber9 = Convert.ToInt16(textBox163.Text);//Ориент.!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            NumbersOfColumns.columnNumber13 = Convert.ToInt16(textBox144.Text);//примечание

            NumbersOfColumns.columnNumber14 = Convert.ToInt16(textBox172.Text);//категория в трубном журнале
            NumbersOfColumns.columnNumber15 = Convert.ToInt16(textBox175.Text);//предел прочности в трубном журнале


            ObjExcel.Quit();


        }
        private void tableArdesTestSODDefectLog()//(для дефектов СОД)метод для проверки правильности адресации ячеек и заполнения экземпляра класса numbersOfColumns()
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileNameDefectLog, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet3;
            string WorksheetName3 = textBox196.Text;//получаем название вкладки из формы импотра


            try
            {
                ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];
            }
            catch (Exception)
            {
                ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3.Replace(".xlsx", "")];
                textBox196.Text = WorksheetName3.Replace(".xlsx", "");
            }



            //номера столбцов для "журлала аномалий"
            int column2Number1 = Convert.ToInt16(textBox195.Text);//дист по одом
            //int column2Number2 = Convert.ToInt16(textBox194.Text);//толщ
            int column2Number3 = Convert.ToInt16(textBox193.Text);//Расст. от ПОПШ
            int column2Number4 = Convert.ToInt16(textBox192.Text);//расст. от реперной т.
            int column2Number5 = Convert.ToInt16(textBox191.Text);//Характ. особ.
            int column2Number6 = Convert.ToInt16(textBox190.Text);//Класс размера
            int column2Number7 = Convert.ToInt16(textBox189.Text);//Ориентац
            int column2Number8 = Convert.ToInt16(textBox188.Text);//Длина
            int column2Number9 = Convert.ToInt16(textBox187.Text);//ширина
            int column2Number10 = Convert.ToInt16(textBox186.Text);//d %


            int column2Number11 = Convert.ToInt16(textBox185.Text);//d мм

            int column2Number12 = Convert.ToInt16(textBox184.Text);//Тип пол.
            int column2Number13 = Convert.ToInt16(textBox183.Text);//КБД
            int column2Number14 = Convert.ToInt16(textBox182.Text);//Оценка
            //int column2Number15 = Convert.ToInt16(textBox81.Text);
            //int column2Number16 = Convert.ToInt16(textBox82.Text);
            //int column2Number17 = Convert.ToInt16(textBox83.Text);
            //int column2Number18 = Convert.ToInt16(textBox84.Text);
            int column2Number19 = Convert.ToInt16(textBox216.Text);//дата устранения
            int column2Number20 = Convert.ToInt16(textBox176.Text);//для номера трубы
            int numb = 128;
            textBox197.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number1].Text);//дист по одом
            //textBox198.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number2].Text);//толщ
            textBox199.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number3].Text);//Расст. от ПОПШ
            textBox200.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number4].Text);//расст. от реперной т.
            textBox201.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number5].Text);//Характ. особ.
            textBox202.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number6].Text);//Класс размера
            textBox203.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number7].Text);//Ориентац
            textBox204.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number8].Text);//Длина
            textBox205.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number9].Text);//ширина
            textBox206.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number10].Text);//d %
            textBox207.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number11].Text);//d мм
            textBox208.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number12].Text);//Тип пол.
            textBox209.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number13].Text);//КБД
            textBox210.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number14].Text);//Оценка
            //textBox60.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number15].Text);
            //textBox61.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number16].Text);
            //textBox62.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number17].Text);
            //textBox63.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number18].Text);
            textBox215.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number19].Text);//дата устранения
            textBox177.Text = Convert.ToString(ObjWorkSheet3.Cells[numb, column2Number20].Text);//для номера трубы



            //номера столбцов для "журлала аномалий"
            NumbersOfColumns.column2Number1 = Convert.ToInt16(textBox195.Text);//дист по одом
            //NumbersOfColumns.column2Number2 = Convert.ToInt16(textBox194.Text);//толщ
            NumbersOfColumns.column2Number3 = Convert.ToInt16(textBox193.Text);//Расст. от ПОПШ
            NumbersOfColumns.column2Number4 = Convert.ToInt16(textBox192.Text);//расст. от реперной т.
            NumbersOfColumns.column2Number5 = Convert.ToInt16(textBox191.Text);//Характ. особ.
            NumbersOfColumns.column2Number6 = Convert.ToInt16(textBox190.Text);//Класс размера
            NumbersOfColumns.column2Number7 = Convert.ToInt16(textBox189.Text);//Ориентац
            NumbersOfColumns.column2Number8 = Convert.ToInt16(textBox188.Text);//Длина
            NumbersOfColumns.column2Number9 = Convert.ToInt16(textBox187.Text);//ширина
            NumbersOfColumns.column2Number10 = Convert.ToInt16(textBox186.Text);//d %
            NumbersOfColumns.column2Number11 = Convert.ToInt16(textBox185.Text);//d мм
            NumbersOfColumns.column2Number12 = Convert.ToInt16(textBox184.Text);//Тип пол.
            NumbersOfColumns.column2Number13 = Convert.ToInt16(textBox183.Text);//КБД
            NumbersOfColumns.column2Number14 = Convert.ToInt16(textBox182.Text);//Оценка
            //NumbersOfColumns.column2Number15 = Convert.ToInt16(textBox81.Text);
            //NumbersOfColumns.column2Number16 = Convert.ToInt16(textBox82.Text);
            //NumbersOfColumns.column2Number17 = Convert.ToInt16(textBox83.Text);
            //NumbersOfColumns.column2Number18 = Convert.ToInt16(textBox84.Text);
            NumbersOfColumns.column2Number19 = Convert.ToInt16(textBox216.Text);//дата устранения
            NumbersOfColumns.column2Number20 = Convert.ToInt16(textBox176.Text);//для номера трубы



            ObjExcel.Quit();


        }
        private void tableArdesTestLineObjects()//(для дефектов СОД)метод для проверки правильности адресации ячеек и заполнения экземпляра класса numbersOfColumns()
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileNameLineObjects, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet4;
            string WorksheetName4 = textBox231.Text;//получаем название вкладки из формы импотра

            try
            {
                ObjWorkSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName4];
            }
            catch (Exception)
            {
                ObjWorkSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName4.Replace(".xlsx", "")];
                textBox231.Text = WorksheetName4.Replace(".xlsx", "");
            }




            //номера столбцов для "журнала элементов обустройства"
            //int column3Number1 = Convert.ToInt16(textBox96.Text);
            int column3Number2 = Convert.ToInt16(textBox229.Text);
            int column3Number3 = Convert.ToInt16(textBox228.Text);
            //int column3Number4 = Convert.ToInt16(textBox99.Text);
            //int column3Number5 = Convert.ToInt16(textBox100.Text);
            int column3Number6 = Convert.ToInt16(textBox227.Text);
            int column3Number7 = Convert.ToInt16(textBox226.Text);
            //int column3Number8 = Convert.ToInt16(textBox103.Text);
            //int column3Number9 = Convert.ToInt16(textBox104.Text);
            //int column3Number10 = Convert.ToInt16(textBox105.Text);
            //int column3Number11 = Convert.ToInt16(textBox106.Text);
            //int column3Number12 = Convert.ToInt16(textBox107.Text);
            int column3Number13 = Convert.ToInt16(textBox224.Text);

            //textBox64.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number1].Text);
            textBox233.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number2].Text);
            textBox234.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number3].Text);
            //textBox85.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number4].Text);
            //textBox86.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number5].Text);
            textBox235.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number6].Text);
            textBox236.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number7].Text);
            //textBox89.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number8].Text);
            //textBox90.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number9].Text);
            //textBox91.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number10].Text);
            //textBox92.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number11].Text);
            //textBox93.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number12].Text);
            textBox238.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number13].Text);

            //номера столбцов для "журнала элементов обустройства"
            //NumbersOfColumns.column3Number1 = Convert.ToInt16(textBox96.Text);
            NumbersOfColumns.column3Number2 = Convert.ToInt16(textBox229.Text);
            NumbersOfColumns.column3Number3 = Convert.ToInt16(textBox228.Text);
            //NumbersOfColumns.column3Number4 = Convert.ToInt16(textBox99.Text);
            //NumbersOfColumns.column3Number5 = Convert.ToInt16(textBox100.Text);
            NumbersOfColumns.column3Number6 = Convert.ToInt16(textBox227.Text);
            NumbersOfColumns.column3Number7 = Convert.ToInt16(textBox226.Text);
            //NumbersOfColumns.column3Number8 = Convert.ToInt16(textBox103.Text);
            //NumbersOfColumns.column3Number9 = Convert.ToInt16(textBox104.Text);
            //NumbersOfColumns.column3Number10 = Convert.ToInt16(textBox105.Text);
            //NumbersOfColumns.column3Number11 = Convert.ToInt16(textBox106.Text);
            //NumbersOfColumns.column3Number12 = Convert.ToInt16(textBox107.Text);
            NumbersOfColumns.column3Number13 = Convert.ToInt16(textBox224.Text);

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
            while (looking)//ищем начало и конец первого журнала
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

        private List<pipeSectionLog> readToClassPipeSectionLog(string fileName)//метод для чтения из файла списка участков
        {
            List<pipeSectionLog> pipeSectionS = new List<pipeSectionLog>();
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = "Сводная таблица";//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            richTextBox7.AppendText(Environment.NewLine + "Выполняется обработка журнала учтенных объектов...");
            richTextBox7.AppendText(Environment.NewLine + "->*");

            int incrementor = 0;//переменная для прогресс - индикатора
            bool noFinalString = true;
            int i = 5;
            while (noFinalString)
            {

                pipeSectionLog PipeSectionLog = new pipeSectionLog();//создаём экземпляр класса строки журнала

                try
                {
                    PipeSectionLog.pipelineID = Convert.ToInt16(ObjWorkSheet2.Cells[i, 1].Text);
                }
                catch (Exception)
                {
                    PipeSectionLog.pipelineID = 0;
                    noFinalString = false;
                }
                // richTextBox7.AppendText(Environment.NewLine + Convert.ToString(i));

                PipeSectionLog.LPUMG_name = Convert.ToString(ObjWorkSheet2.Cells[i, 2].Text);//
                PipeSectionLog.pipelineName = Convert.ToString(ObjWorkSheet2.Cells[i, 3].Text);//
                PipeSectionLog.pipelineSection = Convert.ToString(ObjWorkSheet2.Cells[i, 4].Text);//
                PipeSectionLog.isVTD = Convert.ToString(ObjWorkSheet2.Cells[i, 27].Text);//
                i++;
                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor == 50)
                {
                    richTextBox7.AppendText("*");
                    incrementor = 0;
                }

                if (String.IsNullOrEmpty(PipeSectionLog.LPUMG_name))
                {
                    noFinalString = false;
                }
                else
                {
                    pipeSectionS.Add(PipeSectionLog);//добавляем заполненный экземпляр класса к списку
                }
            }



            richTextBox7.AppendText(Environment.NewLine + "Массив данных из журнала учтенных объектов прочитан. Количество строк: " + pipeSectionS.Count);
            richTextBox7.AppendText(Environment.NewLine + "==========================================");
            return pipeSectionS;
        }
        public class checkedItem
        {
            public int pipelineID;//трубопровод (название)
            public string LPUMG_name;//участок трубопровода
            public string pipelineName;//трубопровод (название)
            public string pipelineSection;//участок трубопровода
        }

        private void get_MG_ID(List<pipeSectionLog> pipeSectionS)
        {
            richTextBox7.AppendText(Environment.NewLine + CheckedItem.LPUMG_name + "_" + CheckedItem.pipelineName + "_" + CheckedItem.pipelineSection);

            for (int i = 0; i < pipeSectionS.Count; i++)
            {

                if (String.Equals(CheckedItem.LPUMG_name, pipeSectionS[i].LPUMG_name) & String.Equals(CheckedItem.pipelineName, pipeSectionS[i].pipelineName) & String.Equals(CheckedItem.pipelineSection, pipeSectionS[i].pipelineSection))
                {
                    MG_ID.Text = Convert.ToString(pipeSectionS[i].pipelineID);
                    if (String.IsNullOrWhiteSpace(pipeSectionS[i].isVTD) == false)
                    {
                        MG_ID.BackColor = Color.Green;
                    }
                    else
                    {
                        MG_ID.BackColor = Color.Empty;
                    }
                    CheckedItem.pipelineID = pipeSectionS[i].pipelineID;
                }
            }




        }
        private void setComboBoxes(List<pipeSectionLog> pipeSectionS)//добавляем в комбобокс с перечнем ЛПУМГ названия ЛПУМГ
        {
            LPUMG_Check.SelectedIndex = -1;
            MG_Check.SelectedIndex = -1;
            pipelineSection_Check.SelectedIndex = -1;
            LPUMG_Check.Text = "";
            MG_Check.Text = "";
            pipelineSection_Check.Text = "";

            for (int i = 0; i < pipeSectionS.Count; i++)
            {
                bool mark = true;
                for (int j = 0; j < LPUMG_Check.Items.Count; j++)
                {
                    if (LPUMG_Check.Items[j].Equals(pipeSectionS[i].LPUMG_name))
                    {
                        mark = false;
                    }
                }
                if (mark)
                {
                    LPUMG_Check.Items.Add(pipeSectionS[i].LPUMG_name);
                }

                //LPUMG_Check.Items.Add(pipeSectionS[i].LPUMG_name);
                //MG_Check.Items.Add(pipeSectionS[i].pipelineName);
                //pipelineSection_Check.Items.Add(pipeSectionS[i].pipelineSection);
            }
        }
        checkedItem CheckedItem = new checkedItem();
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
                AnomalyLogLine.featuresOrientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number7].Text);//ориентация
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
                AnomalyLogLine.defectRepareDate = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number19].Text);//Примечание
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
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Выполняется чтение данных о газопроводе...")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "->*")));
            //richTextBox1.AppendText(Environment.NewLine + "Выполняется чтение данных о газопроводе...");
            //richTextBox1.AppendText(Environment.NewLine + "->*");
            //int pipeListCount = Convert.ToInt16(textBox95.Text);//получаем длину журнала из формы
            pipelineInfo PipelineInfo = new pipelineInfo();

            PipelineInfo.pipelineName = ObjWorkSheet.Cells[NumbersOfColumns.stringNumber1, 4].Text;//трубопровод (название)
            PipelineInfo.pipelineSection = ObjWorkSheet.Cells[NumbersOfColumns.stringNumber2, 4].Text;//участок трубопровода
            String txt = ObjWorkSheet.Cells[NumbersOfColumns.stringNumber3, 4].Text;
            PipelineInfo.pipeDiameter = Convert.ToDouble(txt.Replace(".", ","));//диаметр трубы
            PipelineInfo.principal = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber4, 4].Text);//принципал (хозяин трубы)
            PipelineInfo.examinationDate = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber5, 4].Text);//дата обследования
            txt = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber6, 4].Text);
            PipelineInfo.designPressure = Convert.ToDouble(txt.Replace(".", ","));// проектное давление
            txt = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber7, 4].Text);
            PipelineInfo.operatingPressure = Convert.ToDouble(txt.Replace(".", ","));// рабочее давление
            PipelineInfo.comissioningYear = Convert.ToString(ObjWorkSheet.Cells[NumbersOfColumns.stringNumber8, 4].Text);//год ввода в экспуатацию
            PipelineInfo.contractor = "АО \"Газприборавтоматикасервис\"";
            //mGVTD.pipelineInfo = PipelineInfo;

            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Сведения о газопроводе прочтены")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "==========================================")));

            //richTextBox1.AppendText(Environment.NewLine + "Сведения о газопроводе прочтены");
            //richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return PipelineInfo;
        }
        private pipelineInfo operatingReadToClassPipeInfoSOD()//ДЛЯ СОД!!!метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            pipelineInfo PipelineInfo = new pipelineInfo();


            PipelineInfo.operatingPressure = Convert.ToDouble(textBox147.Text);
            PipelineInfo.pipeDiameter = Convert.ToDouble(textBox146.Text);
            PipelineInfo.pipelineName = textBox230.Text;
            PipelineInfo.pipelineSection = textBox232.Text;
            PipelineInfo.comissioningYear = textBox383.Text;
            PipelineInfo.examinationDate = textBox382.Text;
            PipelineInfo.contractor = "АО \"Газпром оргэнергогаз\" филиал \"Саратоворгдиагностиика\"";
            richTextBox3.AppendText(Environment.NewLine + "================================================");
            richTextBox3.AppendText(Environment.NewLine + "Прочтены сведения о рабочем давлении и диаметре.");
            return PipelineInfo;
        }
        private pipelineInfo operatingReadToClassPipeInfoNPCVTD()//ДЛЯ НПЦВТД!!!метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            pipelineInfo PipelineInfo = new pipelineInfo();


            PipelineInfo.operatingPressure = Convert.ToDouble(textBox_pressureNPCVTD.Text);
            PipelineInfo.pipeDiameter = Convert.ToDouble(textBox_diameterNPCVTD.Text);
            PipelineInfo.pipelineName = textBox_nameNPCVTD.Text;
            PipelineInfo.pipelineSection = textBox_ploteNPCVTD.Text;
            PipelineInfo.examinationDate = textBox_dateNPCVTD.Text;
            PipelineInfo.contractor = "НПЦ \"Внутритрубная диагностика\"";
            richTextBox5.AppendText(Environment.NewLine + "================================================");
            richTextBox5.AppendText(Environment.NewLine + "Прочтены сведения о рабочем давлении и диаметре.");
            return PipelineInfo;
        }
        private pipelineInfo operatingReadToClassPipeInfoBHTTS()//ДЛЯ БХТТС!!!метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            pipelineInfo PipelineInfo = new pipelineInfo();


            PipelineInfo.operatingPressure = Convert.ToDouble(textBox_pressure_BHTTS.Text);
            PipelineInfo.pipeDiameter = Convert.ToDouble(textBox_diam_BHTTS.Text);
            PipelineInfo.pipelineName = textBox_pipeline_BHTTS.Text;
            PipelineInfo.pipelineSection = textBox_plot_BHTTS.Text;
            PipelineInfo.examinationDate = textBox_date_BHTTS.Text;
            PipelineInfo.contractor = "ЗАО \"Бейкер Хьюз Технологии и Трубопроводный Сервис\"";
            richTextBox6.AppendText(Environment.NewLine + "================================================");
            richTextBox6.AppendText(Environment.NewLine + "Прочтены сведения о рабочем давлении и диаметре.");
            return PipelineInfo;
        }
        private List<MGPipe> OperatingReadToClassPipeLog(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета ВТД трубного журнала
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

        private List<anomalyLogLine> OperatingReadToClassAnomalyLog(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета строк журнала аномалий
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
                AnomalyLogLine.featuresOrientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number7].Text);//ориентация


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
                AnomalyLogLine.defectRepareDate = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number19].Text);//Примечание
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
        private List<furnishingsLog> OperatingReadToClassFurnishingsLog(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета строк журнала элементов обустройства
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
        private List<furnishingsLog> OperatingReadToClassFurnishingsLogAutoFin(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета строк журнала элементов обустройства
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

            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала элементов обустройства...")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "->*")));

            //richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала элементов обустройства...");
            //richTextBox1.AppendText(Environment.NewLine + "->*");
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
                    //richTextBox1.AppendText("*");
                    richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "*")));
                    incrementor = 0;
                }
                i++;
            }
            //textBox111.Text = Convert.ToString(i);//записываем в поле номер последней строки.
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала элементов обустройства прочитан, количество строк:" + furnishingsLogS.Count)));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "==========================================")));

            //richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала элементов обустройства прочитан, количество строк:"+ furnishingsLogS.Count);
            //richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return furnishingsLogS;

        }
        private List<furnishingsLog> OperatingReadToClassFurnishingsLogAutoFinSOD(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета строк журнала элементов обустройства
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox231.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            List<furnishingsLog> furnishingsLogS = new List<furnishingsLog>();
            richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "Выполняется обработка журнала элементов обустройства...")));
            richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "->*")));
            //AppendText(Environment.NewLine + "Выполняется обработка журнала элементов обустройства...");
            //richTextBox3.AppendText(Environment.NewLine + "->*");
            //int pipeListCount = Convert.ToInt16(textBox111.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора

            int i = 1;
            bool mark = true;
            while (mark)
            {
                furnishingsLog FurnishingsLog = new furnishingsLog();//создаём экземпляр класса строки журнала аномалий

                //FurnishingsLog.itemNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number1].Text);//номер пункта
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

                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number4].Text);
                try
                {
                    FurnishingsLog.pipeLength = Convert.ToDouble(txt.Replace(".", ","));//длина

                }
                catch (Exception)
                {
                    FurnishingsLog.pipeLength = 0;
                }*/
                //FurnishingsLog.pipeLength = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number4].Text);//длина трубы
                // FurnishingsLog.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number5].Text);//расстояние от поперечного шва, м
                FurnishingsLog.characterFeatures = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number6].Text);// характер особенности
                FurnishingsLog.designations = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number7].Text);//обозначение
                //FurnishingsLog.marker = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number8].Text);//маркер

                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number9].Text);
                try
                {
                    FurnishingsLog.distanceToNextFeature = Convert.ToDouble(txt.Replace(".", ","));//длина

                }
                catch (Exception)
                {
                    FurnishingsLog.distanceToNextFeature = 0;
                }*/
                //FurnishingsLog.distanceToNextFeature = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number9].Text);//расстояние до седующей особенности
                //FurnishingsLog.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number10].Text);//Широта
                //FurnishingsLog.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number11].Text);//Долгота

                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number12].Text);
                try
                {
                    FurnishingsLog.heightAboveSeaLevel = Convert.ToDouble(txt.Replace(".", ","));//длина

                }
                catch (Exception)
                {
                    FurnishingsLog.heightAboveSeaLevel = 0;
                }*/
                //FurnishingsLog.heightAboveSeaLevel = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number12].Text);//H, м
                FurnishingsLog.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number13].Text);//Примечание


                if (String.IsNullOrWhiteSpace(FurnishingsLog.pipeNumber))
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
                    richTextBox3.Invoke(new Action(() => richTextBox3.AppendText("*")));
                    //richTextBox1.AppendText("*");
                    incrementor = 0;
                }
                i++;
            }
            //textBox111.Text = Convert.ToString(i);//записываем в поле номер последней строки
            richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "Массив данных из журнала элементов обустройства прочитан, количество строк:" + furnishingsLogS.Count)));
            richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "==========================================")));

            //richTextBox3.AppendText(Environment.NewLine + "Массив данных из журнала элементов обустройства прочитан, количество строк:" + furnishingsLogS.Count);
            //richTextBox3.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return furnishingsLogS;

        }
        private List<furnishingsLog> OperatingReadToClassFurnishingsLogAutoFinNPCVTD(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета строк журнала элементов обустройства
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox316.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            List<furnishingsLog> furnishingsLogS = new List<furnishingsLog>();

            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "Выполняется обработка журнала элементов обустройства...")));
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "->*")));
            //richTextBox5.AppendText(Environment.NewLine + "Выполняется обработка журнала элементов обустройства...");
            //richTextBox5.AppendText(Environment.NewLine + "->*");
            //int pipeListCount = Convert.ToInt16(textBox111.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора

            int i = 4;
            bool mark = true;
            while (mark)
            {
                furnishingsLog FurnishingsLog = new furnishingsLog();//создаём экземпляр класса строки журнала аномалий

                //FurnishingsLog.itemNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number1].Text);//номер пункта
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

                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number4].Text);
                try
                {
                    FurnishingsLog.pipeLength = Convert.ToDouble(txt.Replace(".", ","));//длина

                }
                catch (Exception)
                {
                    FurnishingsLog.pipeLength = 0;
                }*/
                //FurnishingsLog.pipeLength = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number4].Text);//длина трубы
                // FurnishingsLog.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number5].Text);//расстояние от поперечного шва, м
                FurnishingsLog.characterFeatures = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number6].Text);// характер особенности
                FurnishingsLog.designations = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number7].Text);//обозначение
                //FurnishingsLog.marker = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number8].Text);//маркер

                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number9].Text);
                try
                {
                    FurnishingsLog.distanceToNextFeature = Convert.ToDouble(txt.Replace(".", ","));//длина

                }
                catch (Exception)
                {
                    FurnishingsLog.distanceToNextFeature = 0;
                }*/
                //FurnishingsLog.distanceToNextFeature = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number9].Text);//расстояние до седующей особенности
                //FurnishingsLog.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number10].Text);//Широта
                //FurnishingsLog.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number11].Text);//Долгота

                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number12].Text);
                try
                {
                    FurnishingsLog.heightAboveSeaLevel = Convert.ToDouble(txt.Replace(".", ","));//длина

                }
                catch (Exception)
                {
                    FurnishingsLog.heightAboveSeaLevel = 0;
                }*/
                //FurnishingsLog.heightAboveSeaLevel = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number12].Text);//H, м
                FurnishingsLog.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column3Number13].Text);//Примечание


                if (String.IsNullOrWhiteSpace(FurnishingsLog.pipeNumber))
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
                    richTextBox5.Invoke(new Action(() => richTextBox5.AppendText("*")));
                    incrementor = 0;
                }
                i++;
            }
            //textBox111.Text = Convert.ToString(i);//записываем в поле номер последней строки
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "Массив данных из журнала элементов обустройства прочитан, количество строк:" + furnishingsLogS.Count)));
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "==========================================")));
            ObjExcel.Quit();
            return furnishingsLogS;

        }
        private List<furnishingsLog> FirnishingLogVirtual(MGVTD mGVTD)
        {
            List<furnishingsLog> furnishingsLogS = new List<furnishingsLog>();
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "Журнал элементов обустройства не загружен. Журнал элементов обустройства будет заполнен на основе сведений из трубного журнала...")));
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "Журнал элементов обустройства не загружен. Журнал элементов обустройства будет заполнен на основе сведений из трубного журнала...")));
            richTextBox5.Invoke(new Action(() => richTextBox6.AppendText(Environment.NewLine + "->*")));
            richTextBox5.AppendText(Environment.NewLine + "Журнал элементов обустройства не загружен. Журнал элементов обустройства будет заполнен на основе сведений из трубного журнала...");
            richTextBox6.AppendText(Environment.NewLine + "Журнал элементов обустройства будет заполнен на основе сведений из трубного журнала...");
            richTextBox6.AppendText(Environment.NewLine + "->*");

            for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
            {
                if (String.IsNullOrWhiteSpace(mGVTD.MGPipeS[j].note))
                {
                    furnishingsLog FurnishingsLog = new furnishingsLog();
                    FurnishingsLog.pipeNumber = mGVTD.MGPipeS[j].pipeNumber;//номер трубы
                    FurnishingsLog.odometrDist = mGVTD.MGPipeS[j].odometrDist;//дистанция по одометру
                    FurnishingsLog.characterFeatures = mGVTD.MGPipeS[j].characterFeatures;// характер особенности
                    //richTextBox6.AppendText(Environment.NewLine + FurnishingsLog.characterFeatures);

                    //FurnishingsLog.designations = mGVTD.MGPipeS[j].designations;//обозначение
                    FurnishingsLog.note = mGVTD.MGPipeS[j].note;//Примечание
                    furnishingsLogS.Add(FurnishingsLog);

                }
            }
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "Журнал элементов обустройства заполнен, количество строк:" + furnishingsLogS.Count)));
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "==========================================")));
            richTextBox6.Invoke(new Action(() => richTextBox6.AppendText(Environment.NewLine + "Журнал элементов обустройства заполнен, количество строк:" + furnishingsLogS.Count)));
            richTextBox6.Invoke(new Action(() => richTextBox6.AppendText(Environment.NewLine + "==========================================")));


            return furnishingsLogS;

        }

        private List<pipeCharacteristics> OperatingReadToClassPipeCharacteristics(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета строк журнала элементов обустройства
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

            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала характеристик труб...")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "->*")));

            //richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала характеристик труб...");
            //richTextBox1.AppendText(Environment.NewLine + "->*");
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
                    // richTextBox1.AppendText("*");
                    richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "*")));
                    incrementor = 0;
                }

            }

            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала характеристик труб прочитан. Количество строк:" + pipeCharacteristicseS.Count)));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "==========================================")));

            //richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала характеристик труб прочитан. Количество строк:"+ pipeCharacteristicseS.Count);
            //richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return pipeCharacteristicseS;
        }
        private List<pipelineSectionCategoryLog> OperatingReadToClassPipelineSectionCategoryLog(string fileName, numbersOfColumns NumbersOfColumns)//!!!метод для чтения из файла отчета строк журнала элементов обустройства
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

            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала категорий участков...")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "->*")));

            //richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала категорий участков...");
            //richTextBox1.AppendText(Environment.NewLine + "->*");
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
                    //richTextBox1.AppendText("*");
                    richTextBox1.Invoke(new Action(() => richTextBox1.AppendText("*")));
                    incrementor = 0;
                }

            }
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала категорий учавстков прочитан. Количество строк:" + pipelineSectionCategoryLogS.Count)));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "==========================================")));

            //richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала категорий учавстков прочитан. Количество строк:"+ pipelineSectionCategoryLogS.Count);
            //richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return pipelineSectionCategoryLogS;

        }
        //********************************************************************************
        private List<MGPipe> ShortOperatingReadToClassPipeLog(string fileName, numbersOfColumns NumbersOfColumns)//КОРОТКИЙ!!!метод для чтения из файла отчета ВТД информации о трубопроводе
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
            ObjExcel.Quit();
            return OMGPipeS;
        }
        private List<MGPipe> ShortOperatingReadToClassPipeLogAutoFin(string fileName, numbersOfColumns NumbersOfColumns)//с автофинишем/КОРОТКИЙ!!!метод для чтения из файла отчета ВТД информации о трубопроводе
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
        private List<MGPipe> ShortOperatingReadToClassPipeLogAutoFinSOD(string fileName, numbersOfColumns NumbersOfColumns)//с автофинишем/КОРОТКИЙ!!!метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox170.Text;//получаем название вкладки из формы импотра (трубный журнал)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];


            richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "Выполняется обработка трубного журнала...")));
            richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "->*")));

            //richTextBox3.AppendText(Environment.NewLine + "Выполняется обработка трубного журнала...");
            //richTextBox3.AppendText(Environment.NewLine + "->*");
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
                    mGPipe.odometrDist = Convert.ToDouble(txt.Trim().Replace(".", ","));
                    //richTextBox1.AppendText(Environment.NewLine + "$"+ mGPipe.odometrDist);
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.odometrDist = 0;
                    richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "^")));
                    //richTextBox3.AppendText(Environment.NewLine + "^");
                }


                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber3].Text);
                try
                {
                    mGPipe.thikness = Convert.ToDouble(txt.Trim().Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.thikness = 0;
                }

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber4].Text);

                try
                {
                    mGPipe.pipeLength = Convert.ToDouble(txt.Trim().Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.pipeLength = 0;
                }
                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber7].Text);

                mGPipe.clockOrientation = txt.Replace(".", ",");

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

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber7].Text);
                try
                {
                    mGPipe.jointAngle = Convert.ToDouble(txt.Trim().Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.jointAngle = 0;
                }

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
                    richTextBox3.Invoke(new Action(() => richTextBox3.AppendText("*")));
                    //richTextBox3.AppendText("*");
                    incrementor = 0;
                }
                i++;
            }
            textBox95.Invoke(new Action(() => textBox95.Text = Convert.ToString(i)));
            //textBox95.Text = Convert.ToString(i);//записываем в поле количество труб
            richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "Массив данных из трубного журнала прочитан, количество труб: " + OMGPipeS.Count)));

            //richTextBox3.AppendText(Environment.NewLine + "Массив данных из трубного журнала прочитан, количество труб: " + OMGPipeS.Count);

            richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "==========================================")));
            //richTextBox3.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return OMGPipeS;
        }
        private List<MGPipe> ShortOperatingReadToClassPipeLogAutoFinNPCVTD(string fileName, numbersOfColumns NumbersOfColumns)//с автофинишем/КОРОТКИЙ!!!метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox250.Text;//получаем название вкладки из формы импотра (трубный журнал)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "Выполняется обработка трубного журнала...")));

            //richTextBox5.AppendText(Environment.NewLine + "Выполняется обработка трубного журнала...");
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "->*")));
            //int pipeListCount = Convert.ToInt16(textBox95.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            List<MGPipe> OMGPipeS = new List<MGPipe>();//трубный журнал

            int i = 4;
            bool mark = true;
            while (mark)
            {
                MGPipe mGPipe = new MGPipe();
                mGPipe.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber1].Text);

                String txt;
                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber2].Text);
                try
                {
                    mGPipe.odometrDist = Convert.ToDouble(txt.Trim().Replace(".", ","));
                    //richTextBox1.AppendText(Environment.NewLine + "$"+ mGPipe.odometrDist);
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.odometrDist = 0;

                }


                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber3].Text);
                try
                {
                    mGPipe.thikness = Convert.ToDouble(txt.Trim().Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.thikness = 0;
                }

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber4].Text);
                try
                {
                    mGPipe.pipeLength = Convert.ToDouble(txt.Trim().Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.pipeLength = 0;
                }

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber16].Text);
                try
                {
                    mGPipe.yieldPoint = Convert.ToDouble(txt.Trim().Replace(".", ","));
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    mGPipe.yieldPoint = 0;
                }

                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber16].Text);

                //mGPipe.tensileStrength =Double.Parse( txt.Replace(".", ","));

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

                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber7].Text);
                try
                {
                    mGPipe.jointAngle = Convert.ToDouble(txt.Trim().Replace(".", ","));
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

                //mGPipe.steelGrade = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.columnNumber17].Text);

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
                    richTextBox5.Invoke(new Action(() => richTextBox5.AppendText("*")));
                    incrementor = 0;
                }
                i++;
            }


            //textBox95.Text = Convert.ToString(i);//записываем в поле количество труб
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "Массив данных из трубного журнала прочитан, количество труб: " + OMGPipeS.Count)));
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "==========================================")));
            //richTextBox5.AppendText(Environment.NewLine + "Массив данных из трубного журнала прочитан, количество труб: " + OMGPipeS.Count);
            //richTextBox5.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return OMGPipeS;
        }

        private MGVTD GetCritikalThiknessForAll(MGVTD mGVTD)
        {
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                mGVTD.MGPipeS[i].critikalThikness = GetCritikalThikness(mGVTD.pipelineInfo.operatingPressure, mGVTD.pipelineInfo.pipeDiameter, mGVTD.MGPipeS[i].tensileStrength, Convert.ToInt32(mGVTD.MGPipeS[i].pipelineSectionCategory));
                //richTextBox7.AppendText(Environment.NewLine + "CritikalThikness " + mGVTD.MGPipeS[i].critikalThikness+"_"+ mGVTD.pipelineInfo.operatingPressure + "Diam " + mGVTD.pipelineInfo.pipeDiameter+"_"+ mGVTD.MGPipeS[i].tensileStrength+"_"+ mGVTD.MGPipeS[i].pipelineSectionCategory);
                double Vcorr = (0.01 * mGVTD.MGPipeS[i].MaximumCorrProcent * mGVTD.MGPipeS[i].thikness) / (DateTime.Now.Year - Convert.ToDouble(mGVTD.pipelineInfo.comissioningYear));
                double minimumThikness = mGVTD.MGPipeS[i].thikness - 0.01 * mGVTD.MGPipeS[i].MaximumCorrProcent * mGVTD.MGPipeS[i].thikness;
                //richTextBox7.AppendText(Environment.NewLine + "Vcorr " + Vcorr + "MaxCorrProc " + mGVTD.MGPipeS[i].MaximumCorrProcent+" thikn"+ mGVTD.MGPipeS[i].thikness);
                if (Vcorr > 0)
                {
                    mGVTD.MGPipeS[i].residualResource = (minimumThikness - mGVTD.MGPipeS[i].critikalThikness) / Vcorr;
                }
                else
                {
                    mGVTD.MGPipeS[i].residualResource = 20;
                }
                if (mGVTD.MGPipeS[i].residualResource < 10)
                {
                    richTextBox7.AppendText(Environment.NewLine + "Труба № " + mGVTD.MGPipeS[i].pipeNumber + " имеет ост. ресурс: " + Math.Round(mGVTD.MGPipeS[i].residualResource, 1) + " лет. Макс. глуб. коррозии: " + mGVTD.MGPipeS[i].MaximumCorrProcent + " %");
                }
            }
            return mGVTD;
        }
        private double GetCritikalThikness(double p, double Dh, double Ri, int kategory)//давление, диаметр, сопр. разр., категория участка
        {
            double result = 0;
            double n = 1.2;//коэфф. над. по нагр
            double k1 = 1.1;//Коэфф. над. по метериалу
            double m = 0.825;//коэфф.
            double kn = 1.1;//коэфф. над. по назначению
            if (kategory == 1 | kategory == 2)
            {
                m = 0.825;
            }
            else if (kategory == 3 | kategory == 4)
            {
                m = 0.99;
            }
            else
            {
                m = 0.99;
            }
            if (Dh < 1000)
            {
                kn = 1.1;
            }
            else if (Dh > 1000 & Dh < 1200)
            {
                kn = 1.155;
            }
            else if (Dh > 1200)
            {
                kn = 1.21;
            }
            else
            {
                kn = 1.21;
            }

            //result = (n*p*Dh) / (2*((Ri*m)/((k1+kn)+n*p)));
            //result = n * p * Dh / (2 * (Ri * m) / (k1 * kn) + n * p);
            result = (n * p * Dh) / (2 * (Ri + n * p));
            //double R = (Ri * m) / (k1 * kn);
            //result = (n * p * Dh) / (2 * (R + n * p));
            return result;
        }
        private MGVTD OperatingReadToClassPipeLogHimself(string fileName, string worksheetName)//для БХТТС с автофинишем/КОРОТКИЙ!!!метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = worksheetName;//получаем название вкладки из формы импотра (трубный журнал) "SonarFormat"
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];
            MGVTD result = new MGVTD();
            richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + "Выполняется обработка трубного журнала...")));
            richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + "->*")));
            //richTextBox7.Invoke(new Action(() =>));
            //int pipeListCount = Convert.ToInt16(textBox95.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            List<MGPipe> MGPipes = new List<MGPipe>();//трубный журнал
            List<anomalyLogLine> anomalyLog = new List<anomalyLogLine>();
            List<furnishingsLog> furnishingsLog = new List<furnishingsLog>();


            pipelineInfo info = new pipelineInfo();
            info.comissioningYear = Convert.ToString(ObjWorkSheet2.Cells[1, 28].Text);
            info.contractor = Convert.ToString(ObjWorkSheet2.Cells[1, 31].Text);
            info.designPressure = ConvertToDouble(ObjWorkSheet2.Cells[1, 26].Text);
            info.examinationDate = Convert.ToString(ObjWorkSheet2.Cells[1, 25].Text);
            info.operatingPressure = ConvertToDouble(ObjWorkSheet2.Cells[1, 27].Text);
            info.pipeDiameter = ConvertToDouble(ObjWorkSheet2.Cells[1, 24].Text);
            info.pipelineName = Convert.ToString(ObjWorkSheet2.Cells[1, 22].Text);
            info.pipelineSection = Convert.ToString(ObjWorkSheet2.Cells[1, 23].Text);
            result.pipelineInfo = info;

            int i = 2;
            bool mark = true;
            while (mark)
            {



                MGPipe MG_Pipe = new MGPipe();
                anomalyLogLine anomalyLog_Line = new anomalyLogLine();
                furnishingsLog furnishings_Log = new furnishingsLog();
                //MG_Pipe.pipelineSectionCategory = "1";
                //MG_Pipe.tensileStrength = 550;
                string featuresNumber_BHTTS = Convert.ToString(ObjWorkSheet2.Cells[i, 1].Text);//Номер особенности///1
                string pipeNumber_BHTTS = Convert.ToString(ObjWorkSheet2.Cells[i, 2].Text);//номер трубы///2

                if (String.IsNullOrEmpty(Convert.ToString(ObjWorkSheet2.Cells[i, 1].Text)) & String.IsNullOrEmpty(Convert.ToString(ObjWorkSheet2.Cells[i, 2].Text)) == false)//если "Номер особенности" пустой, значит строка содержит сведения о трубе
                {
                    //MG_Pipe.pipeID = ConvertToInt(ObjWorkSheet2.Cells[i, 2].Text);//
                    MG_Pipe.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i, 2].Text);//номер трубы
                    MG_Pipe.odometrDist = ConvertToDouble(ObjWorkSheet2.Cells[i, 3].Text);//дистанция по одометру
                    MG_Pipe.thikness = ConvertToDouble(ObjWorkSheet2.Cells[i, 7].Text);//толщина трубы
                    MG_Pipe.pipeLength = ConvertToDouble(ObjWorkSheet2.Cells[i, 6].Text);//длина трубы
                    MG_Pipe.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i, 15].Text);//расстояние от реперных точек
                    MG_Pipe.characterFeatures = Convert.ToString(ObjWorkSheet2.Cells[i, 5].Text);// характер особенности
                    MG_Pipe.clockOrientation = Convert.ToString(ObjWorkSheet2.Cells[i, 11].Text);//Ориент., ч:мин
                    //MG_Pipe.bendOfPipe = Convert.ToString(ObjWorkSheet2.Cells[i, 2].Text);//Изгиб, °
                    //MG_Pipe.jointAngle = Convert.ToString(ObjWorkSheet2.Cells[i, 2].Text);//Угол стыка,°
                    MG_Pipe.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i, 25].Text);//Широта
                    MG_Pipe.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i, 26].Text);//Долгота
                    //richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + MG_Pipe.Latitude+"_"+ MG_Pipe.Longitude)));
                    //MG_Pipe.heightAboveSeaLevel = Convert.ToString(ObjWorkSheet2.Cells[i, 2].Text);//H, м
                    MG_Pipe.note = Convert.ToString(ObjWorkSheet2.Cells[i, 14].Text);//Примечание

                    //Следующие поля заполняются после обработки отчета
                    MG_Pipe.pipelineSectionCategory = Convert.ToString(ObjWorkSheet2.Cells[i, 16].Text);//!!!категория участка трубопровода - заполняется при обработке массива
                    MG_Pipe.steelGrade = Convert.ToString(ObjWorkSheet2.Cells[i, 17].Text);//!!!марка стали - заполняется при обработке массива
                    MG_Pipe.yieldPoint = ConvertToDouble(ObjWorkSheet2.Cells[i, 18].Text);//!!!предел текучести - заполняется при обработке массива
                    MG_Pipe.tensileStrength = ConvertToDouble(ObjWorkSheet2.Cells[i, 19].Text);//!!!предел прочности - заполняется при обработке массива
                    MGPipes.Add(MG_Pipe);
                    if (String.IsNullOrEmpty(Convert.ToString(ObjWorkSheet2.Cells[i, 14].Text)) == false)//Запись для журнала элементоов обустройства
                    {
                        //furnishings_Log.itemNumber= Convert.ToString(ObjWorkSheet2.Cells[i, 17].Text);//номер пункта
                        furnishings_Log.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i, 2].Text);//номер трубы
                        furnishings_Log.odometrDist = ConvertToDouble(ObjWorkSheet2.Cells[i, 3].Text);//дистанция по одометру 
                        furnishings_Log.pipeLength = ConvertToDouble(ObjWorkSheet2.Cells[i, 6].Text);//длина трубы
                        furnishings_Log.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i, 15].Text);//расстояние от поперечного шва, м
                        furnishings_Log.characterFeatures = Convert.ToString(ObjWorkSheet2.Cells[i, 5].Text);// характер особенности
                        furnishings_Log.designations = Convert.ToString(ObjWorkSheet2.Cells[i, 14].Text);//обозначение
                        furnishings_Log.marker = Convert.ToString(ObjWorkSheet2.Cells[i, 14].Text);//маркер
                        //furnishings_Log.distanceToNextFeature;//расстояние до седующей особенности
                        furnishings_Log.Latitude=Convert.ToString(ObjWorkSheet2.Cells[i, 25].Text); //Широта
                        furnishings_Log.Longitude=Convert.ToString(ObjWorkSheet2.Cells[i, 26].Text);//Долгота
                        //richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + furnishings_Log.Latitude + "_" + furnishings_Log.Longitude+"_line")));
                        //furnishings_Log.heightAboveSeaLevel;//H, м
                        furnishings_Log.note = Convert.ToString(ObjWorkSheet2.Cells[i, 14].Text);//Примечание
                        furnishingsLog.Add(furnishings_Log);
                    }
                }
                else if (String.IsNullOrEmpty(Convert.ToString(ObjWorkSheet2.Cells[i, 1].Text)) == false & String.IsNullOrEmpty(Convert.ToString(ObjWorkSheet2.Cells[i, 2].Text)) == false)//значит это дефект
                {
                    anomalyLog_Line.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i, 2].Text);//номер трубы
                    anomalyLog_Line.odometrDist = ConvertToDouble(ObjWorkSheet2.Cells[i, 3].Text);//дистанция по одометру
                    anomalyLog_Line.thikness = ConvertToDouble(ObjWorkSheet2.Cells[i, 7].Text);//толщина трубы
                    anomalyLog_Line.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i, 4].Text);//расстояние от поперечного шва, м
                    anomalyLog_Line.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i, 15].Text);//расстояние от реперных точек
                    anomalyLog_Line.featuresCharacter = Convert.ToString(ObjWorkSheet2.Cells[i, 5].Text);//характер особенности
                    //anomalyLog_Line.classOfSize = Convert.ToString(ObjWorkSheet2.Cells[i, 17].Text);//класс размера
                    anomalyLog_Line.featuresOrientation = Convert.ToString(ObjWorkSheet2.Cells[i, 11].Text);//ориентация
                    anomalyLog_Line.length = ConvertToDouble(ObjWorkSheet2.Cells[i, 9].Text);//длина
                    anomalyLog_Line.widht = ConvertToDouble(ObjWorkSheet2.Cells[i, 10].Text);//ширина
                    anomalyLog_Line.depthInProcent = ConvertToDouble(ObjWorkSheet2.Cells[i, 8].Text);//глубина дефекта в процентах
                    anomalyLog_Line.depthInMm = ConvertToDouble(ObjWorkSheet2.Cells[i, 21].Text);//глубина дефекта в миллиметрах
                    //anomalyLog_Line.extOrInt = ConvertToDouble(ObjWorkSheet2.Cells[i, 17].Text);//характер локаизации(внутри или снаружи)
                    //anomalyLog_Line.KBD = Convert.ToString(ObjWorkSheet2.Cells[i, 17].Text);//КБД
                    anomalyLog_Line.defectAssessment = Convert.ToString(ObjWorkSheet2.Cells[i, 13].Text);//оценка дефекта
                    anomalyLog_Line.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i, 25].Text); //Широта
                    anomalyLog_Line.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i, 26].Text);//Долгота
                    //richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + anomalyLog_Line.Latitude + "_" + anomalyLog_Line.Longitude + "_def")));
                    //anomalyLog_Line.heightAboveSeaLevel = Convert.ToString(ObjWorkSheet2.Cells[i, 17].Text);//H, м
                    anomalyLog_Line.note = Convert.ToString(ObjWorkSheet2.Cells[i, 14].Text);//Примечание
                    anomalyLog_Line.defectRepareDate = Convert.ToString(ObjWorkSheet2.Cells[i, 20].Text);//дата устранения дефекта
                    anomalyLog.Add(anomalyLog_Line);
                }

                if (String.IsNullOrWhiteSpace(featuresNumber_BHTTS) & String.IsNullOrWhiteSpace(pipeNumber_BHTTS))
                {
                    mark = false;//дошли до конца трубного журлала
                    result.MGPipeS = MGPipes;
                    result.anomalyLogLineS = anomalyLog;
                    result.furnishingsLogS = furnishingsLog;
                }

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor == 100)
                {

                    richTextBox7.Invoke(new Action(() => richTextBox7.AppendText("*")));
                    incrementor = 0;
                }
                i++;
            }

            textBox95.Text = Convert.ToString(i);//записываем в поле количество труб
            richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + "Массив данных из трубного журнала прочитан, количество труб: " + result.MGPipeS.Count)));
            richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + "Количество строк дефектной ведомости: " + result.anomalyLogLineS.Count)));
            richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + "Количество строк журнала линейных объектов: " + result.furnishingsLogS.Count)));
            richTextBox7.Invoke(new Action(() => richTextBox7.AppendText(Environment.NewLine + "==========================================")));

            ObjExcel.Quit();
            return result;
        }
        private MGVTD OperatingReadToClassPipeLogAutoFinBHTTS(string fileName, numbersOfColumns NumbersOfColumns)// с автофинишем. метод для чтения из файла отчета ВТД информации о трубопроводе
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox379.Text;//задаём название вкладки
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];
            MGVTD result = new MGVTD();
            richTextBox6.AppendText(Environment.NewLine + "Выполняется обработка трубного журнала...");
            richTextBox6.AppendText(Environment.NewLine + "->*");
            //int pipeListCount = Convert.ToInt16(textBox95.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            List<MGPipe> MGPipes = new List<MGPipe>();//трубный журнал
            List<anomalyLogLine> anomalyLog = new List<anomalyLogLine>();
            List<furnishingsLog> furnishingsLog = new List<furnishingsLog>();


            int i = 4;
            bool mark = true;
            double odometr_old = 0;
            while (mark)
            {
                BHTTS_pipelog_String BHTTS_pipelog_string = new BHTTS_pipelog_String();//создаём список для хранения строк трубного журнала БХТТС
                MGPipe MG_Pipe = new MGPipe();
                anomalyLogLine anomalyLog_Line = new anomalyLogLine();
                furnishingsLog furnishings_Log = new furnishingsLog();
                MG_Pipe.pipelineSectionCategory = "1";
                MG_Pipe.tensileStrength = 550;
                string featuresNumber_BHTTS = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.featuresNumber_BHTTS].Text);//Номер особенности///1
                string pipeNumber_BHTTS = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.pipeNumber_BHTTS].Text);//номер трубы///2

                BHTTS_pipelog_string.featuresNumber_BHTTS = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.featuresNumber_BHTTS].Text);//Номер особенности///1
                BHTTS_pipelog_string.pipeNumber_BHTTS = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.pipeNumber_BHTTS].Text);//номер трубы///2
                BHTTS_pipelog_string.odometrDist_BHTTS = ConvertToDouble(ObjWorkSheet2.Cells[i, NumbersOfColumns.odometrDist_BHTTS].Text);//дистанция по одометру///3
                BHTTS_pipelog_string.distanceFromReferencePoints_BHTTS = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.distanceFromReferencePoints_BHTTS].Text);//расстояние от реперных точек///4
                BHTTS_pipelog_string.distanceToNextReferencePoints_BHTTS = ConvertToDouble(ObjWorkSheet2.Cells[i, NumbersOfColumns.distanceToNextReferencePoints_BHTTS].Text);//расстояние до следующей реперной точки///5
                BHTTS_pipelog_string.featuresCharacter_BHTTS = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.featuresCharacter_BHTTS].Text);//характер особенности///6
                BHTTS_pipelog_string.distanceFromTransverseWeld_BHTTS = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.distanceFromTransverseWeld_BHTTS].Text);//расстояние от поперечного шва, м///7
                BHTTS_pipelog_string.featuresOrientation_BHTTS = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.featuresOrientation_BHTTS].Text);//угловая ориентация///8
                BHTTS_pipelog_string.length_BHTTS = ConvertToDouble(ObjWorkSheet2.Cells[i, NumbersOfColumns.length_BHTTS].Text);//длина///9
                BHTTS_pipelog_string.widht_BHTTS = ConvertToDouble(ObjWorkSheet2.Cells[i, NumbersOfColumns.widht_BHTTS].Text);//ширина///10
                BHTTS_pipelog_string.thikness_BHTTS = ConvertToDouble(ObjWorkSheet2.Cells[i, NumbersOfColumns.thikness_BHTTS].Text);//толщина трубы///11
                BHTTS_pipelog_string.depthInProcent_BHTTS = ConvertToDouble(ObjWorkSheet2.Cells[i, NumbersOfColumns.depthInProcent_BHTTS].Text);//глубина дефекта в процентах///12
                BHTTS_pipelog_string.extOrInt_BHTTS = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.extOrInt_BHTTS].Text);//характер локаизации(внутри или снаружи)///13
                BHTTS_pipelog_string.note_BHTTS = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.note_BHTTS].Text);//Примечание///14
                BHTTS_pipelog_string.defectVanishDate = Convert.ToString(ObjWorkSheet2.Cells[i, NumbersOfColumns.defectVanishDate].Text);//Примечание///15
                if (String.IsNullOrWhiteSpace(featuresNumber_BHTTS) & String.IsNullOrWhiteSpace(pipeNumber_BHTTS))
                {
                    mark = false;//дошли до конца трубного журлала
                    result.MGPipeS = MGPipes;
                    result.anomalyLogLineS = anomalyLog;
                }
                else
                {


                    if (String.IsNullOrWhiteSpace(featuresNumber_BHTTS) == false & String.IsNullOrWhiteSpace(pipeNumber_BHTTS) == false)//значит это дефект, и строка идет в журнал дефектов
                    {
                        anomalyLog_Line.pipeNumber = BHTTS_pipelog_string.pipeNumber_BHTTS;//номер трубы
                        anomalyLog_Line.odometrDist = BHTTS_pipelog_string.odometrDist_BHTTS;//дистанция по одометру
                        anomalyLog_Line.thikness = BHTTS_pipelog_string.thikness_BHTTS;//толщина трубы
                        anomalyLog_Line.distanceFromTransverseWeld = BHTTS_pipelog_string.distanceFromTransverseWeld_BHTTS;//расстояние от поперечного шва, м
                        anomalyLog_Line.distanceFromReferencePoints = BHTTS_pipelog_string.distanceFromReferencePoints_BHTTS;//расстояние от реперных точек
                        anomalyLog_Line.featuresCharacter = BHTTS_pipelog_string.featuresCharacter_BHTTS;//характер особенности
                        if (anomalyLog_Line.featuresCharacter.Contains("отер"))
                        {
                            anomalyLog_Line.featuresCharacter = anomalyLog_Line.featuresCharacter + "(коррозия)";
                        }
                        if (anomalyLog_Line.featuresCharacter.Contains("рматур"))
                        {
                            anomalyLog_Line.featuresCharacter = anomalyLog_Line.featuresCharacter + "(кран)";
                            anomalyLog_Line.note = anomalyLog_Line.note + "(кран)";
                        }
                        anomalyLog_Line.classOfSize = "";//класс размера
                        anomalyLog_Line.featuresOrientation = BHTTS_pipelog_string.featuresOrientation_BHTTS;//ориентация
                        anomalyLog_Line.length = BHTTS_pipelog_string.length_BHTTS;//длина
                        anomalyLog_Line.widht = BHTTS_pipelog_string.widht_BHTTS;//ширина
                        anomalyLog_Line.depthInProcent = BHTTS_pipelog_string.depthInProcent_BHTTS;//глубина дефекта в процентах
                        anomalyLog_Line.depthInMm = (BHTTS_pipelog_string.depthInProcent_BHTTS * BHTTS_pipelog_string.thikness_BHTTS) / 100;//глубина дефекта в миллиметрах
                        anomalyLog_Line.extOrInt = BHTTS_pipelog_string.extOrInt_BHTTS;//характер локаизации(внутри или снаружи)
                        anomalyLog_Line.KBD = "";//КБД

                        if (BHTTS_pipelog_string.note_BHTTS.Contains(" А") | BHTTS_pipelog_string.note_BHTTS.Contains(" А"))
                        {
                            anomalyLog_Line.defectAssessment = "A";//оценка дефекта
                        }
                        else if (BHTTS_pipelog_string.note_BHTTS.Contains(" B") | BHTTS_pipelog_string.note_BHTTS.Contains(" В"))
                        {
                            anomalyLog_Line.defectAssessment = "B";//оценка дефекта
                        }
                        else if (BHTTS_pipelog_string.note_BHTTS.Contains(" C") | BHTTS_pipelog_string.note_BHTTS.Contains(" С"))
                        {
                            anomalyLog_Line.defectAssessment = "C";//оценка дефекта                            
                        }

                        anomalyLog_Line.Latitude = "";//Широта
                        anomalyLog_Line.Longitude = "";//Долгота
                        anomalyLog_Line.heightAboveSeaLevel = 0;//H, м
                        anomalyLog_Line.note = BHTTS_pipelog_string.note_BHTTS + " " + anomalyLog_Line.featuresCharacter;//Примечание

                        anomalyLog_Line.defectRepareDate = BHTTS_pipelog_string.defectVanishDate;//дата устранения дефекта

                        anomalyLog.Add(anomalyLog_Line);
                    }
                    else if (String.IsNullOrWhiteSpace(featuresNumber_BHTTS) & String.IsNullOrWhiteSpace(pipeNumber_BHTTS) == false)//значит это труба и строка идет в трубный журнал
                    {
                        MG_Pipe.pipeNumber = BHTTS_pipelog_string.pipeNumber_BHTTS;//номер трубы
                        MG_Pipe.odometrDist = BHTTS_pipelog_string.odometrDist_BHTTS;//дистанция по одометруMG_Pipe.
                        MG_Pipe.thikness = BHTTS_pipelog_string.thikness_BHTTS;//толщина трубыMG_Pipe.

                        if (i > 4)
                        {
                            MG_Pipe.pipeLength = MG_Pipe.odometrDist - odometr_old;//длина трубыMG_Pipe.
                        }
                        else
                        {
                            MG_Pipe.pipeLength = MG_Pipe.odometrDist;
                        }

                        MG_Pipe.distanceFromReferencePoints = BHTTS_pipelog_string.distanceFromReferencePoints_BHTTS;//расстояние от реперных точек
                        MG_Pipe.characterFeatures = BHTTS_pipelog_string.featuresCharacter_BHTTS;// характер особенности
                        MG_Pipe.clockOrientation = BHTTS_pipelog_string.featuresOrientation_BHTTS;//Ориент., ч:мин                        
                        MG_Pipe.note = BHTTS_pipelog_string.note_BHTTS;//Примечание
                        MG_Pipe.yieldPoint = 240;//справочное значение для ст.20
                        MG_Pipe.tensileStrength = 400;//справочное значение для ст.20
                        MG_Pipe.pipelineSectionCategory = "1";//в условиях отсутствия информации все участки для расчета принимаются по жесткому, как для первой категории
                        MGPipes.Add(MG_Pipe);
                        odometr_old = MG_Pipe.odometrDist;
                    }
                }

                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                if (incrementor == 100)
                {
                    richTextBox6.AppendText("*");
                    incrementor = 0;
                }
                i++;
            }

            textBox95.Text = Convert.ToString(i);//записываем в поле количество труб
            richTextBox6.AppendText(Environment.NewLine + "Массив данных из трубного журнала прочитан, количество труб: " + result.MGPipeS.Count);
            richTextBox6.AppendText(Environment.NewLine + "Количество строк дефектной ведомости: " + result.anomalyLogLineS.Count);
            richTextBox6.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return result;
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
                AnomalyLogLine.defectRepareDate = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number19].Text);//Примечание
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

            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала выявленных аномалий...")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "->*")));


            //richTextBox1.AppendText(Environment.NewLine + "Выполняется обработка журнала выявленных аномалий...");
            //richTextBox1.AppendText(Environment.NewLine + "->*");

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
                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number1].Text);
                try
                {
                    AnomalyLogLine.odometrDist = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.odometrDist = 0;
                }


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
                AnomalyLogLine.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number4].Text);//расстояние от реперных точек
                AnomalyLogLine.featuresCharacter = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number5].Text);//характер особенности
                //AnomalyLogLine.classOfSize = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number6].Text);//класс размера
                AnomalyLogLine.featuresOrientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number7].Text);//ориентация


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
                AnomalyLogLine.extOrInt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, 14].Text);//характер локаизации(внутри или снаружи)
                AnomalyLogLine.KBD = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number13].Text);//КБД
                AnomalyLogLine.defectAssessment = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number14].Text);//оценка дефекта
                AnomalyLogLine.Latitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number15].Text);//Широта
                AnomalyLogLine.Longitude = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number16].Text);//Долгота
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
                AnomalyLogLine.defectRepareDate = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number19].Text);//Примечание

                if (String.IsNullOrWhiteSpace(AnomalyLogLine.pipeNumber))
                {
                    mark = false;//дошли до конца трубного журлала
                }
                else
                {
                    if (AnomalyLogLine.featuresCharacter.Contains("Труб") == false)
                    {
                        anomalyLogLineS.Add(AnomalyLogLine);//добавляем заполненный экземпляр класса к списку
                    }

                }
                incrementor++;//сделаем прогресс-индикатор, чтобы было не так скучно ждать.
                i++;
                if (incrementor == 100)
                {
                    //richTextBox1.AppendText("*");
                    richTextBox1.Invoke(new Action(() => richTextBox1.AppendText("*")));
                    incrementor = 0;
                }


            }
            //textBox110.Text = Convert.ToString(i);//записываем в поле количество труб

            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала выявленных аномалий прочитан, количество дефектов:" + anomalyLogLineS.Count)));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "==========================================")));

            //richTextBox1.AppendText(Environment.NewLine + "Массив данных из журнала выявленных аномалий прочитан, количество дефектов:"+ anomalyLogLineS.Count);
            //richTextBox1.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return anomalyLogLineS;
        }

        private List<anomalyLogLine> shortOperatingReadToClassAnomalyLogAutoFinSOD(string fileName, numbersOfColumns NumbersOfColumns)//с автофинишем/КОРОТКИЙ!!!метод для чтения из файла отчета строк журнала аномалий
        {

            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox196.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "Выполняется обработка журнала выявленных аномалий...")));
            //richTextBox3.AppendText(Environment.NewLine + "Выполняется обработка журнала выявленных аномалий...");

            richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "->*")));
            //richTextBox3.AppendText(Environment.NewLine + "->*");

            List<anomalyLogLine> anomalyLogLineS = new List<anomalyLogLine>();
            int pipeListCount = Convert.ToInt16(textBox110.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            int i = 1;
            bool mark = true;
            while (mark)
            {
                anomalyLogLine AnomalyLogLine = new anomalyLogLine();//создаём экземпляр класса строки журнала аномалий
                AnomalyLogLine.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number20].Text);//расстояние от поперечного шва, м


                String txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number1].Text);
                try
                {
                    AnomalyLogLine.odometrDist = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.odometrDist = 0;
                }


                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number2].Text);
                try
                {
                    AnomalyLogLine.thikness = Convert.ToDouble(txt.Trim().Replace(".", ","));//длина

                }
                catch (Exception)
                {
                    AnomalyLogLine.thikness = 0;
                }*/

                AnomalyLogLine.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number3].Text);//расстояние от поперечного шва, м
                //AnomalyLogLine.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number4].Text);//расстояние от реперных точек
                AnomalyLogLine.featuresCharacter = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number5].Text);//характер особенности
                //AnomalyLogLine.classOfSize = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number6].Text);//класс размера
                AnomalyLogLine.featuresOrientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number7].Text);//ориентация
                txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number8].Text);
                try
                {
                    if (checkBox17.Checked)
                    {
                        AnomalyLogLine.length = 1000 * Convert.ToDouble(txt.Replace(".", ","));//длина
                    }
                    else
                    {
                        AnomalyLogLine.length = Convert.ToDouble(txt.Replace(".", ","));//длина    
                    }
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

                    if (checkBox17.Checked)
                    {
                        AnomalyLogLine.widht = 1000 * Convert.ToDouble(txt.Replace(".", ","));//ширина
                    }
                    else
                    {
                        AnomalyLogLine.widht = Convert.ToDouble(txt.Replace(".", ","));//ширина
                    }


                }
                catch (Exception)
                {
                    AnomalyLogLine.widht = 0;
                }

                //**********************************************************
                double pipeThikness = 0;
                for (int f = 0; f < mGVTD.MGPipeS.Count; f++)
                {
                    if (String.Equals(mGVTD.MGPipeS[f].pipeNumber, AnomalyLogLine.pipeNumber))
                    {
                        pipeThikness = mGVTD.MGPipeS[f].thikness;

                    }
                }


                String depthInMMString = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number11].Text);//NumbersOfColumns.column2Number11
                String depthInProcentString = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number10].Text);//NumbersOfColumns.column2Number10

                try
                {
                    AnomalyLogLine.depthInProcent = Convert.ToDouble(depthInProcentString.Trim().Replace(".", ","));//глубина дефекта в процентах

                }
                catch (Exception)
                {
                    AnomalyLogLine.depthInProcent = 0;

                }

                try
                {
                    AnomalyLogLine.depthInMm = 0.1*Convert.ToDouble(depthInMMString.Trim().Replace(".", ","));//глубина дефекта в миллиметрах

                }
                catch (Exception)
                {
                    AnomalyLogLine.depthInMm = 0;
                }

                if (AnomalyLogLine.depthInMm == 0)
                {
                    AnomalyLogLine.depthInMm = AnomalyLogLine.depthInProcent * pipeThikness * 0.001;

                }
                if (AnomalyLogLine.depthInProcent == 0)
                {
                    AnomalyLogLine.depthInProcent = AnomalyLogLine.depthInMm / (pipeThikness * 0.001);

                }
                //********************************************************

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
                //AnomalyLogLine.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number18].Text);//Примечание
                AnomalyLogLine.defectRepareDate = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number19].Text);//Примечание

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
                    richTextBox3.Invoke(new Action(() => richTextBox3.AppendText("*")));
                    //richTextBox3.AppendText("*");
                    incrementor = 0;
                }


            }
            textBox110.Invoke(new Action(() => textBox110.Text = Convert.ToString(i)));
            //textBox110.Text = Convert.ToString(i);//записываем в поле количество труб

            richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "Массив данных из журнала выявленных аномалий прочитан, количество дефектов:" + anomalyLogLineS.Count)));
            richTextBox3.Invoke(new Action(() => richTextBox3.AppendText(Environment.NewLine + "==========================================")));
            //richTextBox3.AppendText(Environment.NewLine + "Массив данных из журнала выявленных аномалий прочитан, количество дефектов:" + anomalyLogLineS.Count);
            //richTextBox3.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return anomalyLogLineS;
        }
        private List<anomalyLogLine> shortOperatingReadToClassAnomalyLogAutoFinNPCVTD(string fileName, numbersOfColumns NumbersOfColumns)//с автофинишем/КОРОТКИЙ!!!метод для чтения из файла отчета строк журнала аномалий
        {

            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileName, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName = textBox274.Text;//получаем название вкладки из формы импотра (журнал выявленных аномалий)
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName];

            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "Выполняется обработка журнала выявленных аномалий...")));
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "->*")));

            //richTextBox5.AppendText(Environment.NewLine + "Выполняется обработка журнала выявленных аномалий...");
            //richTextBox5.AppendText(Environment.NewLine + "->*");

            List<anomalyLogLine> anomalyLogLineS = new List<anomalyLogLine>();
            int pipeListCount = Convert.ToInt16(textBox110.Text);//получаем длину журнала из формы
            int incrementor = 0;//переменная для прогресс - индикатора
            int i = 4;
            bool mark = true;
            while (mark)
            {
                anomalyLogLine AnomalyLogLine = new anomalyLogLine();//создаём экземпляр класса строки журнала аномалий
                AnomalyLogLine.pipeNumber = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number20].Text);//расстояние от поперечного шва, м


                String txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number1].Text);
                try
                {
                    AnomalyLogLine.odometrDist = Convert.ToDouble(txt.Replace(".", ","));//длина
                    //MessageBox.Show("Число");
                }
                catch (Exception)
                {
                    AnomalyLogLine.odometrDist = 0;
                }


                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number2].Text);
                try
                {
                    AnomalyLogLine.thikness = Convert.ToDouble(txt.Trim().Replace(".", ","));//длина

                }
                catch (Exception)
                {
                    AnomalyLogLine.thikness = 0;
                }*/

                AnomalyLogLine.distanceFromTransverseWeld = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number3].Text);//расстояние от поперечного шва, м
                //AnomalyLogLine.distanceFromReferencePoints = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number4].Text);//расстояние от реперных точек
                AnomalyLogLine.featuresCharacter = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number5].Text);//характер особенности
                //AnomalyLogLine.classOfSize = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number6].Text);//класс размера
                AnomalyLogLine.featuresOrientation = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number7].Text);//ориентация


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
                    AnomalyLogLine.depthInProcent = Convert.ToDouble(txt.Trim().Replace(".", ","));//глубина дефекта в процентах                  
                }
                catch (Exception)
                {
                    AnomalyLogLine.depthInProcent = 0;
                }

                for (int g = 0; g < mGVTD.MGPipeS.Count; g++)
                {
                    if (String.Equals(AnomalyLogLine.pipeNumber, mGVTD.MGPipeS[g].pipeNumber))
                    {
                        AnomalyLogLine.thikness = mGVTD.MGPipeS[g].thikness;
                    }
                }

                /*txt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number11].Text);
                try
                {
                    AnomalyLogLine.depthInMm = Convert.ToDouble(txt.Replace(".", ","));//глубина дефекта в миллиметрах                    
                }
                catch (Exception)
                {
                    AnomalyLogLine.depthInMm = 0;
                }*/
                //AnomalyLogLine.depthInMm = Convert.ToDouble(txt.Replace(".", ","));//глубина дефекта в миллиметрах
                //AnomalyLogLine.extOrInt = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number12].Text);//характер локаизации(внутри или снаружи)
                //AnomalyLogLine.KBD = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number13].Text);//КБД


                AnomalyLogLine.defectAssessment = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number14].Text.Replace("(c)", "C").Replace("(a)", "A").Replace("(b)", "B"));//оценка дефекта
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
                //AnomalyLogLine.note = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number18].Text);//Примечание
                AnomalyLogLine.defectRepareDate = Convert.ToString(ObjWorkSheet2.Cells[i + 1, NumbersOfColumns.column2Number19].Text);//Примечание
                AnomalyLogLine.depthInMm = AnomalyLogLine.thikness * AnomalyLogLine.depthInProcent / 100;//в трубном журнале НПЦВТД нет глубины в мм, поэтому вычисляем.
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
                    richTextBox5.Invoke(new Action(() => richTextBox5.AppendText("*")));
                    incrementor = 0;
                }


            }
            //textBox110.Text = Convert.ToString(i);//записываем в поле количество труб
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "Массив данных из журнала выявленных аномалий прочитан, количество дефектов:" + anomalyLogLineS.Count)));
            richTextBox5.Invoke(new Action(() => richTextBox5.AppendText(Environment.NewLine + "==========================================")));
            //richTextBox5.AppendText(Environment.NewLine + "Массив данных из журнала выявленных аномалий прочитан, количество дефектов:" + anomalyLogLineS.Count);
            //richTextBox5.AppendText(Environment.NewLine + "==========================================");
            ObjExcel.Quit();
            return anomalyLogLineS;
        }
        ////
        //********************************************************************************        
        private MGVTD PipeLogWithCategory(MGVTD mGVTD)//расстановка категорий и характеристик труб
        {
            MGVTD MgvtdNew = new MGVTD();
            MgvtdNew = mGVTD;
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Выполняется расстановка характеристик труб в трубном журнале...(" + mGVTD.MGPipeS.Count + " труб)")));
            //richTextBox1.AppendText(Environment.NewLine + "Выполняется расстановка характеристик труб в трубном журнале...(" + mGVTD.MGPipeS.Count + " труб)");
            for (int i = 0; i < mGVTD.pipeCharacteristicsLog.Count; i++)
            {
                for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                {
                    if (String.IsNullOrWhiteSpace(mGVTD.MGPipeS[j].steelGrade))
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
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Расстановка характеристик труб выполнена")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "========================================")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Выполняется расстановка категорий участков в трубном журнале...(" + mGVTD.MGPipeS.Count + " труб)")));

            //richTextBox1.AppendText(Environment.NewLine + "Расстановка характеристик труб выполнена");
            //richTextBox1.AppendText(Environment.NewLine + "========================================");
            //richTextBox1.AppendText(Environment.NewLine + "Выполняется расстановка категорий участков в трубном журнале...(" + mGVTD.MGPipeS.Count + " труб)");
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
                if (mGVTD.anomalyLogLineS[i].thikness < 1)
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
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "Расстановка категорий участков выполнена")));
            richTextBox1.Invoke(new Action(() => richTextBox1.AppendText(Environment.NewLine + "========================================")));

            //richTextBox1.AppendText(Environment.NewLine + "Расстановка категорий участков выполнена");
            //richTextBox1.AppendText(Environment.NewLine + "========================================");
            return mGVTD;
        }
        private MGVTD PipeLogWithThikness(MGVTD mGVTD, bool mark)//расстановка категорий и характеристик труб
        {
            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)//если подрядчики не расставили толщину трубы в в журнале аномалий, расставим сами
            {
                if (mGVTD.anomalyLogLineS[i].thikness < 1)
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

            if (mark)
            {
                richTextBox3.AppendText(Environment.NewLine + "Расстановка толщин труб в журнале аномалий выполнена.");
                richTextBox3.AppendText(Environment.NewLine + "========================================");
            }
            else
            {
                richTextBox5.AppendText(Environment.NewLine + "Расстановка толщин труб в журнале аномалий выполнена.");
                richTextBox5.AppendText(Environment.NewLine + "========================================");
            }

            return mGVTD;
        }
        private void tableExcelReadToClass()//метод для чтения файла в класс
        {
            mGVTD.pipelineInfo = operatingReadToClassPipeInfo(fileName, NumbersOfColumns);//данные о трубе
            mGVTD.MGPipeS = OperatingReadToClassPipeLog(fileName, NumbersOfColumns);//трубный журнал
            mGVTD.anomalyLogLineS = OperatingReadToClassAnomalyLog(fileName, NumbersOfColumns);//журнал аномалий
            mGVTD.furnishingsLogS = OperatingReadToClassFurnishingsLog(fileName, NumbersOfColumns);//элементы обустройства
            mGVTD.pipeCharacteristicsLog = OperatingReadToClassPipeCharacteristics(fileName, NumbersOfColumns);//Характеристики труб
            mGVTD.pipelineSectionCategoryLogs = OperatingReadToClassPipelineSectionCategoryLog(fileName, NumbersOfColumns);//категории участков трубопровода
            mGVTD = PipeLogWithCategory(mGVTD);//расставим в трубном журнале характеристики труб и категории участков
        }
        private void shortTableExcelReadToClass()//КОРОТКИЙ!!! метод для чтения файла в класс
        {
            mGVTD.pipelineInfo = operatingReadToClassPipeInfo(fileName, NumbersOfColumns);//данные о трубе
            mGVTD.MGPipeS = ShortOperatingReadToClassPipeLogAutoFin(fileName, NumbersOfColumns);//трубный журнал
            mGVTD.anomalyLogLineS = shortOperatingReadToClassAnomalyLogAutoFin(fileName, NumbersOfColumns);//журнал аномалий
            mGVTD.furnishingsLogS = OperatingReadToClassFurnishingsLogAutoFin(fileName, NumbersOfColumns);//элементы обустройства
            mGVTD.pipeCharacteristicsLog = OperatingReadToClassPipeCharacteristics(fileName, NumbersOfColumns);//Характеристики труб
            mGVTD.pipelineSectionCategoryLogs = OperatingReadToClassPipelineSectionCategoryLog(fileName, NumbersOfColumns);//категории участков трубопровода

            if (isHaveCategory.Checked == false)
            {
                mGVTD = PipeLogWithCategory(mGVTD);//расставим в трубном журнале характеристики труб и категории участков
            }
        }
        //***************************************************
        private async void shortTableExcelReadToClassSOD()//ДЛЯ СОД!!!!!!!КОРОТКИЙ!!! метод для чтения файла в класс
        {
            mGVTD.pipelineInfo = operatingReadToClassPipeInfoSOD();//данные о трубе
            await Task.Run(() => mGVTD.MGPipeS = ShortOperatingReadToClassPipeLogAutoFinSOD(fileNamePipeLog, NumbersOfColumns));
            //mGVTD.MGPipeS = shortOperatingReadToClassPipeLogAutoFinSOD(fileNamePipeLog, NumbersOfColumns);//трубный журнал
            await Task.Run(() => mGVTD.anomalyLogLineS = shortOperatingReadToClassAnomalyLogAutoFinSOD(fileNameDefectLog, NumbersOfColumns));
            //mGVTD.anomalyLogLineS = shortOperatingReadToClassAnomalyLogAutoFinSOD(fileNameDefectLog, NumbersOfColumns);//журнал аномалий
            await Task.Run(() => mGVTD.furnishingsLogS = OperatingReadToClassFurnishingsLogAutoFinSOD(fileNameLineObjects, NumbersOfColumns));
            mGVTD.furnishingsLogS = OperatingReadToClassFurnishingsLogAutoFinSOD(fileNameLineObjects, NumbersOfColumns);//элементы обустройства
            mGVTD = PipeLogWithThikness(itIsTee(mGVTD), true);//помечаем соответствующие поля у секций, являющихся тройниками. Расставляем толщину стенки в журнале аномалий
            textBox131.Text = mGVTD.MGPipeS[0].pipeNumber;
            textBox136.Text = mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].pipeNumber;
            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)
            {
                for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                {
                    if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                    {
                        mGVTD.anomalyLogLineS[i].thikness = mGVTD.MGPipeS[j].thikness;
                    }
                }
            }
        }
        private async void shortTableExcelReadToClassNPCVTD()//ДЛЯ НПЦВТД!!!!!!!КОРОТКИЙ!!! метод для чтения файла в класс
        {
            mGVTD.pipelineInfo = operatingReadToClassPipeInfoNPCVTD();//данные о трубе
            await Task.Run(() => mGVTD.MGPipeS = ShortOperatingReadToClassPipeLogAutoFinNPCVTD(fileNamePipeLog, NumbersOfColumns));
            await Task.Run(() => mGVTD.anomalyLogLineS = shortOperatingReadToClassAnomalyLogAutoFinNPCVTD(fileNameDefectLog, NumbersOfColumns));

            try
            {
                await Task.Run(() => mGVTD.furnishingsLogS = OperatingReadToClassFurnishingsLogAutoFinNPCVTD(fileNameLineObjects, NumbersOfColumns));
                //mGVTD.furnishingsLogS = operatingReadToClassFurnishingsLogAutoFinNPCVTD(fileNameLineObjects, NumbersOfColumns);//элементы обустройства
            }
            catch (Exception)
            {
                await Task.Run(() => mGVTD.furnishingsLogS = FirnishingLogVirtual(mGVTD));
                //mGVTD.furnishingsLogS = firnishingLogVirtual(mGVTD);//Заполняем журнал элементов обустройства на основе данных трубного журнала
            }
            mGVTD = PipeLogWithThikness(itIsTee(mGVTD), false);//помечаем соответствующие поля у секций, являющихся тройниками. Расставляем толщину стенки в журнале аномалий
            textBox131.Text = mGVTD.MGPipeS[0].pipeNumber;
            textBox136.Text = mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].pipeNumber;
            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)
            {
                for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                {
                    if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.anomalyLogLineS[i].pipeNumber))
                    {
                        mGVTD.anomalyLogLineS[i].thikness = mGVTD.MGPipeS[j].thikness;
                    }
                }
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

                if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate))
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

                if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate))
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
                            if (localThikness > 0)
                            {
                                if (mGVTD.anomalyLogLineS[i].depthInProcent >= procentOfCorrosion)
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
                if (String.IsNullOrEmpty(mGVTD.anomalyLogLineS[i].defectRepareDate))
                {
                    if (mGVTD.anomalyLogLineS[i].depthInMm > 0 | mGVTD.anomalyLogLineS[i].depthInProcent > 0)
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

                    if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].defectRepareDate))
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
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate))
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
                    if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].defectRepareDate))//проверяем, что нет пометки об устранении дефекта
                    {
                        if (mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].depthInProcent >= procentOfCorrosion)//проверяем, что дефект глубже заданного уровня
                        {
                            if (mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].isLostMetal)//для всех труб с коррозией вычисляем ранг опасности и складываем, как того требует п. 6.1.2 СТО 292
                            {
                                double tensileStrength = 500;//ищем по трубному журналу предел прочности
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
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate))
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

                    if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].defectRepareDate))
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
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate))
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
                    if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate))
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

                    if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].featuresCharacter) == false)//ищем строку с каким-нибудь дефектом
                    {
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].defectRepareDate))//проверяем, что поле со сведениями об устранении дефекта пустое
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
                    for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo - 1; i++)
                    {
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].featuresCharacter) == false)//ищем строку с каким-нибудь дефектом
                        {
                            if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate))
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
                    if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate))
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
            double result = 0;
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
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].defectRepareDate))
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
                            if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate))
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
                                if (Rr > 1)//условие, что значение искомой величины по определению не больше единицы
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
                    if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate))
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
                        if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[PlotBoundaries.pipeIdNumberOne].defectRepareDate))
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
                            if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate))
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
            public double pipeOneKilometr;//километр первой трубы
            public double pipeTwoKilometr;//километр второй трубы
            public int pipeIdNumberOne;//порядковый номер первой трубы в таблице дефектов
            public int pipeIdNumberTwo;//порядковый номер второй трубы в таблице дефектов
            public string pipeNumberOnePipeLog;//имя первой трубы участка в трубном журнале
            public string pipeNumberTwoPipeLog;//имя второй трубы участка в трубном журнале
            public int pipeIdNumberOnePipeLog;//номер первой трубы участка в трубном журнале
            public int pipeIdNumberTwoPipeLog;//номер второй трубы участка в трубном журнале
        }

        private List<MGPipe> getAllValvePipes(MGVTD mGVTD)//составляем список всех труб, помеченных как краны
        {
            List<MGPipe> lineValvesPipes = new List<MGPipe>();
            lineValvesPipes.Clear();


            lineValvesPipes.Add(mGVTD.MGPipeS[0]);//добавляем в список первую трубу
            lineValvesPipes[0].note = String.Concat(lineValvesPipes[0].note, " - первая труба участка");


            //List<plotBoundaries> allPlots = new List<plotBoundaries>();//список границ всех межкрановых участков
            int valveID = 1;//порядковый номер крана
            for (int i = 0; i < mGVTD.furnishingsLogS.Count; i++)
            {
                if (mGVTD.furnishingsLogS[i].characterFeatures.Contains("ран") | mGVTD.furnishingsLogS[i].note.Contains("ран") | mGVTD.furnishingsLogS[i].characterFeatures.Contains("раниц"))
                {
                    for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                    {
                        if (String.Equals(mGVTD.MGPipeS[j].pipeNumber, mGVTD.furnishingsLogS[i].pipeNumber))
                        {
                            lineValvesPipes.Add(mGVTD.MGPipeS[j]);
                            string note2 = mGVTD.furnishingsLogS[i].characterFeatures;
                            string note3 = mGVTD.furnishingsLogS[i].note;
                            string note = "";
                            note = String.Concat(note2, "_", note3, ", номер трубы: ", mGVTD.furnishingsLogS[i].pipeNumber);
                            lineValvesPipes[valveID].note = "";
                            lineValvesPipes[valveID].note = note;
                            valveID++;
                            richTextBox4.AppendText(Environment.NewLine + valveID + ") Найден кран " + mGVTD.MGPipeS[j].note + ", номер трубы: " + mGVTD.MGPipeS[j].pipeNumber);
                        }
                    }
                }
            }
            lineValvesPipes.Add(mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1]);//добавляем в список последнюю трубу
            lineValvesPipes[lineValvesPipes.Count - 1].note = String.Concat(lineValvesPipes[lineValvesPipes.Count - 1].note, " - последняя труба участка");
            return lineValvesPipes;
        }
        private void setChekBoxNames(List<MGPipe> lineValvesPipes)
        {
            double a = 0.001;
            if (checkBox16.Checked)
            {
                a = -0.001;
            }

            try
            {
                checkBox1.Text = String.Concat(lineValvesPipes[0].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[0].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox1.Checked = true;
            }
            catch (Exception)
            {
                checkBox1.Enabled = false;
            }
            try
            {
                checkBox2.Text = String.Concat(lineValvesPipes[1].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[1].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox2.Checked = true;
            }
            catch (Exception)
            {
                checkBox2.Enabled = false;
            }
            try
            {
                checkBox3.Text = String.Concat(lineValvesPipes[2].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[2].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox3.Checked = true;
            }
            catch (Exception)
            {
                checkBox3.Enabled = false;
            }
            try
            {
                checkBox4.Text = String.Concat(lineValvesPipes[3].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[3].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox4.Checked = true;
            }
            catch (Exception)
            {
                checkBox4.Enabled = false;
            }
            try
            {
                checkBox5.Text = String.Concat(lineValvesPipes[4].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[4].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox5.Checked = true;
            }
            catch (Exception)
            {
                checkBox5.Enabled = false;
            }
            try
            {
                checkBox6.Text = String.Concat(lineValvesPipes[5].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[5].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox6.Checked = true;
            }
            catch (Exception)
            {
                checkBox6.Enabled = false;
            }
            try
            {
                checkBox7.Text = String.Concat(lineValvesPipes[6].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[6].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox7.Checked = true;
            }
            catch (Exception)
            {
                checkBox7.Enabled = false;
            }
            try
            {
                checkBox8.Text = String.Concat(lineValvesPipes[7].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[7].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox8.Checked = true;
            }
            catch (Exception)
            {
                checkBox8.Enabled = false;
            }
            try
            {
                checkBox9.Text = String.Concat(lineValvesPipes[8].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[8].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox9.Checked = true;
            }
            catch (Exception)
            {
                checkBox9.Enabled = false;
            }
            try
            {
                checkBox10.Text = String.Concat(lineValvesPipes[9].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[9].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox10.Checked = true;
            }
            catch (Exception)
            {
                checkBox10.Enabled = false;
            }
            try
            {
                checkBox11.Text = String.Concat(lineValvesPipes[10].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[10].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox11.Checked = true;
            }
            catch (Exception)
            {
                checkBox11.Enabled = false;
            }
            try
            {
                checkBox12.Text = String.Concat(lineValvesPipes[11].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[11].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox12.Checked = true;
            }
            catch (Exception)
            {
                checkBox12.Enabled = false;
            }
            try
            {
                checkBox13.Text = String.Concat(lineValvesPipes[12].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[12].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox13.Checked = true;
            }
            catch (Exception)
            {
                checkBox13.Enabled = false;
            }
            try
            {
                checkBox14.Text = String.Concat(lineValvesPipes[13].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[13].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox14.Checked = true;
            }
            catch (Exception)
            {
                checkBox14.Enabled = false;
            }
            try
            {
                checkBox15.Text = String.Concat(lineValvesPipes[14].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[14].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox15.Checked = true;
            }
            catch (Exception)
            {
                checkBox15.Enabled = false;
            }
            try
            {
                checkBox18.Text = String.Concat(lineValvesPipes[15].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[15].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox18.Checked = true;
            }
            catch (Exception)
            {
                checkBox18.Enabled = false;
            }
            try
            {
                checkBox19.Text = String.Concat(lineValvesPipes[16].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[16].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox19.Checked = true;
            }
            catch (Exception)
            {
                checkBox19.Enabled = false;
            }
            try
            {
                checkBox20.Text = String.Concat(lineValvesPipes[17].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[17].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox20.Checked = true;
            }
            catch (Exception)
            {
                checkBox20.Enabled = false;
            }
            try
            {
                checkBox21.Text = String.Concat(lineValvesPipes[18].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[18].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox21.Checked = true;
            }
            catch (Exception)
            {
                checkBox21.Enabled = false;
            }
            try
            {
                checkBox22.Text = String.Concat(lineValvesPipes[19].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[19].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox22.Checked = true;
            }
            catch (Exception)
            {
                checkBox22.Enabled = false;
            }
            try
            {
                checkBox23.Text = String.Concat(lineValvesPipes[20].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[20].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox23.Checked = true;
            }
            catch (Exception)
            {
                checkBox23.Enabled = false;
            }
            try
            {
                checkBox24.Text = String.Concat(lineValvesPipes[21].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[21].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox24.Checked = true;
            }
            catch (Exception)
            {
                checkBox24.Enabled = false;
            }
            try
            {
                checkBox25.Text = String.Concat(lineValvesPipes[22].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[22].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox25.Checked = true;
            }
            catch (Exception)
            {
                checkBox25.Enabled = false;
            }
            try
            {
                checkBox26.Text = String.Concat(lineValvesPipes[23].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[23].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox26.Checked = true;
            }
            catch (Exception)
            {
                checkBox26.Enabled = false;
            }
            try
            {
                checkBox27.Text = String.Concat(lineValvesPipes[24].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[24].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox27.Checked = true;
            }
            catch (Exception)
            {
                checkBox27.Enabled = false;
            }
            try
            {
                checkBox28.Text = String.Concat(lineValvesPipes[25].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[25].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox28.Checked = true;
            }
            catch (Exception)
            {
                checkBox28.Enabled = false;
            }
            try
            {
                checkBox29.Text = String.Concat(lineValvesPipes[26].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[26].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox29.Checked = true;
            }
            catch (Exception)
            {
                checkBox29.Enabled = false;
            }
            try
            {
                checkBox30.Text = String.Concat(lineValvesPipes[27].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[27].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox30.Checked = true;
            }
            catch (Exception)
            {
                checkBox30.Enabled = false;
            }
            try
            {
                checkBox31.Text = String.Concat(lineValvesPipes[28].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[28].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox31.Checked = true;
            }
            catch (Exception)
            {
                checkBox31.Enabled = false;
            }
            try
            {
                checkBox32.Text = String.Concat(lineValvesPipes[29].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[29].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox32.Checked = true;
            }
            catch (Exception)
            {
                checkBox32.Enabled = false;
            }
            try
            {
                checkBox33.Text = String.Concat(lineValvesPipes[30].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[30].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox33.Checked = true;
            }
            catch (Exception)
            {
                checkBox33.Enabled = false;
            }
            try
            {
                checkBox34.Text = String.Concat(lineValvesPipes[31].note, ", ", Convert.ToString(Convert.ToDouble(textBox381.Text.Replace(".", ",")) + a * (lineValvesPipes[31].odometrDist - lineValvesPipes[0].odometrDist)), " км.");
                checkBox34.Checked = true;
            }
            catch (Exception)
            {
                checkBox34.Enabled = false;
            }

        }
        private List<plotBoundaries> getAllPlotBoundaries(MGVTD mGVTD, List<MGPipe> allValves)//составляем список всех участков
        {
            List<plotBoundaries> allPlots = new List<plotBoundaries>();//будущий список участков для обсчета
            List<MGPipe> chekedValves = new List<MGPipe>();//список труб (кранов), на которых поставлены галочки
            {
                try
                {
                    if (checkBox1.Checked)
                    {
                        chekedValves.Add(allValves[0]);

                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox2.Checked)
                    {
                        chekedValves.Add(allValves[1]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox3.Checked)
                    {
                        chekedValves.Add(allValves[2]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox4.Checked)
                    {
                        chekedValves.Add(allValves[3]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox5.Checked)
                    {
                        chekedValves.Add(allValves[4]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox6.Checked)
                    {
                        chekedValves.Add(allValves[5]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox7.Checked)
                    {
                        chekedValves.Add(allValves[6]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox8.Checked)
                    {
                        chekedValves.Add(allValves[7]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox9.Checked)
                    {
                        chekedValves.Add(allValves[8]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox10.Checked)
                    {
                        chekedValves.Add(allValves[9]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox11.Checked)
                    {
                        chekedValves.Add(allValves[10]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox12.Checked)
                    {
                        chekedValves.Add(allValves[11]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox13.Checked)
                    {
                        chekedValves.Add(allValves[12]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox14.Checked)
                    {
                        chekedValves.Add(allValves[13]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox15.Checked)
                    {
                        chekedValves.Add(allValves[14]);
                    }
                }
                catch (Exception) { }

                //***********************************************************************************//
                try
                {
                    if (checkBox18.Checked)
                    {
                        chekedValves.Add(allValves[15]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox19.Checked)
                    {
                        chekedValves.Add(allValves[16]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox20.Checked)
                    {
                        chekedValves.Add(allValves[17]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox21.Checked)
                    {
                        chekedValves.Add(allValves[18]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox22.Checked)
                    {
                        chekedValves.Add(allValves[19]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox23.Checked)
                    {
                        chekedValves.Add(allValves[20]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox24.Checked)
                    {
                        chekedValves.Add(allValves[21]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox25.Checked)
                    {
                        chekedValves.Add(allValves[22]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox26.Checked)
                    {
                        chekedValves.Add(allValves[23]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox27.Checked)
                    {
                        chekedValves.Add(allValves[24]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox28.Checked)
                    {
                        chekedValves.Add(allValves[25]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox29.Checked)
                    {
                        chekedValves.Add(allValves[26]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox30.Checked)
                    {
                        chekedValves.Add(allValves[27]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox31.Checked)
                    {
                        chekedValves.Add(allValves[28]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox32.Checked)
                    {
                        chekedValves.Add(allValves[29]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox33.Checked)
                    {
                        chekedValves.Add(allValves[30]);
                    }
                }
                catch (Exception) { }
                try
                {
                    if (checkBox34.Checked)
                    {
                        chekedValves.Add(allValves[31]);
                    }
                }
                catch (Exception) { }
            }

            for (int i = 0; i < chekedValves.Count; i++)
            {
                richTextBox4.AppendText(Environment.NewLine + " Кран " + chekedValves[i].note + ", (номер трубы): " + chekedValves[i].pipeNumber + " добавлен к списку");
            }
            for (int i = 1; i < chekedValves.Count; i++)
            {
                plotBoundaries plot = lookingOfPlotBoundaries(mGVTD, chekedValves[i - 1].pipeNumber, chekedValves[i].pipeNumber);
                allPlots.Add(plot);
            }
            return allPlots;
        }

        private plotBoundaries lookingOfPlotBoundaries(MGVTD mGVTD, string pipeNumberOne, string pipeNumberTwo)//ищем имена и порядковые номера первой и последней труб в ведомости аномалий, попавших в заданный интервал трубного журнала. 
        {
            double apperKm = 0.001;
            if (checkBox16.Checked)
            {
                apperKm = -1 * apperKm;
            }
            plotBoundaries result = new plotBoundaries();
            result.pipeNumberOnePipeLog = pipeNumberOne;//имя первой трубы участка в трубном журнале
            result.pipeNumberTwoPipeLog = pipeNumberTwo;//имя второй трубы участка в трубном журнале
            int firstPipeID = 0;
            int secondPipeID = mGVTD.MGPipeS.Count;
            int marker = 0;//маркер для определения, что первая труба диапазона уже найдена
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)//ищем порядковые номера этих труб в трубном журнале
            {
                if (String.Equals(pipeNumberOne, mGVTD.MGPipeS[i].pipeNumber))
                {
                    firstPipeID = i;
                    result.pipeIdNumberOnePipeLog = firstPipeID;

                    result.pipeOneKilometr = Convert.ToDouble(textBox381.Text.Replace(".", ",")) + apperKm * (mGVTD.MGPipeS[i].odometrDist - mGVTD.MGPipeS[0].odometrDist);
                }
                if (String.Equals(pipeNumberTwo, mGVTD.MGPipeS[i].pipeNumber))
                {
                    secondPipeID = i;
                    result.pipeIdNumberTwoPipeLog = secondPipeID;

                    result.pipeTwoKilometr = Convert.ToDouble(textBox381.Text.Replace(".", ",")) + apperKm * (mGVTD.MGPipeS[i].odometrDist - mGVTD.MGPipeS[0].odometrDist);
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
        private async void button5_Click(object sender, EventArgs e)//чтение данных из таблицы ексель в экземпляр класса
        {
            //tableExcelReadToClass();//чтение данных из файла в экземпляр класса
            await Task.Run(() => shortTableExcelReadToClass());//чтение данных из файла в экземпляр класса (избирательный метод)
            mGVTD = itIsTee(mGVTD);//помечаем соответствующие поля у секций, являющихся тройниками
            textBox131.Text = mGVTD.MGPipeS[0].pipeNumber;
            textBox136.Text = mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].pipeNumber;
        }
        private void button6_Click(object sender, EventArgs e)//выполнение расчета
        {

        }

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
                    result = result + maxDentDamagOfHhisPipe;
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
        private int MaxCorrDefectNumber(MGVTD mGVTD, plotBoundaries PlotBoundaries)//ищем номер строки с максимальным дефектом потери металла
        {
            double maxLostMetalProcent = 0;
            int numberPipeWithMaxDefect = 0;
            for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo; i++)//ищем максимальную глубину дефекта
            {
                if (mGVTD.anomalyLogLineS[i].isLostMetal)
                {
                    if (String.IsNullOrEmpty(mGVTD.anomalyLogLineS[i].defectRepareDate))
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
            int result = 0;
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
            int result = 0;
            for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo; i++)
            {
                if (String.IsNullOrEmpty(mGVTD.anomalyLogLineS[i].defectRepareDate))
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
            int result = 0;
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
        private int numberOfDefectCoilUnderRoads(MGVTD mGVTD, plotBoundaries PlotBoundaries, bool printPipe)//считаем количество дефектных стыков в кожухах
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
                                                    for (int q = startPipeID + 1; q < finishPipeID + 1; q++)//+1 потому, что первый стык находится не в кожухе и мы его пропускаем
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
                            if (String.IsNullOrEmpty(mGVTD.anomalyLogLineS[j].defectRepareDate))//если дефект не помечен как устраненный
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
                                    if (printPipe)
                                    {
                                        richTextBox2.AppendText(Environment.NewLine + "Дефектный стык № " + mGVTD.anomalyLogLineS[j].pipeNumber + " внутри кожуха");
                                    }

                                }
                            }

                        }
                    }
                }
            }

            return result;
        }
        private List<string> namesOfDefectCoilUnderRoads(MGVTD mGVTD, plotBoundaries PlotBoundaries, bool printPipe)//считаем количество дефектных стыков в кожухах
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
                                                    for (int q = startPipeID + 1; q < finishPipeID + 1; q++)//+1 потому, что первый стык находится не в кожухе и мы его пропускаем
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
                            if (String.IsNullOrEmpty(mGVTD.anomalyLogLineS[j].defectRepareDate))//если дефект не помечен как устраненный
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
                                    if (printPipe)
                                    {
                                        richTextBox2.AppendText(Environment.NewLine + "Дефектный стык № " + mGVTD.anomalyLogLineS[j].pipeNumber + " внутри кожуха");
                                    }
                                }
                            }

                        }
                    }
                }
            }

            return defectpipes;
        }
        private int numberOfDefectLongitudinalWelds(MGVTD mGVTD, plotBoundaries PlotBoundaries)//считаем дефектные продольные швы
        {
            int result = 0;
            bool mark = true;
            List<string> defectpipes = new List<string>();//это просто список учтенных труб
            for (int i = PlotBoundaries.pipeIdNumberOne; i < PlotBoundaries.pipeIdNumberTwo; i++)
            {
                if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate))
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

        private void StartCalculations_Click(object sender, EventArgs e)
        {
            goEquation();
        }

        private void goEquation()
        {
            richTextBox2.AppendText(Environment.NewLine + "mGVTD.MGPipeS.Count=" + mGVTD.MGPipeS.Count);
            bool printPipe = true;//выводим на экран номера дефектных труб в кожухах
            allPipeCount = 0;//сумма всех труб участка++
            allPipeWhithСorrosion = 0;//сумма труб с коррозией++
            summCorrosionDamag = 0;//суммарная поврежденность от коррозии++
            allPipeWhithСorrosionPlus = 0;//сумма труб с коррозией++
            summCorrosionDamagPlus = 0;//суммарная поврежденность от коррозии++
            allPipeWhithDent = 0;//количество труб с вмятинами++
            summDentDamag = 0;//суммарная поврежденность от вмятин++        
            technicalConditionIndicatorOfPipesAndSDT = 0;//показатель технического состояния труб и СДТ++
            allPipeWhithJointDefects = 0;//количество труб с дефектами КСС++
            summJointDefectsDamag = 0;//суммарная поврежденность КСС

            double dCoil = 0;
            allDefectsWhithСorrosionPlus = 0;//сумма всех коррозионных дефектов глубиной больше указанного процента
            plotBoundaries PlotBoundaries = new plotBoundaries();
            mGVTD = isLostMetal(mGVTD);//расставляем метки на дефектах потери металла
            richTextBox2.AppendText(Environment.NewLine + "mGVTD.MGPipeS.Count=" + mGVTD.MGPipeS.Count);
            PlotBoundaries = lookingOfPlotBoundaries(mGVTD, textBox131.Text, textBox136.Text);
            allPipeCount = PlotBoundaries.pipeIdNumberTwoPipeLog - PlotBoundaries.pipeIdNumberOnePipeLog + 1;
            richTextBox2.AppendText(Environment.NewLine + "=======================================");
            richTextBox2.AppendText(Environment.NewLine + "Выполняется расчет Pвтд для участка газопровода в заданных границах");

            mGVTD = damagFromСorrosion(damageFromDent(damagOfCoilJoin(mGVTD, PlotBoundaries), PlotBoundaries), PlotBoundaries);
            richTextBox2.AppendText(Environment.NewLine + "mGVTD.MGPipeS.Count=" + mGVTD.MGPipeS.Count);
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
            PvtdReport = Pvtd;
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
            richTextBox2.AppendText(Environment.NewLine + "Доля труб с дефектами потери металла, %: " + Math.Round(100 * Convert.ToDouble(allPipeWhithСorrosion) / allPipeCount, 3));
            richTextBox2.AppendText(Environment.NewLine + "Максимальная глубина дефекта потери металла: " + mGVTD.anomalyLogLineS[MaxCorrDefectNumber(mGVTD, PlotBoundaries)].depthInProcent);
            double x9 = Math.Round(100 * Convert.ToDouble(allPipeWhithСorrosion) / allPipeCount, 3);//Доля труб с дефектами потери металла, %
            double x10 = mGVTD.anomalyLogLineS[MaxCorrDefectNumber(mGVTD, PlotBoundaries)].depthInProcent;//Максимальная глубина дефекта потери металла:
            summ = damagFromСorrosionAllDefects(mGVTD, PlotBoundaries, 15);//


            richTextBox2.AppendText(Environment.NewLine + "Плотность дефектов > 15%: " + 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3));//
            double x11 = 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3);//Плотность дефектов > 15%
            richTextBox2.AppendText(Environment.NewLine + "Доля труб с дефектами геометрии, %: " + Math.Round(100 * Convert.ToDouble(allPipeWhithDent) / allPipeCount, 3));
            double x12 = Math.Round(100 * Convert.ToDouble(allPipeWhithDent) / allPipeCount, 3);//Доля труб с дефектами геометрии, %
            richTextBox2.AppendText(Environment.NewLine + "Общее количество тройников: " + numberOfTriples(mGVTD, PlotBoundaries));
            double x13 = numberOfTriples(mGVTD, PlotBoundaries);//Общее количество тройников
            richTextBox2.AppendText(Environment.NewLine + "Количество дефектных тройников: " + numberOfDefectTriples(mGVTD, PlotBoundaries));
            double x14 = numberOfDefectTriples(mGVTD, PlotBoundaries);//Количество дефектных тройников
            richTextBox2.AppendText(Environment.NewLine + "======================================= ");
            richTextBox2.AppendText(Environment.NewLine + "Количество дефектных труб в кожухах: " + numberOfDefectCoilUnderRoads(mGVTD, PlotBoundaries, printPipe));
            double x15 = numberOfDefectCoilUnderRoads(mGVTD, PlotBoundaries, printPipe);//Количество дефектных труб в кожухах
            richTextBox2.AppendText(Environment.NewLine + "Количество аномальных поперечных швов: " + allPipeWhithJointDefects);
            double x16 = allPipeWhithJointDefects;//Количество аномальных поперечных швов
            richTextBox2.AppendText(Environment.NewLine + "Количество аномальных продольных швов: " + numberOfDefectLongitudinalWelds(mGVTD, PlotBoundaries));
            double x17 = numberOfDefectLongitudinalWelds(mGVTD, PlotBoundaries);//Количество аномальных продольных швов
            richTextBox2.AppendText(Environment.NewLine + x0 + ";" + x1 + ";" + x2 + ";" + x3 + ";" + x4 + ";" + x5 + ";" + x6 + ";" + x7 + ";" + x8 + ";" + x9 + ";" + x10 + ";" +
                x11 + ";" + x12 + ";" + x13 + ";" + x14 + ";" + x15 + ";" + x16 + ";" + x17);
            richTextBox2.AppendText(Environment.NewLine + "mGVTD.MGPipeS.Count=" + mGVTD.MGPipeS.Count);
        }//Проводим цикл вычислений по заданному интервалу

        private MGVTD VDTWhithoutRepareDefects(MGVTD mGVTD)
        {
            MGVTD result = mGVTD;

            result.anomalyLogLineS.Clear();
            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)
            {
                if (String.IsNullOrWhiteSpace(mGVTD.anomalyLogLineS[i].defectRepareDate) == false)
                {
                    result.anomalyLogLineS.Add(mGVTD.anomalyLogLineS[i]);
                }
            }

            return result;
        }

        private void goEquationSpecial(MGVTD mGVTD, plotBoundaries PlotBoundaries)
        {
            allPipeCount = 0;//сумма всех труб участка++
            allPipeWhithСorrosion = 0;//сумма труб с коррозией++
            summCorrosionDamag = 0;//суммарная поврежденность от коррозии++
            allPipeWhithСorrosionPlus = 0;//сумма труб с коррозией++
            summCorrosionDamagPlus = 0;//суммарная поврежденность от коррозии++
            allPipeWhithDent = 0;//количество труб с вмятинами++
            summDentDamag = 0;//суммарная поврежденность от вмятин++        
            technicalConditionIndicatorOfPipesAndSDT = 0;//показатель технического состояния труб и СДТ++
            allPipeWhithJointDefects = 0;//количество труб с дефектами КСС++
            summJointDefectsDamag = 0;//суммарная поврежденность КСС


            double dCoil = 0;
            allDefectsWhithСorrosionPlus = 0;//сумма всех коррозионных дефектов глубиной больше указанного процента
            //plotBoundaries PlotBoundaries = new plotBoundaries();
            mGVTD = isLostMetal(mGVTD);//расставляем метки на дефектах потери металла
            //PlotBoundaries = lookingOfPlotBoundaries(mGVTD, textBox131.Text, textBox136.Text);
            allPipeCount = PlotBoundaries.pipeIdNumberTwoPipeLog - PlotBoundaries.pipeIdNumberOnePipeLog + 1;
            //richTextBox2.AppendText(Environment.NewLine + "=======================================");
            //richTextBox2.AppendText(Environment.NewLine + "Выполняется расчет Pвтд для участка газопровода в заданных границах");

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
            //richTextBox2.AppendText(Environment.NewLine + "Выполнен расчет для участка МГ от трубы № " + textBox131.Text + " до трубы № " + textBox136.Text);

            //richTextBox2.AppendText(Environment.NewLine + "Повреждённость соединительных деталей линейного участка (ф. 5.9 СТО 292): " + Math.Round(dd, 3));
            //richTextBox2.AppendText(Environment.NewLine + "Повреждённость линейного участка МГ от вмятин и гофр (ф. 5.8 СТО 292): " + Math.Round(dr, 3));
            //richTextBox2.AppendText(Environment.NewLine + "Повреждённость линейного участка МГ от от дефектов КСС (ф. 5.10 СТО 292): " + Math.Round(dCoil, 3));
            //richTextBox2.AppendText(Environment.NewLine + "Pвтд= " + Math.Round(Pvtd, 3));
            richTextBox2.AppendText(Environment.NewLine + mGVTD.pipelineInfo.pipelineName + ";" + mGVTD.pipelineInfo.pipelineSection + ";" + mGVTD.pipelineInfo.pipeDiameter + ";" + mGVTD.pipelineInfo.examinationDate + ";" + mGVTD.pipelineInfo.designPressure + ";" + mGVTD.pipelineInfo.operatingPressure + ";" + mGVTD.pipelineInfo.comissioningYear + ";" + PlotBoundaries.pipeNumberOnePipeLog + ";" + PlotBoundaries.pipeNumberTwoPipeLog + ";" + allPipeCount + ";" + allPipeWhithСorrosion + ";" + Math.Round(summCorrosionDamag, 3) + ";" + Math.Round(dk, 3) + ";" + Math.Round(dc, 3) + ";" + Math.Round(Do, 3) + ";" +
            allPipeWhithDent + ";" + Math.Round(summDentDamag, 3) + ";" + Math.Round(dr, 3) + ";" + Math.Round(dd, 3) + ";" + Math.Round(Pt, 3) + ";" + allPipeWhithJointDefects + ";" +
            Math.Round(dJoin, 3) + ";" + Math.Round(dCoil, 3) + ";" + Math.Round(0.85 * dCoil, 3) + ";" + dSigma + ";" + df + ";" + Math.Round(Pvtd, 3));
            double procent = 15;
            //PlotBoundaries.pipeIdNumberTwoPipeLog - PlotBoundaries.pipeIdNumberOnePipeLog
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
            //PlotBoundaries = lookingOfPlotBoundaries(mGVTD, textBox131.Text, textBox136.Text);


            //richTextBox2.AppendText(Environment.NewLine + "=======================================");
            allPipeWhithСorrosionPlus = 0;
            summCorr2 = 0;
            int summ = damagFromСorrosionAllDefects(mGVTD, PlotBoundaries, procent);
            double damagg = damagFromСorrosionProcent(mGVTD, PlotBoundaries, procent);
            double x0 = procent;//процент коррозии из окна на вкладке "анализ"
            double x1 = Math.Round(damagg, 3);//поврежденность
            //richTextBox2.AppendText(Environment.NewLine + "Повреждённость локального участка от коррозии >" + procent + " % (ф. 5.3 СТО 292): " + Math.Round(damagg, 3));
            double x2 = summ;//количество
            //richTextBox2.AppendText(Environment.NewLine + "Количество коррозионных дефектов глубиной >" + procent + " % : " + summ);
            //richTextBox2.AppendText(Environment.NewLine + "=======================================");

            allPipeWhithСorrosionPlus = 0;
            summCorr2 = 0;
            summ = damagFromСorrosionAllDefects(mGVTD, PlotBoundaries, 30);
            damagg = damagFromСorrosionProcent(mGVTD, PlotBoundaries, 30);
            double lengthMG = mGVTD.MGPipeS[PlotBoundaries.pipeIdNumberTwoPipeLog].odometrDist - mGVTD.MGPipeS[PlotBoundaries.pipeIdNumberOnePipeLog].odometrDist;//протяженность участка
            double x3 = Math.Round(damagg, 3);//поврежденность от коррозии 30%
            double x4 = summ;//количество дефектов >30%
            //richTextBox2.AppendText(Environment.NewLine + "Повреждённость локального участка от коррозии >" + 30 + " % (ф. 5.3 СТО 292): " + Math.Round(damagg, 3));
            //richTextBox2.AppendText(Environment.NewLine + "Количество коррозионных дефектов глубиной >" + 30 + " % : " + summ);
            //richTextBox2.AppendText(Environment.NewLine + "Плотность дефектов > 30%: " + 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3));//
            //richTextBox2.AppendText(Environment.NewLine + "=======================================");
            double x5 = 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3);//Плотность дефектов > 30%
            summCorr2 = 0;
            summ = damagFromСorrosionAllDefects(mGVTD, PlotBoundaries, 0);
            damagg = damagFromСorrosionProcent(mGVTD, PlotBoundaries, 0);
            double x6 = Math.Round(damagg, 3);//поврежденность от коррозии
            double x7 = summ;//количество дефектов
            //richTextBox2.AppendText(Environment.NewLine + "Повреждённость локального участка от коррозии (все корр. деф.)  (ф. 5.3 СТО 292): " + Math.Round(damagg, 3));
            //richTextBox2.AppendText(Environment.NewLine + "Количество коррозионных дефектов : " + summ);
            //richTextBox2.AppendText(Environment.NewLine + "Плотность коррозионных дефектов: " + 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3));//
            double x8 = 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3);//Плотность коррозионных дефектов
            //richTextBox2.AppendText(Environment.NewLine + "=======================================");
            //richTextBox2.AppendText(Environment.NewLine + "Доля труб с дефектами потери металла, %: " + Math.Round(100 * Convert.ToDouble(allPipeWhithСorrosion) / allPipeCount, 3));
            //richTextBox2.AppendText(Environment.NewLine + "Максимальная глубина дефекта потери металла: " + mGVTD.anomalyLogLineS[MaxCorrDefectNumber(mGVTD, PlotBoundaries)].depthInProcent);
            double x9 = Math.Round(100 * Convert.ToDouble(allPipeWhithСorrosion) / allPipeCount, 3);//Доля труб с дефектами потери металла, %
            double x10 = mGVTD.anomalyLogLineS[MaxCorrDefectNumber(mGVTD, PlotBoundaries)].depthInProcent;//Максимальная глубина дефекта потери металла:
            summ = damagFromСorrosionAllDefects(mGVTD, PlotBoundaries, 15);//


            //richTextBox2.AppendText(Environment.NewLine + "Плотность дефектов > 15%: " + 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3));//
            double x11 = 1000 * Math.Round(Convert.ToDouble(summ) / lengthMG, 3);//Плотность дефектов > 15%
            //richTextBox2.AppendText(Environment.NewLine + "Доля труб с дефектами геометрии, %: " + Math.Round(100 * Convert.ToDouble(allPipeWhithDent) / allPipeCount, 3));
            double x12 = Math.Round(100 * Convert.ToDouble(allPipeWhithDent) / allPipeCount, 3);//Доля труб с дефектами геометрии, %
            //richTextBox2.AppendText(Environment.NewLine + "Общее количество тройников: " + numberOfTriples(mGVTD, PlotBoundaries));
            double x13 = numberOfTriples(mGVTD, PlotBoundaries);//Общее количество тройников
            //richTextBox2.AppendText(Environment.NewLine + "Количество дефектных тройников: " + numberOfDefectTriples(mGVTD, PlotBoundaries));
            double x14 = numberOfDefectTriples(mGVTD, PlotBoundaries);//Количество дефектных тройников
            //richTextBox2.AppendText(Environment.NewLine + "======================================= ");
            //richTextBox2.AppendText(Environment.NewLine + "Количество дефектных труб в кожухах: " + numberOfDefectCoilUnderRoads(mGVTD, PlotBoundaries));
            bool printPipe = false;
            double x15 = numberOfDefectCoilUnderRoads(mGVTD, PlotBoundaries, printPipe);//Количество дефектных труб в кожухах
            List<string> defectCoils = namesOfDefectCoilUnderRoads(mGVTD, PlotBoundaries, printPipe);
            string x18 = "";
            if (defectCoils.Count > 0)
            {
                for (int i = 0; i < defectCoils.Count; i++)
                {
                    x18 = String.Concat(x18, defectCoils[i], ", ");
                }
            }

            //richTextBox2.AppendText(Environment.NewLine + "Количество аномальных поперечных швов: " + allPipeWhithJointDefects);
            double x16 = allPipeWhithJointDefects;//Количество аномальных поперечных швов
            //richTextBox2.AppendText(Environment.NewLine + "Количество аномальных продольных швов: " + numberOfDefectLongitudinalWelds(mGVTD, PlotBoundaries));
            double x17 = numberOfDefectLongitudinalWelds(mGVTD, PlotBoundaries);//Количество аномальных продольных швов
            richTextBox2.AppendText(";" + x0 + ";" + x1 + ";" + x2 + ";" + x3 + ";" + x4 + ";" + x5 + ";" + x6 + ";" + x7 + ";" + x8 + ";" + x9 + ";" + x10 + ";" +
                x11 + ";" + x12 + ";" + x13 + ";" + x14 + ";" + x15 + ";(" + x18 + ");" + x16 + ";" + x17 + ";" + mGVTD.MGPipeS[PlotBoundaries.pipeIdNumberOnePipeLog].note + ";" + mGVTD.MGPipeS[PlotBoundaries.pipeIdNumberTwoPipeLog].note + ";" + PlotBoundaries.pipeOneKilometr + ";" + PlotBoundaries.pipeTwoKilometr);
        }
        string fileNamePipeLog;
        string fileNameDefectLog;
        string fileNameLineObjects;
        private void ButtonOpenPipeLog_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNamePipeLog = openFileDialog1.FileName;

                ButtonOpenPipeLog.BackColor = Color.Azure;
                //findStart();//поиск начала и конца журналов свойств труб и категорий
                //tableArdesTest();
            }
        }

        private void ButtonOpenDefectLog_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNameDefectLog = openFileDialog1.FileName;
                ButtonOpenDefectLog.BackColor = Color.Azure;
                //findStart();//поиск начала и конца журналов свойств труб и категорий
                //tableArdesTest();//
            }
        }

        private void ButtonOpenLineObjects_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNameLineObjects = openFileDialog1.FileName;
                ButtonOpenLineObjects.BackColor = Color.Azure;
                //findStart();//поиск начала и конца журналов свойств труб и категорий
                //tableArdesTest();
            }
        }

        private void CheckIndexes_Click(object sender, EventArgs e)
        {
            tableArdesTestSODPipeLog();
            tableArdesTestSODDefectLog();
            tableArdesTestLineObjects();
        }

        private void ReadReportToMemory_Click(object sender, EventArgs e)
        {
            shortTableExcelReadToClassSOD();
        }

        private void ReloadDiameter_Click(object sender, EventArgs e)
        {
            mGVTD.pipelineInfo = operatingReadToClassPipeInfoSOD();//данные о трубе
        }

        private void SOD2020_Click(object sender, EventArgs e)
        {
            parsSOD2020();
        }
        private void parsSOD2020()//один из вариантов расположения столбцов журнала дефектов
        {
            textBox195.Text = "7";
            textBox193.Text = "21";
            textBox192.Text = "8";
            textBox191.Text = "7";
            textBox190.Text = "5";
            textBox189.Text = "6";
            textBox188.Text = "14";
            textBox187.Text = "10";
            textBox186.Text = "11";
            textBox185.Text = "12";
            textBox184.Text = "13";
            textBox183.Text = "17";
            textBox182.Text = "18";
            textBox216.Text = "29";
            textBox176.Text = "2";
            tableArdesTestSODDefectLog();
        }
        private void pars2SOD2020()//один из вариантов расположения столбцов журнала дефектов
        {
            textBox195.Text = "7";
            textBox193.Text = "8";
            textBox192.Text = "7";
            textBox191.Text = "4";
            textBox190.Text = "5";
            textBox189.Text = "11";
            textBox188.Text = "8";
            textBox187.Text = "9";
            textBox186.Text = "10";
            textBox185.Text = "23";
            textBox184.Text = "3";
            textBox183.Text = "16";
            textBox182.Text = "15";
            textBox216.Text = "30";
            textBox176.Text = "2";
            tableArdesTestSODDefectLog();
        }
        private void pars3SOD2020()//один из вариантов расположения столбцов журнала дефектов
        {
            textBox195.Text = "7";
            textBox193.Text = "8";
            textBox192.Text = "7";
            textBox191.Text = "5";
            textBox190.Text = "6";
            textBox189.Text = "14";
            textBox188.Text = "10";
            textBox187.Text = "11";
            textBox186.Text = "12";
            textBox185.Text = "29";
            textBox184.Text = "4";
            textBox183.Text = "19";
            textBox182.Text = "22";
            textBox216.Text = "30";
            textBox176.Text = "2";
            tableArdesTestSODDefectLog();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            pars2SOD2020();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pars3SOD2020();
        }
        private void lookingForNumbersOfColumns()
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileNameDefectLog, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet3;

            //*********************Обработка журнала аномалий
            string WorksheetName3 = textBox196.Text;//получаем название вкладки из формы импотра
            try
            {
                ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];
            }
            catch (Exception)
            {
                WorksheetName3 = WorksheetName3.Replace(".xlsx", "");
                ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];
            }

            bool isFindFinish = false;
            for (int i = 1; i < 50; i++)
            {
                string columnName = Convert.ToString(ObjWorkSheet3.Cells[1, i].Text);
                if (columnName.Contains(Convert.ToString(textBox148.Text)))
                {
                    textBox195.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox149.Text)))
                {
                    textBox193.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox150.Text)))
                {
                    textBox192.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox153.Text)))
                {
                    textBox191.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox158.Text)))
                {
                    textBox190.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox159.Text)))
                {
                    textBox189.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox160.Text)))
                {
                    textBox188.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox161.Text)))
                {
                    textBox187.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox162.Text)))
                {
                    textBox186.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox165.Text)))
                {
                    textBox185.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox178.Text)))
                {
                    textBox184.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox179.Text)))
                {
                    textBox183.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox180.Text)))
                {
                    textBox182.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox181.Text)))
                {
                    textBox176.Text = Convert.ToString(i);
                }
                else if (String.IsNullOrEmpty(columnName))
                {
                    if (isFindFinish == false)
                    {
                        textBox216.Text = Convert.ToString(i);
                        isFindFinish = true;
                    }
                }
            }
            ObjExcel.Quit();
            //*********************Обработка трубного журнала
            ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            ObjWorkBook = ObjExcel.Workbooks.Open(fileNamePipeLog, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            WorksheetName3 = textBox170.Text;//получаем название вкладки из формы импотра
            try
            {
                ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];
            }
            catch (Exception)
            {
                WorksheetName3 = WorksheetName3.Replace(".xlsx", "");
                ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];
            }
            isFindFinish = false;
            for (int i = 1; i < 50; i++)
            {
                string columnName = Convert.ToString(ObjWorkSheet3.Cells[1, i].Text);
                if (columnName.Contains(Convert.ToString(textBox198.Text)))
                {
                    textBox169.Text = Convert.ToString(i);
                }

                else if (columnName.Contains(Convert.ToString(textBox211.Text)))
                {
                    textBox168.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox212.Text)))
                {
                    textBox167.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox213.Text)))
                {
                    textBox166.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox214.Text)))
                {
                    textBox164.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox217.Text)))
                {
                    textBox163.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox218.Text)))
                {
                    textBox144.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox219.Text)))
                {
                    textBox172.Text = Convert.ToString(i);
                }
                else if (String.IsNullOrEmpty(columnName))
                {
                    if (isFindFinish == false)
                    {
                        textBox175.Text = Convert.ToString(i);
                        isFindFinish = true;
                    }
                }
            }
            ObjExcel.Quit();
            //*********************Обработка журнала линейных объектов
            ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            ObjWorkBook = ObjExcel.Workbooks.Open(fileNameLineObjects, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            WorksheetName3 = textBox231.Text;//получаем название вкладки из формы импотра

            try
            {
                ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];
            }
            catch (Exception)
            {
                WorksheetName3 = WorksheetName3.Replace(".xlsx", "");
                ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];
            }
            isFindFinish = false;
            for (int i = 1; i < 50; i++)
            {
                string columnName = Convert.ToString(ObjWorkSheet3.Cells[1, i].Text);
                if (columnName.Contains(Convert.ToString(textBox220.Text)))
                {
                    textBox229.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox221.Text)))
                {
                    textBox228.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox222.Text)))
                {
                    textBox227.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox223.Text)))
                {
                    textBox226.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox225.Text)))
                {
                    textBox224.Text = Convert.ToString(i);
                }
            }
            ObjExcel.Quit();
        }
        private void lookingForNumbersOfColumnsNPCVTD()
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileNameDefectLog, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet3;

            //*********************Обработка журнала аномалий
            string WorksheetName3 = textBox274.Text;//получаем название вкладки из формы импотра
            ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];
            bool isFindFinish = false;
            for (int i = 1; i < 50; i++)
            {
                string columnName = Convert.ToString(ObjWorkSheet3.Cells[4, i].Text);
                if (columnName.Contains(Convert.ToString(textBox300.Text)))
                {
                    textBox271.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox270.Text)))
                {
                    textBox288.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox301.Text)))
                {
                    textBox289.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox302.Text)))
                {
                    textBox290.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox303.Text)))
                {
                    textBox291.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox304.Text)))
                {
                    textBox292.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox305.Text)))
                {
                    textBox293.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox306.Text)))
                {
                    textBox294.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox307.Text)))
                {
                    textBox295.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox310.Text)))
                {
                    textBox298.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox311.Text)))
                {
                    textBox299.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox312.Text)))
                {
                    textBox315.Text = Convert.ToString(i);
                }


                else if (String.IsNullOrEmpty(columnName))
                {
                    if (isFindFinish == false)
                    {
                        textBox272.Text = Convert.ToString(i);
                        isFindFinish = true;
                    }
                }
            }
            ObjExcel.Quit();
            //*********************Обработка трубного журнала
            ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            ObjWorkBook = ObjExcel.Workbooks.Open(fileNamePipeLog, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            WorksheetName3 = textBox250.Text;//получаем название вкладки из формы импотра
            ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];
            isFindFinish = false;
            for (int i = 1; i < 50; i++)
            {
                string columnName = Convert.ToString(ObjWorkSheet3.Cells[4, i].Text);
                if (columnName.Contains(Convert.ToString(textBox262.Text)))
                {
                    textBox249.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox263.Text)))
                {
                    textBox248.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox264.Text)))
                {
                    textBox247.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox267.Text)))
                {
                    textBox254.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox265.Text)))
                {
                    textBox246.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox266.Text)))
                {
                    textBox252.Text = Convert.ToString(i);
                }

                else if (columnName.Contains(Convert.ToString(textBox268.Text)))
                {
                    textBox255.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox269.Text)))
                {
                    textBox258.Text = Convert.ToString(i);
                }
                else if (columnName.Contains(Convert.ToString(textBox333.Text)))
                {
                    textBox260.Text = Convert.ToString(i);
                }

            }
            ObjExcel.Quit();
            //*********************Обработка журнала линейных объектов

            if (String.IsNullOrEmpty(fileNameLineObjects) == false)
            {
                ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.                                                                                                                                                        
                ObjWorkBook = ObjExcel.Workbooks.Open(fileNameLineObjects, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                WorksheetName3 = textBox316.Text;//получаем название вкладки из формы импотра
                ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];
                isFindFinish = false;
                for (int i = 1; i < 50; i++)
                {
                    string columnName = Convert.ToString(ObjWorkSheet3.Cells[4, i].Text);
                    if (columnName.Contains(Convert.ToString(textBox328.Text)))
                    {
                        textBox321.Text = Convert.ToString(i);
                    }
                    else if (columnName.Contains(Convert.ToString(textBox329.Text)))
                    {
                        textBox322.Text = Convert.ToString(i);
                    }
                    else if (columnName.Contains(Convert.ToString(textBox330.Text)))
                    {
                        textBox323.Text = Convert.ToString(i);
                    }
                    else if (columnName.Contains(Convert.ToString(textBox331.Text)))
                    {
                        textBox324.Text = Convert.ToString(i);
                    }
                    else if (columnName.Contains(Convert.ToString(textBox332.Text)))
                    {
                        textBox326.Text = Convert.ToString(i);
                    }
                }
                ObjExcel.Quit();
            }


        }

        private void autoLookColumns_Click(object sender, EventArgs e)
        {
            lookingForNumbersOfColumns();
            tableArdesTestSODPipeLog();
            tableArdesTestSODDefectLog();
            tableArdesTestLineObjects();
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            textBox14.Text = "8";
            textBox15.Text = "9";
            textBox16.Text = "10";
            textBox17.Text = "11";
            tableArdesTest();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            textBox14.Text = "9";
            textBox15.Text = "10";
            textBox16.Text = "11";
            textBox17.Text = "12";
            tableArdesTest();
        }

        List<MGPipe> allValves = new List<MGPipe>();
        private void lookingForValves_Click(object sender, EventArgs e)//кнопка для поиска трубных секций, помеченных как краны
        {
            allValves = getAllValvePipes(mGVTD);
            setChekBoxNames(allValves);
        }
        private void makeBoundariesList_Click(object sender, EventArgs e)//кнопка для формирования списка участков для обсчета
        {
            List<plotBoundaries> allPlots = getAllPlotBoundaries(mGVTD, allValves);
            richTextBox2.Clear();

            for (int i = 0; i < allPlots.Count; i++)
            {
                goEquationSpecial(mGVTD, allPlots[i]);
            }
            string text = string.Empty;

            for (int i = 1; i < richTextBox2.Lines.Length; i++)
                text += richTextBox2.Lines[i] + Environment.NewLine;

            Clipboard.SetText(text);
            //Clipboard.SetText(richTextBox2.Text);
            richTextBox4.AppendText(Environment.NewLine + "Вычисления выполнены. Результаты скопированы в буфер обмена.");
        }
        private void exportTo1C_automatic(MGVTD mGVTD, plotBoundaries PlotBoundaries)//Создание Excel файла для экспорта в Сонар
        {
            object Nothing = System.Reflection.Missing.Value;
            var app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Add(Nothing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets[1];

            worksheet.Name = "SonarFormat";

            // Write data
            worksheet.Cells[1, 22] = mGVTD.pipelineInfo.pipelineName;//public string pipelineName;//трубопровод (название)
            worksheet.Cells[1, 23] = mGVTD.pipelineInfo.pipelineSection;//public string pipelineSection;//участок трубопровода
            worksheet.Cells[1, 24] = mGVTD.pipelineInfo.pipeDiameter;//public double pipeDiameter;//диаметр трубы            
            worksheet.Cells[1, 25] = mGVTD.pipelineInfo.examinationDate;//public string examinationDate;//дата обследования
            worksheet.Cells[1, 26] = mGVTD.pipelineInfo.designPressure;//public double designPressure;// проектное давление
            worksheet.Cells[1, 27] = mGVTD.pipelineInfo.operatingPressure;//public double operatingPressure;// рабочее давление
            worksheet.Cells[1, 28] = mGVTD.pipelineInfo.comissioningYear;//public string comissioningYear;//год ввода в экспуатацию
            worksheet.Cells[1, 29] = PlotBoundaries.pipeNumberOnePipeLog;//Первая граница участка
            worksheet.Cells[1, 30] = PlotBoundaries.pipeNumberTwoPipeLog;//вторая граница участка
            worksheet.Cells[1, 31] = mGVTD.pipelineInfo.contractor;//организация - подрядчик
            worksheet.Cells[1, 32] = PlotBoundaries.pipeOneKilometr;//организация - подрядчик
            worksheet.Cells[1, 33] = PlotBoundaries.pipeTwoKilometr;//организация - подрядчик



            worksheet.Cells[1, 1] = "N_OSOB";
            worksheet.Cells[1, 2] = "N_SEK";
            worksheet.Cells[1, 3] = "L_ODOM";
            worksheet.Cells[1, 4] = "OTN_D";
            worksheet.Cells[1, 5] = "OSOB";
            worksheet.Cells[1, 6] = "L_SEK";
            worksheet.Cells[1, 7] = "T_ST";
            worksheet.Cells[1, 8] = "H_PROC";
            worksheet.Cells[1, 9] = "L_DEF";
            worksheet.Cells[1, 10] = "W_DEF";
            worksheet.Cells[1, 11] = "GRAD";
            worksheet.Cells[1, 12] = "TYPE";
            worksheet.Cells[1, 13] = "ABC";
            worksheet.Cells[1, 14] = "PRIM";
            worksheet.Cells[1, 15] = "DistanceFromReferencePoints";//расстояние от реперных точек
            worksheet.Cells[1, 16] = "Category";
            worksheet.Cells[1, 17] = "Grade";
            worksheet.Cells[1, 18] = "yieldPoint";
            worksheet.Cells[1, 19] = "tensileStrength";
            worksheet.Cells[1, 20] = "Repere_date";//дата устранения дефекта
            worksheet.Cells[1, 21] = "Глубина дефекта в мм";//дата устранения дефекта
            int strNunber = 2;//номер строки в формируемой таблице
            for (int i = PlotBoundaries.pipeIdNumberOnePipeLog; i < PlotBoundaries.pipeIdNumberTwoPipeLog + 1; i++)
            {
                string pipeNumber = mGVTD.MGPipeS[i].pipeNumber;//номер текущей трубы
                worksheet.Cells[strNunber, 1] = "";//N_OSOB
                worksheet.Cells[strNunber, 2] = mGVTD.MGPipeS[i].pipeNumber;//N_SEK
                worksheet.Cells[strNunber, 3] = mGVTD.MGPipeS[i].odometrDist;//L_ODOM
                worksheet.Cells[strNunber, 4] = "";//OTN_D
                worksheet.Cells[strNunber, 5] = mGVTD.MGPipeS[i].characterFeatures;//OSOB
                worksheet.Cells[strNunber, 6] = mGVTD.MGPipeS[i].pipeLength;//L_SEK
                worksheet.Cells[strNunber, 7] = mGVTD.MGPipeS[i].thikness;//T_ST
                worksheet.Cells[strNunber, 8] = 0;//H_PROC
                worksheet.Cells[strNunber, 9] = "";//L_DEF
                worksheet.Cells[strNunber, 10] = "";//W_DEF
                worksheet.Cells[strNunber, 11] = mGVTD.MGPipeS[i].jointAngle;//GRAD
                worksheet.Cells[strNunber, 12] = "";//TYPE
                worksheet.Cells[strNunber, 13] = "";//ABC
                worksheet.Cells[strNunber, 14] = mGVTD.MGPipeS[i].note;//PRIM
                worksheet.Cells[strNunber, 15] = mGVTD.MGPipeS[i].distanceFromReferencePoints;//расстояние от реперных точек
                worksheet.Cells[strNunber, 16] = mGVTD.MGPipeS[i].pipelineSectionCategory;//категория участка
                worksheet.Cells[strNunber, 17] = mGVTD.MGPipeS[i].steelGrade;//марка стали
                worksheet.Cells[strNunber, 18] = mGVTD.MGPipeS[i].yieldPoint;//предел текучести
                worksheet.Cells[strNunber, 19] = mGVTD.MGPipeS[i].tensileStrength;//предел прочности
                worksheet.Cells[strNunber, 25] = mGVTD.MGPipeS[i].Latitude;//Широта
                worksheet.Cells[strNunber, 26] = mGVTD.MGPipeS[i].Longitude;//Долгота
                strNunber++;
                for (int j = 0; j < mGVTD.anomalyLogLineS.Count; j++)
                {
                    if (String.Equals(pipeNumber, mGVTD.anomalyLogLineS[j].pipeNumber))
                    {
                        if (String.Equals(mGVTD.anomalyLogLineS[j].featuresCharacter, mGVTD.MGPipeS[i].characterFeatures) == false)
                        {
                            worksheet.Cells[strNunber, 1] = j;//N_OSOB
                            worksheet.Cells[strNunber, 2] = mGVTD.anomalyLogLineS[j].pipeNumber;//N_SEK
                            worksheet.Cells[strNunber, 3] = mGVTD.anomalyLogLineS[j].odometrDist;//L_ODOM
                            worksheet.Cells[strNunber, 4] = mGVTD.anomalyLogLineS[j].distanceFromTransverseWeld;//OTN_D
                            worksheet.Cells[strNunber, 5] = mGVTD.anomalyLogLineS[j].featuresCharacter;//OSOB
                            worksheet.Cells[strNunber, 6] = mGVTD.MGPipeS[i].pipeLength;//L_SEK
                            worksheet.Cells[strNunber, 7] = mGVTD.MGPipeS[i].thikness;//T_ST
                            worksheet.Cells[strNunber, 8] = mGVTD.anomalyLogLineS[j].depthInProcent;//H_PROC
                            worksheet.Cells[strNunber, 9] = mGVTD.anomalyLogLineS[j].length;//L_DEF
                            worksheet.Cells[strNunber, 10] = mGVTD.anomalyLogLineS[j].widht;//W_DEF
                            worksheet.Cells[strNunber, 11] = mGVTD.anomalyLogLineS[j].featuresOrientation;//GRAD
                            worksheet.Cells[strNunber, 12] = mGVTD.anomalyLogLineS[j].extOrInt;//TYPE
                            worksheet.Cells[strNunber, 13] = mGVTD.anomalyLogLineS[j].defectAssessment;//ABC
                            worksheet.Cells[strNunber, 14] = mGVTD.anomalyLogLineS[j].note;//PRIM
                            worksheet.Cells[strNunber, 15] = mGVTD.anomalyLogLineS[j].distanceFromReferencePoints;//distanceFromReferencePoints
                            worksheet.Cells[strNunber, 20] = mGVTD.anomalyLogLineS[j].defectRepareDate;//дата устранения дефекта
                            worksheet.Cells[strNunber, 25] = mGVTD.anomalyLogLineS[j].Latitude;//Широта
                            worksheet.Cells[strNunber, 26] = mGVTD.anomalyLogLineS[j].Longitude;//Долгота
                            if (String.IsNullOrEmpty(Convert.ToString(mGVTD.anomalyLogLineS[j].depthInMm)) == false)
                            {
                                if (mGVTD.anomalyLogLineS[j].depthInMm > 0)
                                {
                                    worksheet.Cells[strNunber, 21] = mGVTD.anomalyLogLineS[j].depthInMm;
                                }
                                else
                                {
                                    if (mGVTD.anomalyLogLineS[j].featuresCharacter.Contains("мятин"))
                                    {
                                        worksheet.Cells[strNunber, 21] = 0.01 * mGVTD.anomalyLogLineS[j].depthInProcent * mGVTD.pipelineInfo.pipeDiameter;
                                    }
                                    else
                                    {
                                        worksheet.Cells[strNunber, 21] = 0.01 * mGVTD.anomalyLogLineS[j].depthInProcent * mGVTD.MGPipeS[i].thikness;
                                    }

                                }
                            }
                            else
                            {
                                if (mGVTD.anomalyLogLineS[j].featuresCharacter.Contains("мятин"))
                                {
                                    worksheet.Cells[strNunber, 21] = 0.01 * mGVTD.anomalyLogLineS[j].depthInProcent * mGVTD.pipelineInfo.pipeDiameter;
                                }
                                else
                                {
                                    worksheet.Cells[strNunber, 21] = 0.01 * mGVTD.anomalyLogLineS[j].depthInProcent * mGVTD.MGPipeS[i].thikness;
                                }
                            }

                            strNunber++;
                        }
                    }
                }

            }

            /*string filename = textBox_adres.Text + mGVTD.pipelineInfo.pipelineName + "-" + PlotBoundaries.pipeNumberOnePipeLog+"-"+ PlotBoundaries.pipeNumberOnePipeLog+"xlsx";
            worksheet.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(false, Type.Missing, Type.Missing);
            app.Quit();*/

            // Show save file dialog
            //SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            /*if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {*/

            String dir = textBox_adres.Text;
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            saveFileDialog1.FileName = textBox_adres.Text + mGVTD.pipelineInfo.pipelineName + "-" + PlotBoundaries.pipeNumberOnePipeLog + "-" + PlotBoundaries.pipeNumberTwoPipeLog + "-(" + PlotBoundaries.pipeOneKilometr + "-" + PlotBoundaries.pipeTwoKilometr + " км).xlsx";
            //richTextBox4.AppendText(Environment.NewLine + saveFileDialog1.FileName);
            worksheet.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            //worksheet2.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);

            workBook.Close(false, Type.Missing, Type.Missing);
            app.Quit();


            /*}*/
        }
        private void exportAnomalylogToIUST(MGVTD mGVTD)//Создание Excel файла с журналом дефектов для ИУС Т
        {
            object Nothing = System.Reflection.Missing.Value;
            var app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Add(Nothing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets[1];

            worksheet.Name = "AnomalyLog";

            // Write data
            worksheet.Cells[1, 1] = "№";//
            worksheet.Cells[1, 2] = "Номер трубы";//
            worksheet.Cells[1, 3] = "Тип дефекта";//          
            worksheet.Cells[1, 4] = "Код дефекта";//
            worksheet.Cells[1, 5] = "Расстояние от кольцевого шва";//
            worksheet.Cells[1, 6] = "Расстояние от продольного шва";//
            worksheet.Cells[1, 7] = "Начальный угол";//
            worksheet.Cells[1, 8] = "Длина, мм";//
            worksheet.Cells[1, 9] = "Ширина, мм";//
            worksheet.Cells[1, 10] = "Глубина, мм";//
            worksheet.Cells[1, 11] = "Положение";//наружный-внутренний
            worksheet.Cells[1, 12] = "Расположение дефекта на трубе";//
            worksheet.Cells[1, 13] = "Уровень опасности";//
            worksheet.Cells[1, 14] = "Требуется дополнительное обследование";//
            worksheet.Cells[1, 15] = "Комментарий";//
            int strNunber = 2;
            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)
            {
                worksheet.Cells[strNunber, 1] = i + 1;//"№";//
                worksheet.Cells[strNunber, 2] = mGVTD.anomalyLogLineS[i].pipeNumber;//"Номер трубы"
                worksheet.Cells[strNunber, 3] = mGVTD.anomalyLogLineS[i].defectType;//"Тип дефекта";          
                worksheet.Cells[strNunber, 4] = mGVTD.anomalyLogLineS[i].defectCode;// "Код дефекта";
                worksheet.Cells[strNunber, 5] = Math.Round(mGVTD.anomalyLogLineS[i].distanceFromTransverseWeldIUST, 3);// "Расстояние от кольцевого шва";
                worksheet.Cells[strNunber, 6] = Math.Round(mGVTD.anomalyLogLineS[i].distanceFromLongitudinalWeld, 3);// "Расстояние от продольного шва";
                worksheet.Cells[strNunber, 7] = mGVTD.anomalyLogLineS[i].start_angle;// "Начальный угол";

                if (mGVTD.anomalyLogLineS[i].length > 0)
                {
                    if (mGVTD.anomalyLogLineS[i].defectCode.Contains("ANCW") == false)
                    {
                        worksheet.Cells[strNunber, 8] = Math.Round(mGVTD.anomalyLogLineS[i].length, 3);// "Длина, мм";
                    }
                }
                if (mGVTD.anomalyLogLineS[i].widht > 0)
                {
                    worksheet.Cells[strNunber, 9] = Math.Round(mGVTD.anomalyLogLineS[i].widht, 3);// "Ширина, мм";
                }
                if (mGVTD.anomalyLogLineS[i].depthInMm > 0)
                {
                    worksheet.Cells[strNunber, 10] = Math.Round(mGVTD.anomalyLogLineS[i].depthInMm, 3);// "Глубина, мм";
                }

                worksheet.Cells[strNunber, 11] = mGVTD.anomalyLogLineS[i].inside_or_outside;// "Положение";//наружный-внутренний
                worksheet.Cells[strNunber, 12] = mGVTD.anomalyLogLineS[i].defect_location;// "Расположение дефекта на трубе";
                worksheet.Cells[strNunber, 13] = mGVTD.anomalyLogLineS[i].danger_level;// "Уровень опасности";
                worksheet.Cells[strNunber, 14] = "";// "Требуется дополнительное обследование";
                worksheet.Cells[strNunber, 15] = mGVTD.anomalyLogLineS[i].note;// "Комментарий";

                strNunber++;
            }

            /*string filename = textBox_adres.Text + mGVTD.pipelineInfo.pipelineName + "-" + PlotBoundaries.pipeNumberOnePipeLog+"-"+ PlotBoundaries.pipeNumberOnePipeLog+"xlsx";
            worksheet.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(false, Type.Missing, Type.Missing);
            app.Quit();*/

            // Show save file dialog
            //SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            /*if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {*/

            String dir = textBox_adres.Text;
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            saveFileDialog1.FileName = textBox_adres.Text + mGVTD.pipelineInfo.pipelineName + "-" + " Журнал_дефектов_ИУС_Т.xlsx";
            //richTextBox4.AppendText(Environment.NewLine + saveFileDialog1.FileName);
            try
            {
                worksheet.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception)
            {

            }
            //worksheet2.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);

            //workBook.Close(false, Type.Missing, Type.Missing);
            //app.Quit();


            /*}*/
        }

        private void exportPipeLogToIUST(MGVTD mGVTD)//Создание Excel файла с трубным журналом для ИУС Т
        {
            object Nothing = System.Reflection.Missing.Value;
            var app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Add(Nothing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets[1];

            worksheet.Name = "PipeLog";

            // Write data
            worksheet.Cells[1, 1] = "№";//
            worksheet.Cells[1, 2] = "Номер трубы";//
            worksheet.Cells[1, 3] = "Начало";//          
            worksheet.Cells[1, 4] = "Географические координаты начала трубы";//
            worksheet.Cells[1, 5] = "Длина трубы";//
            worksheet.Cells[1, 6] = "Тип трубы";//
            worksheet.Cells[1, 7] = "Диаметр";//
            worksheet.Cells[1, 8] = "Толщина стенки измеренная";//
            worksheet.Cells[1, 9] = "Ориентация первого шва";//
            worksheet.Cells[1, 10] = "Ориентация второго шва";//
            worksheet.Cells[1, 11] = "Овализация трубы";//
            worksheet.Cells[1, 12] = "Радиус упругого изгиба";//
            worksheet.Cells[1, 13] = "Процент повреждения изоляции";//
            worksheet.Cells[1, 14] = "Комментарий";//
            worksheet.Cells[1, 15] = "Завод - изготовитель";//
            worksheet.Cells[1, 16] = "Класс прочности стали трубы";//
            worksheet.Cells[1, 17] = "Коэф.надежности по внутреннему давлению";//
            worksheet.Cells[1, 18] = "Коэф-т надежности по назн.трубопровода";//
            worksheet.Cells[1, 19] = "Коэффициент надежности по материалу k1";//
            worksheet.Cells[1, 20] = "Марка стали";//
            worksheet.Cells[1, 21] = "Наличие наружного балластного покрытия";//
            worksheet.Cells[1, 22] = "Проект.категорийность уч.трубопровода";//
            worksheet.Cells[1, 23] = "Способ производства трубы";//
            worksheet.Cells[1, 24] = "Стандарт изготовления";//
            worksheet.Cells[1, 25] = "Страна завода-изготовителя";//
            worksheet.Cells[1, 26] = "Тип защитного покрытия";//
            int strNunber = 2;
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {

                worksheet.Cells[strNunber, 1] = i+1;//"№";//
                worksheet.Cells[strNunber, 2] = mGVTD.MGPipeS[i].pipeNumber;//"Номер трубы";//
                worksheet.Cells[strNunber, 3] = mGVTD.MGPipeS[i].odometrDist;// "Начало";//          
                //worksheet.Cells[strNunber, 4] = "Географические координаты начала трубы";//
                worksheet.Cells[strNunber, 5] = mGVTD.MGPipeS[i].pipeLength;//"Длина трубы";//

                if (mGVTD.MGPipeS[i].isTwoJoint)
                {
                    worksheet.Cells[strNunber, 6] = "двухшовная";//"Тип трубы";//
                    worksheet.Cells[strNunber, 9] = mGVTD.MGPipeS[i].firstJointAngle;//"Ориентация первого шва";//
                    worksheet.Cells[strNunber, 10] = mGVTD.MGPipeS[i].secondJointAngle;//"Ориентация второго шва";//
                }
                else
                {
                    worksheet.Cells[strNunber, 6] = "одношовная";//"Тип трубы";//
                    worksheet.Cells[strNunber, 9] = mGVTD.MGPipeS[i].firstJointAngle;//"Ориентация первого шва";//
                }
                
                worksheet.Cells[strNunber, 7] = mGVTD.pipelineInfo.pipeDiameter;//"Диаметр";//
                worksheet.Cells[strNunber, 8] = mGVTD.MGPipeS[i].thikness;//"Толщина стенки измеренная";//



                //worksheet.Cells[strNunber, 11] = "Овализация трубы";//
                //worksheet.Cells[strNunber, 12] = "Радиус упругого изгиба";//
                //worksheet.Cells[strNunber, 13] = "Процент повреждения изоляции";//
                worksheet.Cells[strNunber, 14] = mGVTD.MGPipeS[i].note;//"Комментарий";//
                //worksheet.Cells[strNunber, 15] = "Завод - изготовитель";//
                worksheet.Cells[strNunber, 16] = textBox447.Text;//"Класс прочности стали трубы";//
                worksheet.Cells[strNunber, 17] = textBox444.Text;//"Коэф.надежности по внутреннему давлению";//
                worksheet.Cells[strNunber, 18] = textBox445.Text;//"Коэф-т надежности по назн.трубопровода";//
                worksheet.Cells[strNunber, 19] = textBox446.Text;//"Коэффициент надежности по материалу k1";//
                worksheet.Cells[strNunber, 20] = mGVTD.MGPipeS[i].steelGrade;//"Марка стали";//
                //worksheet.Cells[strNunber, 21] = "Наличие наружного балластного покрытия";//
                worksheet.Cells[strNunber, 22] = mGVTD.MGPipeS[i].pipelineSectionCategory;//"Проект.категорийность уч.трубопровода";//
                //worksheet.Cells[strNunber, 23] = "Способ производства трубы";//
                //worksheet.Cells[strNunber, 24] = "Стандарт изготовления";//
                //worksheet.Cells[strNunber, 25] = "Страна завода-изготовителя";//
                //worksheet.Cells[strNunber, 26] = "Тип защитного покрытия";//
                strNunber++;
            }

            /*string filename = textBox_adres.Text + mGVTD.pipelineInfo.pipelineName + "-" + PlotBoundaries.pipeNumberOnePipeLog+"-"+ PlotBoundaries.pipeNumberOnePipeLog+"xlsx";
            worksheet.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(false, Type.Missing, Type.Missing);
            app.Quit();*/

            // Show save file dialog
            //SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            /*if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {*/

            String dir = textBox_adres.Text;
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            saveFileDialog1.FileName = textBox_adres.Text + mGVTD.pipelineInfo.pipelineName + "-" + " Трубный_журнал_ИУС_Т.xlsx";
            //richTextBox4.AppendText(Environment.NewLine + saveFileDialog1.FileName);
            try
            {
                worksheet.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception)
            {

            }
            //worksheet2.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);

            //workBook.Close(false, Type.Missing, Type.Missing);
            //app.Quit();


            /*}*/
        }
        private void exportFurnishingLogToIUST(MGVTD mGVTD)//Создание Excel файла с журналом элементов обустройства для ИУС Т
        {
            object Nothing = System.Reflection.Missing.Value;
            var app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Add(Nothing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets[1];

            worksheet.Name = "furnishingsLog";
            double startKm = 0;
            try
            {
                startKm = Convert.ToDouble(textBox448.Text.Replace(".",","));
            }
            catch (Exception)
            {
                
            }
                
            // Write data
            worksheet.Cells[1, 1] = "№";//
            worksheet.Cells[1, 2] = "Километр расположения";//
            worksheet.Cells[1, 3] = "Тип";//          
            worksheet.Cells[1, 4] = "Начало";//
            worksheet.Cells[1, 5] = "Длина";//
            worksheet.Cells[1, 6] = "Начальный угол";//
            worksheet.Cells[1, 7] = "Конечный угол";//
            worksheet.Cells[1, 8] = "Описание";//
            worksheet.Cells[1, 9] = "Комментарий";//
            worksheet.Cells[1, 10] = "Номер трубы";//
           
            int strNunber = 2;
            for (int i = 0; i < mGVTD.furnishingsLogS.Count; i++)
            {
                double aa = 0.001;
                worksheet.Cells[strNunber, 1] = i + 1;//"№";//
                worksheet.Cells[strNunber, 2] = startKm + aa * mGVTD.furnishingsLogS[i].odometrDist;//"Километр расположения";//
                worksheet.Cells[strNunber, 3] = mGVTD.furnishingsLogS[i].typeForIUST;//"Тип";//          
                worksheet.Cells[strNunber, 4] = Convert.ToString(1000 * mGVTD.furnishingsLogS[i].odometrDist);// "Начало";//
                worksheet.Cells[strNunber, 5] = mGVTD.furnishingsLogS[i].pipeLength;//"Длина";//
                //worksheet.Cells[strNunber, 6] = "Начальный угол";//
                //worksheet.Cells[strNunber, 7] = "Конечный угол";//
                //worksheet.Cells[strNunber, 8] = "Описание";//
                worksheet.Cells[strNunber, 9] = mGVTD.furnishingsLogS[i].note;//"Комментарий";//
                worksheet.Cells[strNunber, 10] = mGVTD.furnishingsLogS[i].pipeNumber;//"Номер трубы";//
                strNunber++;
            }

            /*string filename = textBox_adres.Text + mGVTD.pipelineInfo.pipelineName + "-" + PlotBoundaries.pipeNumberOnePipeLog+"-"+ PlotBoundaries.pipeNumberOnePipeLog+"xlsx";
            worksheet.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(false, Type.Missing, Type.Missing);
            app.Quit();*/

            // Show save file dialog
            //SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            /*if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {*/

            String dir = textBox_adres.Text;
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            saveFileDialog1.FileName = textBox_adres.Text + mGVTD.pipelineInfo.pipelineName + "-" + " журнал_особенностей_ИУС_Т.xlsx";
            //richTextBox4.AppendText(Environment.NewLine + saveFileDialog1.FileName);
            try
            {
                worksheet.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception)
            {

            }
            //worksheet2.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);

            //workBook.Close(false, Type.Missing, Type.Missing);
            //app.Quit();


            /*}*/
        }

        private void exportTo1C(MGVTD mGVTD, plotBoundaries PlotBoundaries)//Создание Excel файла для экспорта в Сонар
        {
            object Nothing = System.Reflection.Missing.Value;
            var app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook workBook = app.Workbooks.Add(Nothing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets[1];
            worksheet.Name = "SonarFormat";
            // Write data
            worksheet.Cells[1, 20] = mGVTD.pipelineInfo.pipelineName;//public string pipelineName;//трубопровод (название)
            worksheet.Cells[1, 21] = mGVTD.pipelineInfo.pipelineSection;//public string pipelineSection;//участок трубопровода
            worksheet.Cells[1, 22] = mGVTD.pipelineInfo.pipeDiameter;//public double pipeDiameter;//диаметр трубы            
            worksheet.Cells[1, 23] = mGVTD.pipelineInfo.examinationDate;//public string examinationDate;//дата обследования
            worksheet.Cells[1, 24] = mGVTD.pipelineInfo.designPressure;//public double designPressure;// проектное давление
            worksheet.Cells[1, 25] = mGVTD.pipelineInfo.operatingPressure;//public double operatingPressure;// рабочее давление
            worksheet.Cells[1, 26] = mGVTD.pipelineInfo.comissioningYear;//public string comissioningYear;//год ввода в экспуатацию
            worksheet.Cells[1, 27] = PlotBoundaries.pipeNumberOnePipeLog;//Первая граница участка
            worksheet.Cells[1, 28] = PlotBoundaries.pipeNumberTwoPipeLog;//вторая граница участка

            worksheet.Cells[1, 1] = "N_OSOB";
            worksheet.Cells[1, 2] = "N_SEK";
            worksheet.Cells[1, 3] = "L_ODOM";
            worksheet.Cells[1, 4] = "OTN_D";
            worksheet.Cells[1, 5] = "OSOB";
            worksheet.Cells[1, 6] = "L_SEK";
            worksheet.Cells[1, 7] = "T_ST";
            worksheet.Cells[1, 8] = "H_PROC";
            worksheet.Cells[1, 9] = "L_DEF";
            worksheet.Cells[1, 10] = "W_DEF";
            worksheet.Cells[1, 11] = "GRAD";
            worksheet.Cells[1, 12] = "TYPE";
            worksheet.Cells[1, 13] = "ABC";
            worksheet.Cells[1, 14] = "PRIM";
            worksheet.Cells[1, 15] = "PRIM";
            worksheet.Cells[1, 16] = "Category";
            worksheet.Cells[1, 17] = "Grade";
            worksheet.Cells[1, 18] = "yieldPoint";
            worksheet.Cells[1, 19] = "tensileStrength";

            int strNunber = 2;//номер строки в формируемой таблице
            for (int i = PlotBoundaries.pipeIdNumberOnePipeLog; i < PlotBoundaries.pipeIdNumberTwoPipeLog; i++)
            {
                string pipeNumber = mGVTD.MGPipeS[i].pipeNumber;//номер текущей трубы
                worksheet.Cells[strNunber, 1] = "";//N_OSOB
                worksheet.Cells[strNunber, 2] = mGVTD.MGPipeS[i].pipeNumber;//N_SEK
                worksheet.Cells[strNunber, 3] = mGVTD.MGPipeS[i].odometrDist;//L_ODOM
                worksheet.Cells[strNunber, 4] = "";//OTN_D
                worksheet.Cells[strNunber, 5] = mGVTD.MGPipeS[i].characterFeatures;//OSOB
                worksheet.Cells[strNunber, 6] = mGVTD.MGPipeS[i].pipeLength;//L_SEK
                worksheet.Cells[strNunber, 7] = mGVTD.MGPipeS[i].thikness;//T_ST
                worksheet.Cells[strNunber, 8] = 0;//H_PROC
                worksheet.Cells[strNunber, 9] = "";//L_DEF
                worksheet.Cells[strNunber, 10] = "";//W_DEF
                worksheet.Cells[strNunber, 11] = mGVTD.MGPipeS[i].jointAngle;//GRAD
                worksheet.Cells[strNunber, 12] = "";//TYPE
                worksheet.Cells[strNunber, 13] = "";//ABC
                worksheet.Cells[strNunber, 14] = mGVTD.MGPipeS[i].note;//PRIM
                worksheet.Cells[strNunber, 15] = "";//KBD
                worksheet.Cells[strNunber, 16] = mGVTD.MGPipeS[i].pipelineSectionCategory;//категория участка
                worksheet.Cells[strNunber, 17] = mGVTD.MGPipeS[i].steelGrade;//марка стали
                worksheet.Cells[strNunber, 18] = mGVTD.MGPipeS[i].yieldPoint;//предел текучести
                worksheet.Cells[strNunber, 19] = mGVTD.MGPipeS[i].tensileStrength;//предел прочности
                strNunber++;
                for (int j = 0; j < mGVTD.anomalyLogLineS.Count; j++)
                {
                    if (String.Equals(pipeNumber, mGVTD.anomalyLogLineS[j].pipeNumber))
                    {
                        if (String.Equals(mGVTD.anomalyLogLineS[j].featuresCharacter, mGVTD.MGPipeS[i].characterFeatures) == false)
                        {
                            worksheet.Cells[strNunber, 1] = j;//N_OSOB
                            worksheet.Cells[strNunber, 2] = mGVTD.anomalyLogLineS[j].pipeNumber;//N_SEK
                            worksheet.Cells[strNunber, 3] = mGVTD.anomalyLogLineS[j].odometrDist;//L_ODOM
                            worksheet.Cells[strNunber, 4] = mGVTD.anomalyLogLineS[j].distanceFromTransverseWeld;//OTN_D
                            worksheet.Cells[strNunber, 5] = mGVTD.anomalyLogLineS[j].featuresCharacter;//OSOB
                            worksheet.Cells[strNunber, 6] = mGVTD.MGPipeS[i].pipeLength;//L_SEK
                            worksheet.Cells[strNunber, 7] = mGVTD.MGPipeS[i].thikness;//T_ST
                            worksheet.Cells[strNunber, 8] = mGVTD.anomalyLogLineS[j].depthInProcent;//H_PROC
                            worksheet.Cells[strNunber, 9] = mGVTD.anomalyLogLineS[j].length;//L_DEF
                            worksheet.Cells[strNunber, 10] = mGVTD.anomalyLogLineS[j].widht;//W_DEF
                            worksheet.Cells[strNunber, 11] = mGVTD.anomalyLogLineS[j].featuresOrientation;//GRAD
                            worksheet.Cells[strNunber, 12] = mGVTD.anomalyLogLineS[j].extOrInt;//TYPE
                            worksheet.Cells[strNunber, 13] = mGVTD.anomalyLogLineS[j].defectAssessment;//ABC
                            worksheet.Cells[strNunber, 14] = mGVTD.anomalyLogLineS[j].note;//PRIM
                            worksheet.Cells[strNunber, 15] = mGVTD.anomalyLogLineS[j].KBD;//KBD
                            strNunber++;
                        }
                    }
                }

            }
            /*string filename = "C:\\Users\\Nivhu\\Desktop\\" + mGVTD.pipelineInfo.pipelineName + "-" + PlotBoundaries.pipeNumberOnePipeLog+"-"+ PlotBoundaries.pipeNumberOnePipeLog;
            worksheet.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(false, Type.Missing, Type.Missing);
            app.Quit();*/

            // Show save file dialog
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                richTextBox4.AppendText(Environment.NewLine + saveFileDialog1.FileName);
                worksheet.SaveAs(saveFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
                workBook.Close(false, Type.Missing, Type.Missing);
                app.Quit();
            }
        }

        private async void Export_1C_button_Click_1(object sender, EventArgs e)
        {
            List<plotBoundaries> allPlots = getAllPlotBoundaries(mGVTD, allValves);

            for (int i = 0; i < allPlots.Count; i++)
            {
                await Task.Run(() => exportTo1C_automatic(mGVTD, allPlots[i]));
            }
            System.Diagnostics.Process.Start("explorer", textBox_adres.Text);
        }

        private void textBox316_TextChanged(object sender, EventArgs e)
        {

        }


        private void ButtonOpenPipeLogNPCVTD_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNamePipeLog = openFileDialog1.FileName;

                ButtonOpenPipeLogNPCVTD.BackColor = Color.Azure;
                //findStart();//поиск начала и конца журналов свойств труб и категорий
                //tableArdesTest();
            }
        }

        private void ButtonOpenDefectLogNPCVTD_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNameDefectLog = openFileDialog1.FileName;
                ButtonOpenDefectLogNPCVTD.BackColor = Color.Azure;
                //findStart();//поиск начала и конца журналов свойств труб и категорий
                //tableArdesTest();//
            }
        }

        private void ButtonOpenLineObjectsNPCVTD_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNameLineObjects = openFileDialog1.FileName;
                ButtonOpenLineObjectsNPCVTD.BackColor = Color.Azure;
                //findStart();//поиск начала и конца журналов свойств труб и категорий
                //tableArdesTest();
            }
        }

        private void get_report_inform_NPCVTD()
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileNamePipeLog, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = "Общая информация";//получаем название вкладки из формы импотра
            try
            {
                ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];
            }
            catch (Exception)//рудимент. может пригодиться в будущем, если будут варианты нейминга вкладок
            {
                ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2.Replace(".xlsx", "")];
                textBox170.Text = WorksheetName2.Replace(".xlsx", "");
            }
            string text = "";
            string test = Convert.ToString(ObjWorkSheet2.Cells[9, 2].Text).Substring(4, 1);
            if (test == " ")
            {
                text = Convert.ToString(ObjWorkSheet2.Cells[9, 2].Text).Substring(0, 4);
            }
            else if (test != " ")
            {
                text = Convert.ToString(ObjWorkSheet2.Cells[9, 2].Text).Substring(0, 3);
            }

            textBox_diameterNPCVTD.Text = text;//диаметр трубы 
            textBox_nameNPCVTD.Text = Convert.ToString(ObjWorkSheet2.Cells[6, 2].Text);//наименование газопровода
            textBox_ploteNPCVTD.Text = Convert.ToString(ObjWorkSheet2.Cells[7, 2].Text);//наименование участка
            textBox_dateNPCVTD.Text = Convert.ToString(ObjWorkSheet2.Cells[4, 2].Text);//дата составления отчета
            //textBox_pressureNPCVTD = Convert.ToString(ObjWorkSheet2.Cells[4, 2].Text);//
        }
        private void get_report_inform_BHTTS()
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileNamePipeLog, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox379.Text;//получаем название вкладки из формы импотра

            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];
            //Газопровод - отвод к ГРС-16 пгт. Алексеевка, Ду800 Участок: 0 – 0.8 км
            string text = Convert.ToString(ObjWorkSheet2.Cells[1, 1].Text);
            string pipename = "0";
            string diameter = "0";
            string plote = "0";
            string dateCreationFile = "0";

            int DiameterStartPosition = text.IndexOf("Ду", 0);

            if (DiameterStartPosition > -1)
            {
                int DiameterFinishPosition = text.IndexOf(" ", DiameterStartPosition);
                pipename = text.Substring(0, DiameterStartPosition - 1);
                diameter = text.Substring(DiameterStartPosition + 2, 4).Replace(" ", ""); ;
                plote = text.Substring(DiameterFinishPosition + 1, text.Length - DiameterFinishPosition - 1);
                dateCreationFile = System.IO.File.GetLastWriteTime(fileNamePipeLog).ToString();
            }


            textBox_diam_BHTTS.Text = diameter;
            textBox_pipeline_BHTTS.Text = pipename;
            textBox_plot_BHTTS.Text = plote;
            textBox_date_BHTTS.Text = dateCreationFile;

            richTextBox6.AppendText(Environment.NewLine + "============================================");
            richTextBox6.AppendText(Environment.NewLine + "=Получены сведения об обследованном участке=");
            richTextBox6.AppendText(Environment.NewLine + "============================================");
        }
        private void CheckIndexesNPCVTD_Click(object sender, EventArgs e)//проверка адресации НПЦ ВТД
        {
            get_report_inform_NPCVTD();//получаем данные об отчете из отчета
            tableArdesTestPipeLogNPCVTD();
            tableArdesTestDefectLogNPCVTD();
            if (String.IsNullOrEmpty(fileNameLineObjects) == false)
            {
                tableArdesTestLineObjectsNPCVTD();
            }
        }

        private void ReloadDiameterNPCVTD_Click(object sender, EventArgs e)
        {
            //get_report_inform_NPCVTD();//получаем данные об отчете из отчета
            operatingReadToClassPipeInfoNPCVTD();
        }
        private void tableArdesTestPipeLogNPCVTD()//(для трубника НПЦВТД)метод для проверки правильности адресации ячеек и заполнения экземпляра класса numbersOfColumns()
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileNamePipeLog, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox250.Text;//получаем название вкладки из формы импотра
            try
            {
                ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];
            }
            catch (Exception)
            {
                ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2.Replace(".xlsx", "")];
                textBox250.Text = WorksheetName2.Replace(".xlsx", "");
            }



            //получаем номера столбцов для "трубного журлала"
            int columnNumber1 = Convert.ToInt16(textBox249.Text);//номер трубы
            int columnNumber2 = Convert.ToInt16(textBox248.Text);//дист
            int columnNumber3 = Convert.ToInt16(textBox247.Text);//толщина
            int columnNumber4 = Convert.ToInt16(textBox246.Text);//Длина трубы
            int columnNumber6 = Convert.ToInt16(textBox252.Text);//Характер особ.
            int columnNumber16 = Convert.ToInt16(textBox254.Text);//Ориент.//предел текучести            
            int columnNumber13 = Convert.ToInt16(textBox255.Text);//примечание//комментарий
            int columnNumber14 = Convert.ToInt16(textBox258.Text);//Категория
            int columnNumber15 = Convert.ToInt16(textBox260.Text);//Предел прочности
            //int columnNumber15 = Convert.ToInt16(textBox434.Text);//Предел прочности

            //выводим значения соответствующих ячеек для проверки
            textBox244.Text = Convert.ToString(ObjWorkSheet2.Cells[5, columnNumber1].Text);//номер трубы
            textBox243.Text = Convert.ToString(ObjWorkSheet2.Cells[5, columnNumber2].Text);//дист
            textBox242.Text = Convert.ToString(ObjWorkSheet2.Cells[5, columnNumber3].Text);//толщина
            textBox241.Text = Convert.ToString(ObjWorkSheet2.Cells[5, columnNumber4].Text);//Длина трубы
            textBox251.Text = Convert.ToString(ObjWorkSheet2.Cells[5, columnNumber6].Text);//Характер особ.
            textBox253.Text = Convert.ToString(ObjWorkSheet2.Cells[5, columnNumber16].Text);//Ориент.//предел текучести
            textBox256.Text = Convert.ToString(ObjWorkSheet2.Cells[5, columnNumber13].Text);//примечание//комментарий
            textBox257.Text = Convert.ToString(ObjWorkSheet2.Cells[5, columnNumber14].Text);//Категория
            textBox259.Text = Convert.ToString(ObjWorkSheet2.Cells[5, columnNumber15].Text);//Предел прочности


            //заполняем ссылки на номера столбцов для "трубного журлала"
            NumbersOfColumns.columnNumber1 = Convert.ToInt16(textBox249.Text);//номер трубы
            NumbersOfColumns.columnNumber2 = Convert.ToInt16(textBox248.Text);//дист
            NumbersOfColumns.columnNumber3 = Convert.ToInt16(textBox247.Text);//толщина
            NumbersOfColumns.columnNumber4 = Convert.ToInt16(textBox246.Text);//Длина трубы
            NumbersOfColumns.columnNumber6 = Convert.ToInt16(textBox252.Text);//Характер особ.
            NumbersOfColumns.columnNumber16 = Convert.ToInt16(textBox254.Text);//Ориент.//предел текучести            
            NumbersOfColumns.columnNumber13 = Convert.ToInt16(textBox255.Text);//примечание//марка стали
            NumbersOfColumns.columnNumber14 = Convert.ToInt16(textBox258.Text);//категория в трубном журнале
            NumbersOfColumns.columnNumber15 = Convert.ToInt16(textBox260.Text);//предел прочности в трубном журнале


            ObjExcel.Quit();


        }
        private void tableArdesTestPipeLogBHTTS()//(для трубника БХТТС)метод для проверки правильности адресации ячеек и заполнения экземпляра класса numbersOfColumns()
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileNamePipeLog, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet2;
            string WorksheetName2 = textBox379.Text;//получаем название вкладки из формы импотра
            ObjWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName2];

            //получаем номера столбцов для "трубного журлала"
            NumbersOfColumns.featuresNumber_BHTTS = Convert.ToInt16(textBox349.Text);//Номер особенности///1
            NumbersOfColumns.pipeNumber_BHTTS = Convert.ToInt16(textBox350.Text);//номер трубы///2
            NumbersOfColumns.odometrDist_BHTTS = Convert.ToInt16(textBox351.Text);//дистанция по одометру///3
            NumbersOfColumns.distanceFromReferencePoints_BHTTS = Convert.ToInt16(textBox352.Text);//расстояние от реперных точек///4distanceToNextReferencePoints_BHTTS
            NumbersOfColumns.distanceToNextReferencePoints_BHTTS = Convert.ToInt16(textBox353.Text);//расстояние от реперных точек///4distanceToNextReferencePoints_BHTTS
            NumbersOfColumns.featuresCharacter_BHTTS = Convert.ToInt16(textBox354.Text);//характер особенности///6
            NumbersOfColumns.distanceFromTransverseWeld_BHTTS = Convert.ToInt16(textBox355.Text);//расстояние от поперечного шва, м///7
            NumbersOfColumns.featuresOrientation_BHTTS = Convert.ToInt16(textBox356.Text);//угловая ориентация///8
            NumbersOfColumns.length_BHTTS = Convert.ToInt16(textBox357.Text);//длина///9
            NumbersOfColumns.widht_BHTTS = Convert.ToInt16(textBox358.Text);//ширина///10
            NumbersOfColumns.thikness_BHTTS = Convert.ToInt16(textBox359.Text);//толщина трубы///11
            NumbersOfColumns.depthInProcent_BHTTS = Convert.ToInt16(textBox360.Text);//глубина дефекта в процентах///12
            NumbersOfColumns.extOrInt_BHTTS = Convert.ToInt16(textBox361.Text);//характер локаизации(внутри или снаружи)///13
            NumbersOfColumns.note_BHTTS = Convert.ToInt16(textBox362.Text);//Примечание///14
            NumbersOfColumns.defectVanishDate = NumbersOfColumns.note_BHTTS + 1;
            int number_of_string = Convert.ToInt16(textBox112.Text);//получаем из формы номер строки для тестирования
            //выводим значения соответствующих ячеек для проверки
            textBox334.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.featuresNumber_BHTTS].Text);
            textBox335.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.pipeNumber_BHTTS].Text);
            textBox336.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.odometrDist_BHTTS].Text);
            textBox337.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.distanceFromReferencePoints_BHTTS].Text);
            textBox338.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.distanceToNextReferencePoints_BHTTS].Text);
            textBox339.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.featuresCharacter_BHTTS].Text);
            textBox340.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.distanceFromTransverseWeld_BHTTS].Text);
            textBox341.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.featuresOrientation_BHTTS].Text);
            textBox342.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.length_BHTTS].Text);
            textBox343.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.widht_BHTTS].Text);
            textBox344.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.thikness_BHTTS].Text);
            textBox345.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.depthInProcent_BHTTS].Text);
            textBox346.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.extOrInt_BHTTS].Text);
            textBox347.Text = Convert.ToString(ObjWorkSheet2.Cells[number_of_string, NumbersOfColumns.note_BHTTS].Text);

            ObjExcel.Quit();


        }

        private void tableArdesTestDefectLogNPCVTD()//(для дефектов НПЦВТД)метод для проверки правильности адресации ячеек и заполнения экземпляра класса numbersOfColumns()
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileNameDefectLog, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);


            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet3;
            string WorksheetName3 = textBox274.Text;//получаем название вкладки из формы импотра


            try
            {
                ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3];
            }
            catch (Exception)
            {
                ObjWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName3.Replace(".xlsx", "")];
                textBox274.Text = WorksheetName3.Replace(".xlsx", "");
            }



            //номера столбцов для "журлала аномалий"
            int column2Number1 = Convert.ToInt16(textBox271.Text);//дист по одом
            //int column2Number2 = Convert.ToInt16(textBox194.Text);//толщ
            int column2Number3 = Convert.ToInt16(textBox288.Text);//Расст. от ПОПШ
            int column2Number4 = Convert.ToInt16(textBox289.Text);//расст. от реперной т.
            int column2Number5 = Convert.ToInt16(textBox290.Text);//Характ. особ.
            int column2Number6 = Convert.ToInt16(textBox291.Text);//Класс размера
            int column2Number7 = Convert.ToInt16(textBox292.Text);//Ориентац
            int column2Number8 = Convert.ToInt16(textBox293.Text);//Длина
            int column2Number9 = Convert.ToInt16(textBox294.Text);//ширина
            int column2Number10 = Convert.ToInt16(textBox295.Text);//d %
            //int column2Number11 = Convert.ToInt16(textBox185.Text);//d мм
            //int column2Number12 = Convert.ToInt16(textBox184.Text);//Тип пол.
            int column2Number13 = Convert.ToInt16(textBox298.Text);//КБД
            int column2Number14 = Convert.ToInt16(textBox299.Text);//Оценка
            //int column2Number15 = Convert.ToInt16(textBox81.Text);
            //int column2Number16 = Convert.ToInt16(textBox82.Text);
            //int column2Number17 = Convert.ToInt16(textBox83.Text);
            //int column2Number18 = Convert.ToInt16(textBox84.Text);
            int column2Number19 = Convert.ToInt16(textBox272.Text);//дата устранения
            int column2Number20 = Convert.ToInt16(textBox315.Text);//для номера трубы

            textBox275.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number1].Text);//дист по одом
            //textBox198.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number2].Text);//толщ
            textBox276.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number3].Text);//Расст. от ПОПШ
            textBox277.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number4].Text);//расст. от реперной т.
            textBox278.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number5].Text);//Характ. особ.
            textBox279.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number6].Text);//Класс размера
            textBox280.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number7].Text);//Ориентац
            textBox281.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number8].Text);//Длина
            textBox282.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number9].Text);//ширина
            textBox283.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number10].Text);//d %
            //textBox207.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number11].Text);//d мм
            //textBox208.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number12].Text);//Тип пол.
            textBox286.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number13].Text);//КБД
            textBox287.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number14].Text);//Оценка
            //textBox60.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number15].Text);
            //textBox61.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number16].Text);
            //textBox62.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number17].Text);
            //textBox63.Text = Convert.ToString(ObjWorkSheet3.Cells[3, column2Number18].Text);
            textBox273.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number19].Text);//дата устранения
            textBox314.Text = Convert.ToString(ObjWorkSheet3.Cells[5, column2Number20].Text);//для номера трубы



            //номера столбцов для "журлала аномалий"
            NumbersOfColumns.column2Number1 = Convert.ToInt16(textBox271.Text);//дист по одом
            //NumbersOfColumns.column2Number2 = Convert.ToInt16(textBox194.Text);//толщ
            NumbersOfColumns.column2Number3 = Convert.ToInt16(textBox288.Text);//Расст. от ПОПШ
            NumbersOfColumns.column2Number4 = Convert.ToInt16(textBox289.Text);//расст. от реперной т.
            NumbersOfColumns.column2Number5 = Convert.ToInt16(textBox290.Text);//Характ. особ.
            NumbersOfColumns.column2Number6 = Convert.ToInt16(textBox291.Text);//Класс размера
            NumbersOfColumns.column2Number7 = Convert.ToInt16(textBox292.Text);//Ориентац
            NumbersOfColumns.column2Number8 = Convert.ToInt16(textBox293.Text);//Длина
            NumbersOfColumns.column2Number9 = Convert.ToInt16(textBox294.Text);//ширина
            NumbersOfColumns.column2Number10 = Convert.ToInt16(textBox295.Text);//d %
            //NumbersOfColumns.column2Number11 = Convert.ToInt16(textBox186.Text);//d мм
            //NumbersOfColumns.column2Number12 = Convert.ToInt16(textBox184.Text);//Тип пол.
            NumbersOfColumns.column2Number13 = Convert.ToInt16(textBox298.Text);//КБД
            NumbersOfColumns.column2Number14 = Convert.ToInt16(textBox299.Text);//Оценка
            //NumbersOfColumns.column2Number15 = Convert.ToInt16(textBox81.Text);
            //NumbersOfColumns.column2Number16 = Convert.ToInt16(textBox82.Text);
            //NumbersOfColumns.column2Number17 = Convert.ToInt16(textBox83.Text);
            //NumbersOfColumns.column2Number18 = Convert.ToInt16(textBox84.Text);
            NumbersOfColumns.column2Number19 = Convert.ToInt16(textBox272.Text);//дата устранения
            NumbersOfColumns.column2Number20 = Convert.ToInt16(textBox315.Text);//для номера трубы



            ObjExcel.Quit();


        }
        private void tableArdesTestLineObjectsNPCVTD()//(для объектов НПЦВТД)метод для проверки правильности адресации ячеек и заполнения экземпляра класса numbersOfColumns()
        {
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(fileNameLineObjects, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet4;
            string WorksheetName4 = textBox316.Text;//получаем название вкладки из формы импотра

            try
            {
                ObjWorkSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName4];
            }
            catch (Exception)
            {
                ObjWorkSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[WorksheetName4.Replace(".xlsx", "")];
                textBox316.Text = WorksheetName4.Replace(".xlsx", "");
            }

            //номера столбцов для "журнала элементов обустройства"
            //int column3Number1 = Convert.ToInt16(textBox96.Text);
            int column3Number2 = Convert.ToInt16(textBox321.Text);//номер трубы
            int column3Number3 = Convert.ToInt16(textBox322.Text);//дистанция по одометру
            //int column3Number4 = Convert.ToInt16(textBox99.Text);
            //int column3Number5 = Convert.ToInt16(textBox100.Text);
            int column3Number6 = Convert.ToInt16(textBox323.Text);//тип особенности
            int column3Number7 = Convert.ToInt16(textBox324.Text);//обозначение
            //int column3Number8 = Convert.ToInt16(textBox103.Text);
            //int column3Number9 = Convert.ToInt16(textBox104.Text);
            //int column3Number10 = Convert.ToInt16(textBox105.Text);
            //int column3Number11 = Convert.ToInt16(textBox106.Text);
            //int column3Number12 = Convert.ToInt16(textBox107.Text);
            int column3Number13 = Convert.ToInt16(textBox326.Text);//примечание

            //textBox64.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number1].Text);
            textBox317.Text = Convert.ToString(ObjWorkSheet4.Cells[5, column3Number2].Text);//номер трубы
            textBox318.Text = Convert.ToString(ObjWorkSheet4.Cells[5, column3Number3].Text);//дистанция по одометру
            //textBox85.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number4].Text);
            //textBox86.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number5].Text);
            textBox319.Text = Convert.ToString(ObjWorkSheet4.Cells[5, column3Number6].Text);//тип особенности
            textBox320.Text = Convert.ToString(ObjWorkSheet4.Cells[5, column3Number7].Text);//обозначение
            //textBox89.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number8].Text);
            //textBox90.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number9].Text);
            //textBox91.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number10].Text);
            //textBox92.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number11].Text);
            //textBox93.Text = Convert.ToString(ObjWorkSheet4.Cells[2, column3Number12].Text);
            textBox325.Text = Convert.ToString(ObjWorkSheet4.Cells[5, column3Number13].Text);//примечание

            //номера столбцов для "журнала элементов обустройства"
            //NumbersOfColumns.column3Number1 = Convert.ToInt16(textBox96.Text);
            NumbersOfColumns.column3Number2 = Convert.ToInt16(textBox321.Text);//номер трубы
            NumbersOfColumns.column3Number3 = Convert.ToInt16(textBox322.Text);//дистанция по одометру
            //NumbersOfColumns.column3Number4 = Convert.ToInt16(textBox99.Text);
            //NumbersOfColumns.column3Number5 = Convert.ToInt16(textBox100.Text);
            NumbersOfColumns.column3Number6 = Convert.ToInt16(textBox323.Text);//тип особенности
            NumbersOfColumns.column3Number7 = Convert.ToInt16(textBox324.Text);//обозначение
            //NumbersOfColumns.column3Number8 = Convert.ToInt16(textBox103.Text);
            //NumbersOfColumns.column3Number9 = Convert.ToInt16(textBox104.Text);
            //NumbersOfColumns.column3Number10 = Convert.ToInt16(textBox105.Text);
            //NumbersOfColumns.column3Number11 = Convert.ToInt16(textBox106.Text);
            //NumbersOfColumns.column3Number12 = Convert.ToInt16(textBox107.Text);
            NumbersOfColumns.column3Number13 = Convert.ToInt16(textBox326.Text);//примечание

            ObjExcel.Quit();


        }

        private void autoLookColumnsNPCVTD_Click(object sender, EventArgs e)
        {
            lookingForNumbersOfColumnsNPCVTD();
            get_report_inform_NPCVTD();//получаем данные об отчете из отчета
            tableArdesTestPipeLogNPCVTD();
            tableArdesTestDefectLogNPCVTD();
            if (String.IsNullOrEmpty(fileNameLineObjects) == false)
            {
                tableArdesTestLineObjectsNPCVTD();
            }
        }

        private void ReadReportToMemoryNPCVTD_Click(object sender, EventArgs e)
        {
            shortTableExcelReadToClassNPCVTD();
        }

        private void Button_open_BHTTS_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNamePipeLog = openFileDialog1.FileName;

                Button_open_BHTTS.BackColor = Color.Azure;
            }
        }

        private void buttonAdressCheckBHTTS_Click(object sender, EventArgs e)
        {
            get_report_inform_BHTTS();
            tableArdesTestPipeLogBHTTS();
        }

        private void button_pipeinfo_to_memory_BHTTS_Click(object sender, EventArgs e)
        {
            operatingReadToClassPipeInfoBHTTS();
        }
        private void textBox349_TextChanged(object sender, EventArgs e)
        {
        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
        private void button9_Click(object sender, EventArgs e)
        {
            mGVTD = OperatingReadToClassPipeLogAutoFinBHTTS(fileNamePipeLog, NumbersOfColumns);//трубный журнал и журнал аномалий
            List<furnishingsLog> furnishingsLogS_ = FirnishingLogVirtual(mGVTD);//Заполняем журнал элементов обустройства на основе данных трубного журнала
            mGVTD.furnishingsLogS = furnishingsLogS_;
            mGVTD.pipelineInfo = operatingReadToClassPipeInfoBHTTS();
        }

        private void ArrangementElements_Click(object sender, EventArgs e)
        {
        }
        private void button8_Click(object sender, EventArgs e)
        {
            string filename = DIrectory.Text + "VTD.xlsx";
            pipeSectionS = readToClassPipeSectionLog(filename);
            setComboBoxes(pipeSectionS);

            /*if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filename = openFileDialog1.FileName;

                button8.BackColor = Color.Azure;
                pipeSectionS = readToClassPipeSectionLog(filename);
                setComboBoxes(pipeSectionS);
            }*/
        }

        private void checkItems_for_LPU()//заполняем комбобоксы только значениями, которые есть в выбранном ЛПУ
        {
            MG_Check.Items.Clear();
            pipelineSection_Check.Items.Clear();
            for (int i = 0; i < pipeSectionS.Count; i++)
            {
                if (String.Equals(LPUMG_Check.SelectedItem, pipeSectionS[i].LPUMG_name))
                {
                    bool mark = true;
                    for (int j = 0; j < MG_Check.Items.Count; j++)
                    {
                        if (MG_Check.Items[j].Equals(pipeSectionS[i].pipelineName))
                        {
                            mark = false;
                        }
                    }
                    if (mark)
                    {
                        MG_Check.Items.Add(pipeSectionS[i].pipelineName);
                    }

                    mark = true;
                    for (int j = 0; j < pipelineSection_Check.Items.Count; j++)
                    {
                        if (pipelineSection_Check.Items[j].Equals(pipeSectionS[i].pipelineSection))
                        {
                            mark = false;
                        }
                    }
                    if (mark)
                    {
                        pipelineSection_Check.Items.Add(pipeSectionS[i].pipelineSection);
                    }

                }
            }


            /*                LPUMG_Check.Items.Add(pipeSectionS[i].LPUMG_name);
                MG_Check.Items.Add(pipeSectionS[i].pipelineName);
                pipelineSection_Check.Items.Add(pipeSectionS[i].pipelineSection);*/

        }
        private void checkItems_for_MG()//заполняем комбобоксы только значениями, которые есть в выбранном МГ
        {
            LPUMG_Check.Items.Clear();
            pipelineSection_Check.Items.Clear();
            for (int i = 0; i < pipeSectionS.Count; i++)
            {
                if (String.Equals(MG_Check.SelectedItem, pipeSectionS[i].pipelineName))
                {
                    bool mark = true;
                    for (int j = 0; j < LPUMG_Check.Items.Count; j++)
                    {
                        if (LPUMG_Check.Items[j].Equals(pipeSectionS[i].LPUMG_name))
                        {
                            mark = false;
                        }
                    }
                    if (mark)
                    {
                        LPUMG_Check.Items.Add(pipeSectionS[i].LPUMG_name);
                    }

                    mark = true;
                    for (int j = 0; j < pipelineSection_Check.Items.Count; j++)
                    {
                        if (pipelineSection_Check.Items[j].Equals(pipeSectionS[i].pipelineSection))
                        {
                            mark = false;
                        }
                    }
                    if (mark)
                    {
                        pipelineSection_Check.Items.Add(pipeSectionS[i].pipelineSection);
                    }

                }
            }


            /*  LPUMG_Check.Items.Add(pipeSectionS[i].LPUMG_name);
                MG_Check.Items.Add(pipeSectionS[i].pipelineName);
                pipelineSection_Check.Items.Add(pipeSectionS[i].pipelineSection);*/

        }
        private void LPUMG_Check_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckedItem.LPUMG_name = Convert.ToString(LPUMG_Check.SelectedItem);
            checkItems_for_LPU();
            //richTextBox7.AppendText(Environment.NewLine + LPUMG_Check.SelectedIndex);
        }

        private void MG_Check_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckedItem.pipelineName = Convert.ToString(MG_Check.SelectedItem);
            checkItems_for_MG();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            setComboBoxes(pipeSectionS);
        }

        private async void button10_Click(object sender, EventArgs e)
        {
            get_MG_ID(pipeSectionS);
            //string fileName = DIrectory.Text + MG_ID.Text;
            string fileName = DIrectory.Text + "VTD.xlsx";
            await Task.Run(() => mGVTD = OperatingReadToClassPipeLogHimself(fileName, MG_ID.Text));
            //List<ListOfTees> listOfTees = GetListOfTees(fileName);//читаем таблицу принадлежности тройников
            //mGVTD = VDTWhithoutRepareDefects(mGVTD);//удаляет сведения об устраненных дефектах
            textBox131.Text = mGVTD.MGPipeS[0].pipeNumber;
            textBox136.Text = mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].pipeNumber;
            textBox_contractor.Text = mGVTD.pipelineInfo.contractor;
            textBox_examinationDate.Text = mGVTD.pipelineInfo.examinationDate;
            textBox_diameter.Text = Convert.ToString(mGVTD.pipelineInfo.pipeDiameter);
            textBox_comissioningYear.Text = mGVTD.pipelineInfo.comissioningYear;
            textBox_operatingPressure.Text = Convert.ToString(mGVTD.pipelineInfo.operatingPressure);
            textBox_pipelineCount.Text = Convert.ToString(mGVTD.MGPipeS.Count);
            richTextBox7.AppendText(Environment.NewLine + "Подрядчик: " + mGVTD.pipelineInfo.contractor);
            richTextBox7.AppendText(Environment.NewLine + "Дата обследования: " + mGVTD.pipelineInfo.examinationDate);
        }

        private void pipelineSection_Check_SelectedIndexChanged(object sender, EventArgs e)
        {
            CheckedItem.pipelineSection = Convert.ToString(pipelineSection_Check.SelectedItem);
        }

        private void DrawMyDiagram(MGVTD mGVTD)//График длин катушек.
        {
            Graphics g = pictureBox1.CreateGraphics();
            g.Clear(Color.White);
            // Создаем объекты-кисти для закрашивания фигур
            SolidBrush myCorp = new SolidBrush(Color.DarkMagenta);
            SolidBrush myTrum = new SolidBrush(Color.DarkOrchid);
            SolidBrush myTrub = new SolidBrush(Color.DeepPink);
            SolidBrush mySeа = new SolidBrush(Color.Blue);
            //Выбираем перо myPen желтого цвета толщиной в 2 пикселя:
            Pen myWindYellow = new Pen(Color.Yellow, 2);
            //Выбираем перо myPen черного цвета толщиной в 2 пикселя:
            Pen myWindBlack = new Pen(Color.Black, 1);
            Pen myWindBlackBold = new Pen(Color.Black, 2);
            //Выбираем перо myPen Голубого цвета толщиной в 2 пикселя:
            Pen myWindBlue = new Pen(Color.Blue, 1);
            //Выбираем перо myPen Красного цвета толщиной в 2 пикселя:
            Pen myWindRed = new Pen(Color.Red, 1);
            //Выбираем перо myPen зелёного цвета толщиной в 2 пикселя:
            Pen myWindGreen = new Pen(Color.Lime, 1);
            //g.DrawLine(myWindRed, 100, 100, 200, 200);
            //запрашиваем ширину и высоту окна с графикой
            float strip = 20;//высота полосы для служебной информации
            float gWidht = (float)pictureBox1.Width;
            float gHeight2 = (float)pictureBox1.Height;
            float gHeight = (float)pictureBox1.Height - strip;
            float halfgWidht = (float)pictureBox1.Width / 2;
            float halfgHeight = pictureBox1.Height / 2;
            float minim = Math.Min(gWidht, gHeight);
            float xCoeff = 1500 / minim;
            float deltaX = gWidht / mGVTD.MGPipeS.Count;
            double maxHeight = 15;//максимальное ожидаемое значение аргумента (будет соответствовать высоте окна гистограммы)
            double trend = gHeight - gHeight * 0.001 * mGVTD.pipelineInfo.pipeDiameter / maxHeight;

            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                double dd = gHeight * mGVTD.MGPipeS[i].pipeLength / maxHeight;
                float alpha = (float)dd;
                if (mGVTD.MGPipeS[i].pipeLength < 0.001 * mGVTD.pipelineInfo.pipeDiameter)
                {
                    g.DrawLine(myWindRed, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
                else
                {
                    g.DrawLine(myWindGreen, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
            }
            g.DrawLine(myWindBlack, ((float)(0)), ((float)(trend)), ((float)(gWidht)), ((float)(trend)));

            double fullPath = 0;//полный путь
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                fullPath = mGVTD.MGPipeS[i].odometrDist- mGVTD.MGPipeS[0].odometrDist;
                //richTextBox7.AppendText("/"+fullPath);
                double a = 1000;
                double b = 100;

                double z = fullPath / b - Math.Round(fullPath / b);//остаток от целочисленного деления
                if (Math.Abs(z) < 0.06)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlue, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 10)));
                }

                z = fullPath / a - Math.Round(fullPath / a);
                if (Math.Abs(z) < 0.007)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlackBold, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 20)));
                }
            }
        }
        private void DrawMyDiagram2(MGVTD mGVTD)//График поврежденности от коррозии.
        {
            Graphics g = pictureBox2.CreateGraphics();
            g.Clear(Color.White);
            // Создаем объекты-кисти для закрашивания фигур
            SolidBrush myCorp = new SolidBrush(Color.DarkMagenta);
            SolidBrush myTrum = new SolidBrush(Color.DarkOrchid);
            SolidBrush myTrub = new SolidBrush(Color.DeepPink);
            SolidBrush mySeа = new SolidBrush(Color.Blue);
            //Выбираем перо myPen желтого цвета толщиной в 2 пикселя:
            Pen myWindYellow = new Pen(Color.Yellow, 2);
            //Выбираем перо myPen черного цвета толщиной в 2 пикселя:
            Pen myWindBlack = new Pen(Color.Black, 1);
            Pen myWindBlackBold = new Pen(Color.Black, 2);
            //Выбираем перо myPen Голубого цвета толщиной в 2 пикселя:
            Pen myWindBlue = new Pen(Color.Blue, 1);
            //Выбираем перо myPen Красного цвета толщиной в 2 пикселя:
            Pen myWindRed = new Pen(Color.Red, 1);
            //Выбираем перо myPen зелёного цвета толщиной в 2 пикселя:
            Pen myWindGreen = new Pen(Color.Lime, 3);
            //g.DrawLine(myWindRed, 100, 100, 200, 200);
            //запрашиваем ширину и высоту окна с графикой
            float strip = 20;//высота полосы для служебной информации
            float gWidht = (float)pictureBox1.Width;
            float gHeight2 = (float)pictureBox1.Height;
            float gHeight = (float)pictureBox1.Height - strip;
            float halfgWidht = (float)pictureBox1.Width / 2;
            float halfgHeight = pictureBox1.Height / 2;
            float minim = Math.Min(gWidht, gHeight);
            float xCoeff = 1500 / minim;
            float deltaX = gWidht / mGVTD.MGPipeS.Count;
            double maxHeight = 1;//максимальное ожидаемое значение аргумента (будет соответствовать высоте окна гистограммы)
            double trend = gHeight - gHeight * 0.001 * mGVTD.pipelineInfo.pipeDiameter / maxHeight;



            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                double dd = gHeight * mGVTD.MGPipeS[i].MaximumDamageCorr / maxHeight;
                float alpha = (float)dd;
                if (mGVTD.MGPipeS[i].MaximumDamageCorr >= 0.1)
                {
                    g.DrawLine(myWindRed, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
                else
                {
                    g.DrawLine(myWindGreen, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
            }
            g.DrawLine(myWindBlack, ((float)(0)), ((float)(trend)), ((float)(gWidht)), ((float)(trend)));

            double fullPath = 0;//полный путь
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                fullPath = mGVTD.MGPipeS[i].odometrDist - mGVTD.MGPipeS[0].odometrDist;
                double a = 1000;
                double b = 100;

                double z = fullPath / b - Math.Round(fullPath / b);//остаток от целочисленного деления
                if (Math.Abs(z) < 0.06)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlue, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 10)));
                }

                z = fullPath / a - Math.Round(fullPath / a);
                if (Math.Abs(z) < 0.007)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlackBold, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 20)));
                }
            }
        }
        private void DrawMyDiagram3(MGVTD mGVTD)//График поврежденности от вмятин.
        {
            Graphics g = pictureBox3.CreateGraphics();
            g.Clear(Color.White);
            // Создаем объекты-кисти для закрашивания фигур
            SolidBrush myCorp = new SolidBrush(Color.DarkMagenta);
            SolidBrush myTrum = new SolidBrush(Color.DarkOrchid);
            SolidBrush myTrub = new SolidBrush(Color.DeepPink);
            SolidBrush mySeа = new SolidBrush(Color.Blue);
            //Выбираем перо myPen желтого цвета толщиной в 2 пикселя:
            Pen myWindYellow = new Pen(Color.Yellow, 2);
            //Выбираем перо myPen черного цвета толщиной в 2 пикселя:
            Pen myWindBlack = new Pen(Color.Black, 1);
            Pen myWindBlackBold = new Pen(Color.Black, 2);
            //Выбираем перо myPen Голубого цвета толщиной в 2 пикселя:
            Pen myWindBlue = new Pen(Color.Blue, 1);
            //Выбираем перо myPen Красного цвета толщиной в 2 пикселя:
            Pen myWindRed = new Pen(Color.Red, 1);
            //Выбираем перо myPen зелёного цвета толщиной в 2 пикселя:
            Pen myWindGreen = new Pen(Color.Lime, 3);
            //g.DrawLine(myWindRed, 100, 100, 200, 200);
            //запрашиваем ширину и высоту окна с графикой
            float strip = 20;//высота полосы для служебной информации
            float gWidht = (float)pictureBox1.Width;
            float gHeight2 = (float)pictureBox1.Height;
            float gHeight = (float)pictureBox1.Height - strip;
            float halfgWidht = (float)pictureBox1.Width / 2;
            float halfgHeight = pictureBox1.Height / 2;
            float minim = Math.Min(gWidht, gHeight);
            float xCoeff = 1500 / minim;
            float deltaX = gWidht / mGVTD.MGPipeS.Count;
            double maxHeight = 3;//максимальное ожидаемое значение аргумента (будет соответствовать высоте окна гистограммы)
            double trend = gHeight - gHeight * 0.001 * mGVTD.pipelineInfo.pipeDiameter / maxHeight;



            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                double dd = gHeight * mGVTD.MGPipeS[i].MaximumDentDamage / maxHeight;
                float alpha = (float)dd;
                if (mGVTD.MGPipeS[i].MaximumDentDamage >= 0.1)
                {
                    g.DrawLine(myWindRed, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
                else
                {
                    g.DrawLine(myWindGreen, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
            }
            //g.DrawLine(myWindBlack, ((float)(0)), ((float)(trend)), ((float)(gWidht)), ((float)(trend)));

            double fullPath = 0;//полный путь
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                fullPath = mGVTD.MGPipeS[i].odometrDist - mGVTD.MGPipeS[0].odometrDist;
                double a = 1000;
                double b = 100;

                double z = fullPath / b - Math.Round(fullPath / b);//остаток от целочисленного деления
                if (Math.Abs(z) < 0.06)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlue, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 10)));
                }

                z = fullPath / a - Math.Round(fullPath / a);
                if (Math.Abs(z) < 0.007)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlackBold, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 20)));
                }
            }
        }
        private void DrawMyDiagram4(MGVTD mGVTD)//График поврежденности от дефектов сварнных швов.
        {
            Graphics g = pictureBox4.CreateGraphics();
            g.Clear(Color.White);
            // Создаем объекты-кисти для закрашивания фигур
            SolidBrush myCorp = new SolidBrush(Color.DarkMagenta);
            SolidBrush myTrum = new SolidBrush(Color.DarkOrchid);
            SolidBrush myTrub = new SolidBrush(Color.DeepPink);
            SolidBrush mySeа = new SolidBrush(Color.Blue);
            //Выбираем перо myPen желтого цвета толщиной в 2 пикселя:
            Pen myWindYellow = new Pen(Color.Yellow, 2);
            //Выбираем перо myPen черного цвета толщиной в 2 пикселя:
            Pen myWindBlack = new Pen(Color.Black, 1);
            Pen myWindBlackBold = new Pen(Color.Black, 2);
            //Выбираем перо myPen Голубого цвета толщиной в 2 пикселя:
            Pen myWindBlue = new Pen(Color.Blue, 1);
            //Выбираем перо myPen Красного цвета толщиной в 2 пикселя:
            Pen myWindRed = new Pen(Color.Red, 1);
            //Выбираем перо myPen зелёного цвета толщиной в 2 пикселя:
            Pen myWindGreen = new Pen(Color.Lime, 3);
            //g.DrawLine(myWindRed, 100, 100, 200, 200);
            //запрашиваем ширину и высоту окна с графикой
            float strip = 20;//высота полосы для служебной информации
            float gWidht = (float)pictureBox1.Width;
            float gHeight2 = (float)pictureBox1.Height;
            float gHeight = (float)pictureBox1.Height - strip;
            float halfgWidht = (float)pictureBox1.Width / 2;
            float halfgHeight = pictureBox1.Height / 2;
            float minim = Math.Min(gWidht, gHeight);
            float xCoeff = 1500 / minim;
            float deltaX = gWidht / mGVTD.MGPipeS.Count;
            double maxHeight = 3;//максимальное ожидаемое значение аргумента (будет соответствовать высоте окна гистограммы)
            double trend = gHeight - gHeight * 0.001 * mGVTD.pipelineInfo.pipeDiameter / maxHeight;



            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                double dd = gHeight * mGVTD.MGPipeS[i].MaximumJoinDamage / maxHeight;
                float alpha = (float)dd;
                if (mGVTD.MGPipeS[i].MaximumJoinDamage >= 0.1)
                {
                    g.DrawLine(myWindRed, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
                else
                {
                    g.DrawLine(myWindGreen, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
            }
            //g.DrawLine(myWindBlack, ((float)(0)), ((float)(trend)), ((float)(gWidht)), ((float)(trend)));

            double fullPath = 0;//полный путь
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                fullPath = mGVTD.MGPipeS[i].odometrDist- mGVTD.MGPipeS[0].odometrDist;
                double a = 1000;
                double b = 100;

                double z = fullPath / b - Math.Round(fullPath / b);//остаток от целочисленного деления
                if (Math.Abs(z) < 0.06)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlue, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 10)));
                }

                z = fullPath / a - Math.Round(fullPath / a);
                if (Math.Abs(z) < 0.007)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlackBold, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 20)));
                }
            }
        }
        private void DrawMyDiagram5(MGVTD mGVTD)//График поврежденности от дефектов сварнных швов.
        {
            Graphics g = pictureBox5.CreateGraphics();
            g.Clear(Color.White);
            // Создаем объекты-кисти для закрашивания фигур
            SolidBrush myCorp = new SolidBrush(Color.DarkMagenta);
            SolidBrush myTrum = new SolidBrush(Color.DarkOrchid);
            SolidBrush myTrub = new SolidBrush(Color.DeepPink);
            SolidBrush mySeа = new SolidBrush(Color.Blue);
            //Выбираем перо myPen желтого цвета толщиной в 2 пикселя:
            Pen myWindYellow = new Pen(Color.Yellow, 2);
            //Выбираем перо myPen черного цвета толщиной в 2 пикселя:
            Pen myWindBlack = new Pen(Color.Black, 1);
            Pen myWindBlackBold = new Pen(Color.Black, 2);
            //Выбираем перо myPen Голубого цвета толщиной в 2 пикселя:
            Pen myWindBlue = new Pen(Color.Blue, 1);
            //Выбираем перо myPen Красного цвета толщиной в 2 пикселя:
            Pen myWindRed = new Pen(Color.Red, 1);
            //Выбираем перо myPen зелёного цвета толщиной в 2 пикселя:
            Pen myWindGreen = new Pen(Color.Lime, 1);
            //g.DrawLine(myWindRed, 100, 100, 200, 200);
            //запрашиваем ширину и высоту окна с графикой
            float strip = 20;//высота полосы для служебной информации
            float gWidht = (float)pictureBox1.Width;
            float gHeight2 = (float)pictureBox1.Height;
            float gHeight = (float)pictureBox1.Height - strip;
            float halfgWidht = (float)pictureBox1.Width / 2;
            float halfgHeight = pictureBox1.Height / 2;
            float minim = Math.Min(gWidht, gHeight);
            float xCoeff = 1500 / minim;
            float deltaX = gWidht / mGVTD.MGPipeS.Count;
            double maxCorrpro = 0;
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                if (maxCorrpro < mGVTD.MGPipeS[i].MaximumCorrProcent)
                {
                    maxCorrpro = mGVTD.MGPipeS[i].MaximumCorrProcent;
                }
            }

            //richTextBox7.AppendText(""+ maxCorrpro);
            double maxHeight = 110;//максимальное ожидаемое значение аргумента (будет соответствовать высоте окна гистограммы)
            double trend = gHeight - gHeight * 0.001 * mGVTD.pipelineInfo.pipeDiameter / maxHeight;



            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                double dd = gHeight * mGVTD.MGPipeS[i].MaximumCorrProcent / maxHeight;
                float alpha = (float)dd;
                if (mGVTD.MGPipeS[i].MaximumCorrProcent >= 30)
                {
                    g.DrawLine(myWindRed, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
                else
                {
                    g.DrawLine(myWindGreen, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
            }
            //g.DrawLine(myWindBlack, ((float)(0)), ((float)(trend)), ((float)(gWidht)), ((float)(trend)));

            double fullPath = 0;//полный путь
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                fullPath = mGVTD.MGPipeS[i].odometrDist - mGVTD.MGPipeS[0].odometrDist;
                double a = 1000;
                double b = 100;

                double z = fullPath / b - Math.Round(fullPath / b);//остаток от целочисленного деления
                if (Math.Abs(z) < 0.06)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlue, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 10)));
                }

                z = fullPath / a - Math.Round(fullPath / a);
                if (Math.Abs(z) < 0.007)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlackBold, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 20)));
                }
            }
        }
        private void DrawMyDiagram6(MGVTD mGVTD)//График поврежденности от дефектов сварнных швов.
        {
            Graphics g = pictureBox6.CreateGraphics();
            g.Clear(Color.White);
            // Создаем объекты-кисти для закрашивания фигур
            SolidBrush myCorp = new SolidBrush(Color.DarkMagenta);
            SolidBrush myTrum = new SolidBrush(Color.DarkOrchid);
            SolidBrush myTrub = new SolidBrush(Color.DeepPink);
            SolidBrush mySeа = new SolidBrush(Color.Blue);
            //Выбираем перо myPen желтого цвета толщиной в 2 пикселя:
            Pen myWindYellow = new Pen(Color.Yellow, 2);
            //Выбираем перо myPen черного цвета толщиной в 2 пикселя:
            Pen myWindBlack = new Pen(Color.Black, 1);
            Pen myWindBlackBold = new Pen(Color.Black, 2);
            //Выбираем перо myPen Голубого цвета толщиной в 2 пикселя:
            Pen myWindBlue = new Pen(Color.Blue, 1);
            //Выбираем перо myPen Красного цвета толщиной в 2 пикселя:
            Pen myWindRed = new Pen(Color.Red, 3);
            //Выбираем перо myPen зелёного цвета толщиной в 2 пикселя:
            Pen myWindGreen = new Pen(Color.Lime, 1);
            //g.DrawLine(myWindRed, 100, 100, 200, 200);
            //запрашиваем ширину и высоту окна с графикой
            float strip = 20;//высота полосы для служебной информации
            float gWidht = (float)pictureBox1.Width;
            float gHeight2 = (float)pictureBox1.Height;
            float gHeight = (float)pictureBox1.Height - strip;
            float halfgWidht = (float)pictureBox1.Width / 2;
            float halfgHeight = pictureBox1.Height / 2;
            float minim = Math.Min(gWidht, gHeight);
            float xCoeff = 1500 / minim;
            float deltaX = gWidht / mGVTD.MGPipeS.Count;
            double maxHeight = 110;//максимальное ожидаемое значение аргумента (будет соответствовать высоте окна гистограммы)
            double trend = gHeight - gHeight * 0.001 * mGVTD.pipelineInfo.pipeDiameter / maxHeight;
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)//перегоняем одометр так, чтобы отсчет участка начинался от нуля
            {
                mGVTD.MGPipeS[i].odometrDist = mGVTD.MGPipeS[i].odometrDist - mGVTD.MGPipeS[0].odometrDist;
            }


            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                double dd = gHeight * (100 - mGVTD.MGPipeS[i].MaximumCorrProcent) / maxHeight;
                float alpha = (float)dd;
                if (100 - mGVTD.MGPipeS[i].MaximumCorrProcent <= 70)
                {
                    g.DrawLine(myWindRed, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
                else
                {
                    g.DrawLine(myWindGreen, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
            }
            //g.DrawLine(myWindBlack, ((float)(0)), ((float)(trend)), ((float)(gWidht)), ((float)(trend)));

            double fullPath = 0;//полный путь
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                fullPath = mGVTD.MGPipeS[i].odometrDist - mGVTD.MGPipeS[0].odometrDist;
                double a = 1000;
                double b = 100;
                fullPath = fullPath - mGVTD.MGPipeS[0].odometrDist;
                double z = fullPath / b - Math.Round(fullPath / b);//остаток от целочисленного деления
                if (Math.Abs(z) < 0.06)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlue, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 10)));
                }

                z = fullPath / a - Math.Round(fullPath / a);
                if (Math.Abs(z) < 0.007)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlackBold, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 20)));
                }
            }
        }
        private void DrawMyDiagram7(MGVTD mGVTD)//График поврежденности от дефектов сварнных швов.
        {
            Graphics g = pictureBox7.CreateGraphics();
            g.Clear(Color.White);
            // Создаем объекты-кисти для закрашивания фигур
            SolidBrush myCorp = new SolidBrush(Color.DarkMagenta);
            SolidBrush myTrum = new SolidBrush(Color.DarkOrchid);
            SolidBrush myTrub = new SolidBrush(Color.DeepPink);
            SolidBrush mySeа = new SolidBrush(Color.Blue);
            //Выбираем перо myPen желтого цвета толщиной в 2 пикселя:
            Pen myWindYellow = new Pen(Color.Yellow, 2);
            //Выбираем перо myPen черного цвета толщиной в 2 пикселя:
            Pen myWindBlack = new Pen(Color.Black, 1);
            Pen myWindBlackBold = new Pen(Color.Black, 2);
            //Выбираем перо myPen Голубого цвета толщиной в 2 пикселя:
            Pen myWindBlue = new Pen(Color.Blue, 1);
            //Выбираем перо myPen Красного цвета толщиной в 2 пикселя:
            Pen myWindRed = new Pen(Color.Red, 3);
            //Выбираем перо myPen зелёного цвета толщиной в 2 пикселя:
            Pen myWindGreen = new Pen(Color.Lime, 1);
            //g.DrawLine(myWindRed, 100, 100, 200, 200);
            //запрашиваем ширину и высоту окна с графикой
            float strip = 20;//высота полосы для служебной информации
            float gWidht = (float)pictureBox1.Width;
            float gHeight2 = (float)pictureBox1.Height;
            float gHeight = (float)pictureBox1.Height - strip;
            float halfgWidht = (float)pictureBox1.Width / 2;
            float halfgHeight = pictureBox1.Height / 2;
            float minim = Math.Min(gWidht, gHeight);
            float xCoeff = 1500 / minim;
            float deltaX = gWidht / mGVTD.MGPipeS.Count;
            double maxHeight = 20;//максимальное ожидаемое значение аргумента (будет соответствовать высоте окна гистограммы)
            double trend = gHeight - gHeight * 0.001 * mGVTD.pipelineInfo.pipeDiameter / maxHeight;
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)//перегоняем одометр так, чтобы отсчет участка начинался от нуля
            {
                mGVTD.MGPipeS[i].odometrDist = mGVTD.MGPipeS[i].odometrDist - mGVTD.MGPipeS[0].odometrDist;
            }


            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                double dd = gHeight * (mGVTD.MGPipeS[i].residualResource) / maxHeight;
                float alpha = (float)dd;
                //g.DrawLine(myWindRed, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));

                if (mGVTD.MGPipeS[i].residualResource <= 10)
                {
                    g.DrawLine(myWindRed, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
                else
                {
                    g.DrawLine(myWindGreen, ((float)(deltaX + i * deltaX)), ((float)(gHeight)), ((float)(deltaX + i * deltaX)), ((float)(gHeight - alpha)));
                }
            }
            //g.DrawLine(myWindBlack, ((float)(0)), ((float)(trend)), ((float)(gWidht)), ((float)(trend)));

            double fullPath = 0;//полный путь
            for (int i = 0; i < mGVTD.MGPipeS.Count; i++)
            {
                fullPath = mGVTD.MGPipeS[i].odometrDist - mGVTD.MGPipeS[0].odometrDist;
                double a = 1000;
                double b = 100;
                fullPath = fullPath - mGVTD.MGPipeS[0].odometrDist;
                double z = fullPath / b - Math.Round(fullPath / b);//остаток от целочисленного деления
                if (Math.Abs(z) < 0.06)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlue, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 10)));
                }

                z = fullPath / a - Math.Round(fullPath / a);
                if (Math.Abs(z) < 0.007)
                {
                    float X0 = (float)(gWidht * fullPath / (mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].odometrDist - mGVTD.MGPipeS[0].odometrDist));
                    g.DrawLine(myWindBlackBold, ((float)(X0)), ((float)(gHeight2)), ((float)(X0)), ((float)(gHeight2 - 20)));
                }
            }
        }
        MGVTD cropPipesLog(MGVTD input, int start, int stop)
        {
            MGVTD result = new MGVTD();
  
            for (int i = start; i < stop; i++)
            {
                result.MGPipeS.Add(input.MGPipeS[i]);
            }
            result.pipelineInfo = input.pipelineInfo;
            /*for (int i = 0; i < result.MGPipeS.Count; i++)
            {
                result.MGPipeS[i].odometrDist = result.MGPipeS[i].odometrDist - result.MGPipeS[0].odometrDist;
            }*/
            return result;
        }
        private void DrawDiagrams()
        {
            int min = Math.Min(hScrollBar1.Value, hScrollBar2.Value);
            int max = Math.Max(hScrollBar1.Value, hScrollBar2.Value);

            if (max - min > 10)
            {
                MGVTD VTD = cropPipesLog(mGVTD, min, max);
                VTD.MGPipeS[0].odometrDist = 0;
                for (int i = 1; i < VTD.MGPipeS.Count; i++)
                {
                    VTD.MGPipeS[i].odometrDist = VTD.MGPipeS[i - 1].odometrDist + VTD.MGPipeS[i].pipeLength;
                }
                DrawMyDiagram(VTD);
                DrawMyDiagram2(VTD);
                DrawMyDiagram3(VTD);
                DrawMyDiagram4(VTD);
                DrawMyDiagram5(VTD);
                DrawMyDiagram6(VTD);
                DrawMyDiagram7(VTD);
            }
        }

        private void button_DrawGraph_Click(object sender, EventArgs e)
        {
            DrawDiagrams();
        }
        List<ListOfTees> listOfTees = new List<ListOfTees>();
        List<ListOfTees> InfoOfTees = new List<ListOfTees>();
        
        private async void button1_Click_2(object sender, EventArgs e)
        {
            fileName = DIrectory.Text + "VTD.xlsx";
            await Task.Run(new Action(() => InfoOfTees = GetListOfNumbersTees(fileName)));//читаем таблицу принадлежности тройников GetListOfTees
            string fileName2 = DIrectory.Text + "tees.xlsx";
            await Task.Run(new Action(() => listOfTees = GetListOfTees(fileName2)));//читаем данные о тройниках
            listOfTees = SetIDtoTees(listOfTees, InfoOfTees);
        }

        private async void button14_Click(object sender, EventArgs e)
        {
            richTextBox7.Invoke(new Action(() => richTextBox7.Clear()));
            await Task.Run(new Action(() => listOfTees = getInfoAbouteTees(listOfTees, pipeSectionS)));
            for (int i = 0; i < listOfTees.Count; i++)
            {
                if (listOfTees[i].isSorted == false)
                {
                    richTextBox7.AppendText(Convert.ToString(listOfTees[i].teeName));
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            int a = 5;
            double b = 5.5;
            double c = a + b;
            richTextBox7.AppendText(Convert.ToString(c));
        }
        private void ConvertIusT_Click(object sender, EventArgs e)
        {
            mGVTD = SetPipeLengthToAnomalyLog(mGVTD);
            mGVTD = SetVolumesToAnomalyLogForIUST(mGVTD);
            mGVTD = setJointAnglesToMgPipesGPAS(mGVTD);
            mGVTD = setTypesForIUSTToFirnishingLog(mGVTD);
            //MessageBox.Show("OK!");
            exportAnomalylogToIUST(mGVTD);
            exportPipeLogToIUST(mGVTD);
            exportFurnishingLogToIUST(mGVTD);
        }
        private void testButton_Click(object sender, EventArgs e)
        {
            textBox437.Text = GetDefectTypeGPAS(textBox436.Text).defectType + "_" + GetDefectTypeGPAS(textBox436.Text).defectCode;
        }
        private void button13_Click(object sender, EventArgs e)//3.04 / -8.54
        {
            textBox437.Text = Convert.ToString(convertJointToHour(textBox436.Text));
        }
        private void button15_Click(object sender, EventArgs e)
        {
            textBox437.Text = Convert.ToString(GetStartAngle(textBox436.Text));
        }
        private void button16_Click(object sender, EventArgs e)
        {
            textBox437.Text = Convert.ToString(GetDistanceFromTranswersWeldGPAS(textBox436.Text));
        }
        private void button17_Click(object sender, EventArgs e)
        {
            //GetdistanceFromLongitudinalWeld
            textBox437.Text = Convert.ToString(GetdistanceFromLongitudinalWeld(textBox436.Text, 0, 1420));
        }
        double PvtdReport = 0;
        private async void getVtdData_Click(object sender, EventArgs e)
        {
            string fileName = DIrectory.Text + "VTD.xlsx";
            await Task.Run(() => mGVTD = OperatingReadToClassPipeLogHimself(fileName, textBox438.Text));
            textBox439.Text = mGVTD.pipelineInfo.pipelineName;
            textBox440.Text = mGVTD.pipelineInfo.pipelineSection;
            textBox441.Text = Convert.ToString(mGVTD.pipelineInfo.pipeDiameter);
            textBox443.Text = Convert.ToString(mGVTD.anomalyLogLineS.Count);

            textBox131.Text = mGVTD.MGPipeS[0].pipeNumber;
            textBox136.Text = mGVTD.MGPipeS[mGVTD.MGPipeS.Count - 1].pipeNumber;
            goEquation();
            textBox442.Text = Convert.ToString(Math.Round(PvtdReport,3));
            richTextBox2.AppendText(Environment.NewLine + "mGVTD.MGPipeS.Count=" + mGVTD.MGPipeS.Count); textBox441.Text = Convert.ToString(mGVTD.pipelineInfo.pipeDiameter);
            textBox_contractor.Text = mGVTD.pipelineInfo.contractor;
            textBox_examinationDate.Text = mGVTD.pipelineInfo.examinationDate;
            textBox_diameter.Text = Convert.ToString(mGVTD.pipelineInfo.pipeDiameter);
            textBox_comissioningYear.Text = mGVTD.pipelineInfo.comissioningYear;
            textBox_operatingPressure.Text = Convert.ToString(mGVTD.pipelineInfo.operatingPressure);
            textBox_pipelineCount.Text = Convert.ToString(mGVTD.MGPipeS.Count);
            richTextBox7.AppendText(Environment.NewLine + "Подрядчик: " + mGVTD.pipelineInfo.contractor);
            richTextBox7.AppendText(Environment.NewLine + "Дата обследования: " + mGVTD.pipelineInfo.examinationDate);
            /*textBoxGraphStart.Text = mGVTD.MGPipeS[0].pipeNumber;
            textBoxGraphStop.Text = mGVTD.MGPipeS[mGVTD.MGPipeS.Count].pipeNumber;*/
            mGVTD = GetMaximumValues(mGVTD);
            mGVTD = GetMaximumValuesCorrerionInProcent(mGVTD);
            mGVTD = GetCritikalThiknessForAll(mGVTD);
            SetScrollBars(mGVTD);
            DrawDiagrams();
        }
        private void SetScrollBars(MGVTD input)
        {
            hScrollBar1.Minimum = 1;
            hScrollBar1.Maximum = input.MGPipeS.Count;
            richTextBox2.AppendText(Environment.NewLine + "mGVTD.MGPipeS.Count=" + input.MGPipeS.Count);
            hScrollBar2.Minimum = 1;
            hScrollBar2.Maximum = input.MGPipeS.Count;
            hScrollBar1.Value = 1;
            hScrollBar2.Value = input.MGPipeS.Count;
            textBoxGraphStart.Text = String.Concat(mGVTD.MGPipeS[hScrollBar1.Value - 1].pipeNumber, "_", hScrollBar1.Value - 1);
            textBoxGraphStop.Text = String.Concat(mGVTD.MGPipeS[hScrollBar2.Value - 1].pipeNumber, "_", hScrollBar2.Value - 1);
        }
        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            textBoxGraphStart.Text = String.Concat(mGVTD.MGPipeS[hScrollBar1.Value - 1].pipeNumber, "_", hScrollBar1.Value - 1);
            //DrawDiagrams();
        }

        private void hScrollBar2_Scroll(object sender, ScrollEventArgs e)
        {
            textBoxGraphStop.Text = String.Concat(mGVTD.MGPipeS[hScrollBar2.Value - 1].pipeNumber, "_", hScrollBar2.Value - 1);
            //DrawDiagrams();
        }
        private void button18_Click(object sender, EventArgs e)
        {
            gMapControl1.MapProvider = GMapProviders.GoogleHybridMap;
            gMapControl1.Position = new PointLatLng(53.219527, 50.154535);
            
            gMapControl1.MinZoom = 5;
            gMapControl1.MaxZoom = 50;
            gMapControl1.Zoom = 10;
            gMapControl1.DragButton = MouseButtons.Left;
            markers.Clear();
        }
        
        private void button19_Click(object sender, EventArgs e)
        {
            setMarkersOfDefectsCorr(mGVTD);
        }
        GMapOverlay markers = new GMapOverlay("markers");
        private void setMarkersOfDefects(List<InsulationDefects> input)
        {
            gMapControl1.MapProvider = GMapProviders.GoogleHybridMap;
            gMapControl1.MinZoom = 5;
            gMapControl1.MaxZoom = 50;
            gMapControl1.Zoom = 15;
            gMapControl1.DragButton = MouseButtons.Left;
            //GMapOverlay markers = new GMapOverlay("markers");
            
            for (int i = 0; i < input.Count; i++)
            {                
                PointLatLng point = new PointLatLng(ConvertDegreeAngleToDoubleLat(input[i].defectCoordinates), ConvertDegreeAngleToDoubleLon(input[i].defectCoordinates));
                GMapMarker marker = new GMarkerGoogle(point, GMarkerGoogleType.red_dot);
                marker.ToolTipText = String.Concat(Convert.ToString(input[i].defectLength)," м");
                markers.Markers.Add(marker);
            }
                        gMapControl1.Overlays.Add(markers);
            
            int centrPoint = 0;
            if (input.Count%2==0)
            {
                centrPoint = input.Count / 2;
            }
            else
            {
                centrPoint = (input.Count-1) / 2;
            }
            gMapControl1.Position = new PointLatLng(ConvertDegreeAngleToDoubleLat(input[centrPoint].defectCoordinates), ConvertDegreeAngleToDoubleLon(input[centrPoint].defectCoordinates));
        }

        List<defectsOfInsulation> insulationDefectPipes = new List<defectsOfInsulation>();
        MGVTD setInsulationDefectsNew(MGVTD inputVTD, List<InsulationDefects> insulationDefects)
        {
            MGVTD result = new MGVTD();
            for (int i = 0; i < inputVTD.MGPipeS.Count; i++)
            {
                inputVTD.MGPipeS[i].isInsulationDefect = false;
            }
            for (int i = 0; i < insulationDefects.Count; i++)
            {
                int nearPipe = 2;
                double minimumDist = 10;
                bool mark = false;
                for (int j = 0; j < inputVTD.MGPipeS.Count; j++)
                {
                    GeoCoordinate defectPoint = new GeoCoordinate(ConvertDegreeAngleToDoubleLat(insulationDefects[i].defectCoordinates), ConvertDegreeAngleToDoubleLon(insulationDefects[i].defectCoordinates));
                    GeoCoordinate pipePoint = new GeoCoordinate(ConvertDegreeAngleToDouble(inputVTD.MGPipeS[j].Latitude), ConvertDegreeAngleToDouble(inputVTD.MGPipeS[j].Longitude));
                    double distanceTo = pipePoint.GetDistanceTo(defectPoint);

                        if (distanceTo< minimumDist)
                        {
                            minimumDist = distanceTo;
                            nearPipe = j;
                            mark = true;
                        }                    
                }

                
                if (mark)
                {
                    defectsOfInsulation defect = new defectsOfInsulation();
                    double defectLength = insulationDefects[i].defectLength;
                    defect.defectLength = insulationDefects[i].defectLength;
                    defect.pipeNumber = inputVTD.MGPipeS[nearPipe].pipeNumber;
                    List<int> pipesList = new List<int>();
                    while (defectLength > 0)
                    {
                        inputVTD.MGPipeS[nearPipe].isInsulationDefect = true;
                        defectLength -= inputVTD.MGPipeS[nearPipe].pipeLength;
                        
                        
                        if (nearPipe < inputVTD.MGPipeS.Count)
                        {
                            pipesList.Add(nearPipe);
                            nearPipe++;                            
                        }
                        else
                        {
                            defectLength = 0;
                        }
                    }
                    defect.numbersOfPipes = pipesList;
                    insulationDefectPipes.Add(defect);
                }
            }

            return result;
        }
        private void printInsulatoinDefects(MGVTD inputVTD, List<defectsOfInsulation> inputDefects)
        {
            for (int i = 0; i < inputDefects.Count; i++)
            {
                int numb = i + 1;
                richTextBox9.Invoke(new Action(() => richTextBox9.AppendText(Environment.NewLine + "Дефект №" + numb + ",длина:" + inputDefects[i].defectLength + " м, тр.№" + inputVTD.MGPipeS[inputDefects[i].numbersOfPipes[0]].pipeNumber+ "-тр.№" + inputVTD.MGPipeS[inputDefects[i].numbersOfPipes[inputDefects[i].numbersOfPipes.Count-1]].pipeNumber)));
            }
        }


        MGVTD setInsulationDefects(MGVTD inputVTD, List<InsulationDefects> insulationDefects)
        {
            MGVTD result = new MGVTD();
            for (int i = 0; i < inputVTD.MGPipeS.Count; i++)
            {
                inputVTD.MGPipeS[i].isInsulationDefect = false;
            }
            for (int i = 0; i < insulationDefects.Count; i++)
            {
                for (int j = 0; j < inputVTD.MGPipeS.Count; j++)
                {
                    GeoCoordinate defectPoint = new GeoCoordinate(ConvertDegreeAngleToDoubleLat(insulationDefects[i].defectCoordinates), ConvertDegreeAngleToDoubleLon(insulationDefects[i].defectCoordinates));
                    GeoCoordinate pipePoint = new GeoCoordinate(ConvertDegreeAngleToDouble(inputVTD.MGPipeS[j].Latitude), ConvertDegreeAngleToDouble(inputVTD.MGPipeS[j].Longitude));
                    double distanceTo = pipePoint.GetDistanceTo(defectPoint);
                    List<int> nearPipes = new List<int>();
                    if (distanceTo < 10)
                    {
                        nearPipes.Add(i);
                    }
                    
                    
                    if (distanceTo<10)
                    {
                        int a = j;
                        double defectLength = insulationDefects[i].defectLength;
                        while (defectLength>0)
                        {
                            inputVTD.MGPipeS[a].isInsulationDefect = true;
                            defectLength -= inputVTD.MGPipeS[a].pipeLength;

                            richTextBox9.Invoke(new Action(() => richTextBox9.AppendText(Environment.NewLine + inputVTD.MGPipeS[a].pipeNumber + "_" + insulationDefects[i].defectLength)));
                            if (a< inputVTD.MGPipeS.Count)
                            {
                                a++;
                            }
                            else
                            {
                                defectLength = 0;
                            }
                        }
                    }
                }
            }

            return result;
        }

        private void setMarkersOfPipes(MGVTD input)
        {
            gMapControl1.MapProvider = GMapProviders.GoogleHybridMap;
            gMapControl1.MinZoom = 5;
            gMapControl1.MaxZoom = 50;
            gMapControl1.Zoom = 15;
            gMapControl1.DragButton = MouseButtons.Left;
            //GMapOverlay markers = new GMapOverlay("markers");

            for (int i = 0; i < input.MGPipeS.Count; i++)
            {
                /*if (input.anomalyLogLineS[i].depthInProcent>0)
                {*/
                    PointLatLng point = new PointLatLng(ConvertDegreeAngleToDouble(input.MGPipeS[i].Latitude), ConvertDegreeAngleToDouble(input.MGPipeS[i].Longitude));
                    GMapMarker marker = new GMarkerGoogle(point, GMarkerGoogleType.green_dot);
                    markers.Markers.Add(marker);
               // }
            }
            gMapControl1.Overlays.Add(markers);
            gMapControl1.Position = new PointLatLng(ConvertDegreeAngleToDouble(input.MGPipeS[0].Latitude), ConvertDegreeAngleToDouble(input.MGPipeS[0].Longitude));
        }
        private void setMarkersOfDefectsCorr(MGVTD input)
        {
            gMapControl1.MapProvider = GMapProviders.GoogleHybridMap;
            gMapControl1.MinZoom = 5;
            gMapControl1.MaxZoom = 50;
            gMapControl1.Zoom = 15;
            gMapControl1.DragButton = MouseButtons.Left;
            

            for (int i = 0; i < input.anomalyLogLineS.Count; i++)
            {
                if (input.anomalyLogLineS[i].depthInProcent>Convert.ToDouble(textBox449.Text))
                {
                PointLatLng point = new PointLatLng(ConvertDegreeAngleToDouble(input.anomalyLogLineS[i].Latitude), ConvertDegreeAngleToDouble(input.anomalyLogLineS[i].Longitude));
                GMapMarker marker = new GMarkerGoogle(point, GMarkerGoogleType.yellow_dot);
                    marker.ToolTipText = String.Concat("Труба №", input.anomalyLogLineS[i].pipeNumber,", глубина дефекта:", input.anomalyLogLineS[i].depthInProcent, " %");
                markers.Markers.Add(marker);
                }
            }
            gMapControl1.Overlays.Add(markers);
            gMapControl1.Position = new PointLatLng(ConvertDegreeAngleToDouble(input.anomalyLogLineS[0].Latitude), ConvertDegreeAngleToDouble(input.anomalyLogLineS[0].Longitude));
        }
        private void button20_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog1.FileName;
                insulationDefects = GetInsulationDefectsFromFile(fileName);
            }
            setMarkersOfDefects(insulationDefects);
            setInsulationDefectsNew(mGVTD, insulationDefects);
            addLineInsulatoinDefects(mGVTD);
            printInsulatoinDefects(mGVTD, insulationDefectPipes);

        }
        private void addLine()//пример рисования линий на карте
        {
            GMapOverlay routes = new GMapOverlay("routes"); //Создаем объект наложения (Overlay)
            List<PointLatLng> points = new List<PointLatLng>(); //Создаем лист, где будут наши точки пути.
            points.Add(new PointLatLng(48.866383, 2.323575)); //Добавляем координаты
            points.Add(new PointLatLng(48.863868, 2.321554));
            points.Add(new PointLatLng(48.861017, 2.330030));
            GMapRoute route = new GMapRoute(points, "A walk in the park"); //Создаем из полученных точнек маршрут и даем ей имя.
            route.Stroke = new Pen(Color.Red, 3); //Задаем цвет и ширину линии
            routes.Routes.Add(route); //Добавляем на наш Overlay маршрут
            gMapControl1.Overlays.Add(routes); //Накладываем Overlay на карту.

            GMapOverlay markersOverlay = new GMapOverlay("marker"); //Создаем Overlay
            GMarkerGoogle markerStart = new GMarkerGoogle(points.FirstOrDefault(), GMarkerGoogleType.blue); //Создаем новую точку и даем ей координаты первого элемента из листа координат и синий цвет
            GMarkerGoogle markerEnd = new GMarkerGoogle(points.LastOrDefault(), GMarkerGoogleType.red); //Тоже самое, но красный цвет и последний из списка координат.
            markerStart.ToolTip = new GMapRoundedToolTip(markerStart); //Указываем тип всплывающей подсказки для точки старта
            markerEnd.ToolTip = new GMapBaloonToolTip(markerEnd); //Другой тип подсказки для точки окончания (для теста)
            markerStart.ToolTipText = "Точка старта"; //Текст всплывающих подсказок при наведении
            markerEnd.ToolTipText = "Точка окончания";

            markersOverlay.Markers.Add(markerStart); //Добавляем точки
            markersOverlay.Markers.Add(markerEnd); //В наш оверлей маркеров

            gMapControl1.Overlays.Add(markersOverlay); //Добавляем оверлей на карту
            gMapControl1.Position = new PointLatLng(48.861017, 2.330030);

        }
        private void addLineInsulatoinDefects(MGVTD input) 
        {
            GMapOverlay routes = new GMapOverlay("routes"); //Создаем объект наложения (Overlay)
            GMapOverlay markersOverlay = new GMapOverlay("marker"); //Создаем Overlay
            for (int i = 0; i < input.MGPipeS.Count-1; i++)
            {
                if (input.MGPipeS[i].isInsulationDefect)
                {
                    List<PointLatLng> points = new List<PointLatLng>(); //Создаем лист, где будут наши точки пути.
                    points.Add(new PointLatLng(ConvertDegreeAngleToDouble(input.MGPipeS[i].Latitude), ConvertDegreeAngleToDouble(input.MGPipeS[i].Longitude))); //Добавляем координаты
                    points.Add(new PointLatLng(ConvertDegreeAngleToDouble(input.MGPipeS[i+1].Latitude), ConvertDegreeAngleToDouble(input.MGPipeS[i+1].Longitude)));
                    GMapRoute route = new GMapRoute(points, "pipe"); //Создаем из полученных точнек маршрут и даем ей имя.
                    route.Stroke = new Pen(Color.Red, 3); //Задаем цвет и ширину линии
                    routes.Routes.Add(route); //Добавляем на наш Overlay маршрут


                    PointLatLng point = new PointLatLng(ConvertDegreeAngleToDouble(input.MGPipeS[i].Latitude), ConvertDegreeAngleToDouble(input.MGPipeS[i].Longitude));
                    GMapMarker markerStart = new GMarkerGoogle(point, GMarkerGoogleType.green_dot);

                    //PointLatLng pipePoint = new PointLatLng(ConvertDegreeAngleToDouble(input.MGPipeS[i].Latitude), ConvertDegreeAngleToDouble(input.MGPipeS[i].Longitude));
                    //GMarkerGoogle markerStart = new GMarkerGoogle(pipePoint, GMarkerGoogleType.blue); //Создаем новую точку и даем ей координаты первого элемента из листа координат и синий цвет
                    //GMarkerGoogle markerEnd = new GMarkerGoogle(points.LastOrDefault(), GMarkerGoogleType.red); //Тоже самое, но красный цвет и последний из списка координат.
                    //markerStart.ToolTip = new GMapRoundedToolTip(markerStart); //Указываем тип всплывающей подсказки для точки старта
                    //markerEnd.ToolTip = new GMapBaloonToolTip(markerEnd); //Другой тип подсказки для точки окончания (для теста)
                    markerStart.ToolTipText = input.MGPipeS[i].pipeNumber; //Текст всплывающих подсказок при наведении
                    //markerEnd.ToolTipText = "Точка окончания";

                    markersOverlay.Markers.Add(markerStart); //Добавляем точки
                    //markersOverlay.Markers.Add(markerEnd); //В наш оверлей маркеров
                }
            }
            gMapControl1.Overlays.Add(routes); //Накладываем Overlay на карту.
            gMapControl1.Overlays.Add(markersOverlay); //Добавляем оверлей на карту
        }

        private void button21_Click(object sender, EventArgs e)
        {
            setMarkersOfPipes(mGVTD);
        }

        private void button22_Click(object sender, EventArgs e)
        {
            //addLine();

        }
    }
}
