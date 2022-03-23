using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VTDinfo
{

    public partial class Form1
    {
        public DefectsTypes GetDefectTypeSOD(string defectName)
        {
            DefectsTypes types = new DefectsTypes();
            types.defectType = "аномалия";
            types.defectCode = "ANML";
            if (String.Equals(defectName, "Технологический дефект TECH"))
            {
                types.defectType = "аномалия";
                types.defectCode = "ANML";
            }
            if (String.Equals(defectName, "Аномалия кольцевого шва GWAN"))
            {
                types.defectType = "аномалия кольцевого шва";
                types.defectCode = "ANCW";
            }
            if (String.Equals(defectName, "Аномалия продольного шва LWAN"))
            {
                types.defectType = "аномалия продольного шва";
                types.defectCode = "ANLW";
            }
            if (String.Equals(defectName, "Вмятина DENT"))
            {
                types.defectType = "вмятина";
                types.defectCode = "DENT";
            }
            if (String.Equals(defectName, "Коррозия CORR"))
            {
                types.defectType = "коррозия";
                types.defectCode = "CORR";
            }
            if (String.Equals(defectName, "Механическое повреждение ARTD"))
            {
                types.defectType = "Механическое повреждение";
                types.defectCode = "MCIN";
            }
            if (String.Equals(defectName, "Заводской дефект MIAN"))
            {
                types.defectType = "Заводской дефект";
                types.defectCode = "FACT";
            }
            if (String.Equals(defectName, "Металл снаружи TMTM"))
            {
                types.defectType = "аномалия";
                types.defectCode = "ANML";
            }
            if (String.Equals(defectName, "Расслоение LAMI"))
            {
                types.defectType = "аномалия";
                types.defectCode = "ANML";
            }
            return types;
        }
        public DefectsTypes GetDefectTypeGPAS(string defectName)
        {
            DefectsTypes types = new DefectsTypes();
            types.defectType = "аномалия";
            types.defectCode = "ANML";
            if (String.Equals(defectName, "Технологический дефект"))
            {
                types.defectType = "аномалия";
                types.defectCode = "ANML";
            }
            if (String.Equals(defectName, "Аномалия кольцевого шва"))
            {
                types.defectType = "аномалия кольцевого шва";
                types.defectCode = "ANCW";
            }
            if (String.Equals(defectName, "Аномалия продольного шва"))
            {
                types.defectType = "аномалия продольного шва";
                types.defectCode = "ANLW";
            }
            if (String.Equals(defectName, "Вмятина"))
            {
                types.defectType = "вмятина";
                types.defectCode = "DENT";
            }
            if (String.Equals(defectName, "Кластер коррозии"))
            {
                types.defectType = "коррозия";
                types.defectCode = "CORR";
            }
            if (String.Equals(defectName, "Механическое повреждение"))
            {
                types.defectType = "Механическое повреждение";
                types.defectCode = "MCIN";
            }
            if (String.Equals(defectName, "Заводской дефект"))
            {
                types.defectType = "Заводской дефект";
                types.defectCode = "FACT";
            }
            if (String.Equals(defectName, "Коррозия"))
            {
                types.defectType = "коррозия";
                types.defectCode = "CORR";
            }
            if (String.Equals(defectName, "Металл снаружи"))
            {
                types.defectType = "аномалия";
                types.defectCode = "ANML";
            }
            if (String.Equals(defectName, "Несваренный стык патрона"))
            {
                types.defectType = "аномалия";
                types.defectCode = "ANML";
            }
            return types;
        }
        public double GetDistanceFromTranswersWeldGPAS(string distanceFromTransverseWeld)//расстояние от поперечного шва
        {
            double result = 0;
            int indexOfChar = 0;
            distanceFromTransverseWeld = distanceFromTransverseWeld.Replace(" ", "").Replace(".", ",");
            if (distanceFromTransverseWeld.Contains("/"))
            {
                indexOfChar = distanceFromTransverseWeld.IndexOf("/");
            }
            if (distanceFromTransverseWeld.Contains("-"))
            {
                indexOfChar = distanceFromTransverseWeld.IndexOf("-");
            }
            result = Convert.ToDouble(distanceFromTransverseWeld.Substring(0, indexOfChar-1));
            return result;
        }
        public double convertJointToHour(string inputAngle)//конвертация из формата чч:мм в десятичные часы.
        {
            double result = 0;
            int indexOfChar = inputAngle.IndexOf(":");
            int lengthOfString = inputAngle.Length;
            if (lengthOfString>0)
            {
                double hours = Convert.ToDouble(inputAngle.Substring(0, indexOfChar).Trim().Replace("-", ""));
                double minuts = Convert.ToDouble(inputAngle.Substring(indexOfChar + 1, lengthOfString - indexOfChar - 1).Trim());
                if (hours > 12)
                {
                    hours = hours - 12;
                }
                if (minuts > 60)
                {
                    minuts = 60;
                }
                result = Math.Round(hours + 0.01666666666 * minuts, 2);
            }
            return result;
        }
        public double GetdistanceFromLongitudinalWeld(string featuresOrientation, double clockOrientation,  double pipeDiameter)//расстояние от продольного шва
        {
            double result = 0;
            string substring = "";
            double koeff = (Math.PI * pipeDiameter) / 12;
            if (featuresOrientation.Contains("-"))
            {
                int indexOfChar = featuresOrientation.IndexOf("-");
                substring = featuresOrientation.Substring(0, indexOfChar);
                result = koeff * convertJointToHour(substring)- clockOrientation;
            }
            else if (featuresOrientation.Contains("/"))
            {
                int indexOfChar = featuresOrientation.IndexOf("/");
                substring = featuresOrientation.Substring(0, indexOfChar);
                result = koeff * convertJointToHour(substring)-clockOrientation;
            }
            else if (featuresOrientation.Contains(":")& featuresOrientation.Contains(":")==false & featuresOrientation.Contains("/") == false)
            {
                result = koeff * convertJointToHour(featuresOrientation)- clockOrientation;
            }
            else
            {
                result = 0;
            }
            return result;
        }
        public double GetStartAngle(string featuresOrientation)
        {
            double result = 0;
            featuresOrientation = featuresOrientation.Replace(" ","");
            if (featuresOrientation.Contains("-"))
            {
                result = convertJointToHour(featuresOrientation.Substring(0, featuresOrientation.IndexOf("-")));
            }
            else if (featuresOrientation.Contains("/"))
            {
                result = convertJointToHour(featuresOrientation.Substring(0, featuresOrientation.IndexOf("/")));
            }
            else
            {
                result = convertJointToHour(featuresOrientation);
            }
            return result;
        }
        public string GetExtOrInt(string extOrInt)
        {
            string result = "";
            DefectsTypes types = new DefectsTypes();
            types.defect_location = "";
            if (String.Equals(extOrInt, "EXT"))
            {
                types.defect_location = "Наружный";
            }
            if (String.Equals(extOrInt, "INT"))
            {
                types.defect_location = "Внутренний";
            }
            return result;
        }
        public MGVTD SetPipeLengthToAnomalyLog(MGVTD input)//расставим в журнале аномалий длины труб
        {
            
            for (int i = 0; i < mGVTD.anomalyLogLineS.Count; i++)
            {
                for (int j = 0; j < mGVTD.MGPipeS.Count; j++)
                {
                    if (String.Equals(input.anomalyLogLineS[i].pipeNumber, input.MGPipeS[j].pipeNumber))
                    {
                        input.anomalyLogLineS[i].pipeLength = input.MGPipeS[j].pipeLength;
                    }
                }

            }
                       
            return input;

        }

        public string GetDefectLocation(MGVTD input, int defectNumber)
        {
            string result = "";
            double distanceFromJoint = Math.Min(input.anomalyLogLineS[defectNumber].distanceFromLongitudinalWeld, input.anomalyLogLineS[defectNumber].pipeLength- input.anomalyLogLineS[defectNumber].distanceFromLongitudinalWeld);
            if (distanceFromJoint==0)
            {
                result = "Сварной шов";
            }
            else if (distanceFromJoint <150)
            {
                result = "Околошовная зона";
            }
            else
            {
                result = "Основной металл";
            }
            return result;
        }
        public MGVTD SetVolumesToAnomalyLogForIUST(MGVTD input)
        {
            MGVTD result = input;
            for (int i = 0; i < input.anomalyLogLineS.Count; i++)//расставляем в журнале аномалий ориентацию сварного шва
            {
                for (int j = 0; j < input.MGPipeS.Count; j++)
                {
                    if (String.Equals(input.anomalyLogLineS[i].pipeNumber, input.MGPipeS[j].pipeNumber))
                    {
                        input.anomalyLogLineS[i].clockOrientation = GetStartAngle(input.MGPipeS[j].clockOrientation);
                        //MessageBox.Show("GetStartAngle");
                        input.anomalyLogLineS[i].distanceFromLongitudinalWeld = GetdistanceFromLongitudinalWeld(input.anomalyLogLineS[i].featuresOrientation, input.anomalyLogLineS[i].clockOrientation, input.pipelineInfo.pipeDiameter);
                        //MessageBox.Show("GetdistanceFromLongitudinalWeld");
                        input.anomalyLogLineS[i].defectType = GetDefectTypeGPAS(input.anomalyLogLineS[i].featuresCharacter).defectType;
                        //MessageBox.Show("GetDefectTypeGPAS");
                        input.anomalyLogLineS[i].defectCode = GetDefectTypeGPAS(input.anomalyLogLineS[i].featuresCharacter).defectCode;
                        //MessageBox.Show("GetDefectTypeGPAS");
                        input.anomalyLogLineS[i].start_angle = GetStartAngle(input.anomalyLogLineS[i].featuresOrientation);
                        //MessageBox.Show("GetStartAngle");
                        input.anomalyLogLineS[i].inside_or_outside = GetExtOrInt(input.anomalyLogLineS[i].extOrInt);
                        //MessageBox.Show("GetExtOrInt");
                        input.anomalyLogLineS[i].defect_location = GetDefectLocation(input, i);
                        //MessageBox.Show("GetDefectLocation");
                    }
                }
            }



            return result;
        }

    }
        public class DefectsTypes
        {
            public int defectNumber;//номер дефекта по порядку
            public string pipeNumber;//номер трубы
            public string defectType;//тип дефекта
            public string defectCode;//код дефекта
            public double distanceFromTransverseWeld;//расстояние от первого поперечного шва
            public double distanceFromLongitudinalWeld;//расстояние от продольного шва
            public double start_angle;//начальный угол дефекта
            public double length;//длина
            public double widht;//ширина
            public double depthInMm;//глубина дефекта в миллиметрах
            public string inside_or_outside;//внутренний, наружный, внутристенный
            public string defect_location;//расположение дефекта (основной металл, сварной шов, околошовная зона)
            public string danger_level;//уровень опасности (закритический, критический, допустимый)
        }    

}
