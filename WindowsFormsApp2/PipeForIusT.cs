using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VTDinfo
{
    internal class PipeForIusT
    {
        public int defectNumber;//номер дефекта по порядку
        public string pipeNumber;//номер трубы
        public string defectType;//тип дефекта
        public string defectCode;//код дефекта
        public double distance_from_transverse_weld1;//расстояние от первого поперечного шва
        public double distance_from_longitudinal_weld;//расстояние от продольного шва
        public double start_angle;//начальный угол дефекта
        public double length;//длина
        public double widht;//ширина
        public double depthInMm;//глубина дефекта в миллиметрах
        public string inside_or_outside;//внутренний, наружный, внутристенный
        public string defect_location;//расположение дефекта (основной металл, сварной шов, околошовная зона)
        public string danger_level;//уровень опасности (закритический, критический, допустимый)


        public double odometrDist_mm;//дистанция по одометру в мм
        public string geographical_coordinates;//географические координаты
        public double pipeLength;//длина трубы
        public string characterFeatures;// характер особенности
        public double pipeDiameter;//диаметр трубы
        public double thikness;//толщина трубы





        public double distance_from_transverse_weld2;//расстояние от второго поперечного шва



    }
}
