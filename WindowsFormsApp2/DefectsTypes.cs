using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VTDinfo
{
    public class DefectsTypes
    {
        public string defectType;
        public string defectCode;
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
    
    }
    
}
