using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VTDinfo
{
    internal class JointAngle
    {
        public double angle1 { get; set; }
        public double angle2 { get; set; }
        bool isOneJoint;
        public double convertJointToHour(string inputAngle)//конвертация из формата чч:мм в десятичные часы.
        {
            double result = 0;
            int indexOfChar = inputAngle.IndexOf(":");
            double hours= Convert.ToDouble(inputAngle.Substring(0, indexOfChar).Trim());
            double minuts = Convert.ToDouble(inputAngle.Substring(indexOfChar+1).Trim());
            result = hours + 1.66 * minuts;
            return result;
        }
        public JointAngle GetJointAngle(string angle)//конвертация из "2:14 / 8:14" в экземпляр класса JointAngle
        {
            angle = angle.Replace(" ", "");
            JointAngle jointAngles= new JointAngle();
            jointAngles.angle1 = 0;
            jointAngles.angle1 = 6;
            isOneJoint = false;

            if (angle.Contains("/"))
            {
                int indexOfChar = angle.IndexOf("/");
                string ang1=angle.Substring(0, indexOfChar);    
                string ang2=angle.Substring(indexOfChar+1);
                jointAngles.angle1 = convertJointToHour(ang1);
                jointAngles.angle2 = convertJointToHour(ang2);
                isOneJoint = false;
            }
            else
            {
                jointAngles.angle1 = convertJointToHour(angle.Trim());
                jointAngles.angle2 = 0;
                isOneJoint = true;
            }
            return jointAngles;
        }
    }
}
