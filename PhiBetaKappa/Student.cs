using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PhiBettaKappa
{
    public class Student
    {
        private static int[] areaIndex = { 0, 7, 10, 13, 14, 20, 21, 22, 32, 35, 37, 55, 66, 68, 77, 79, 88, 89, 104, 149, 160, 199, 202, 237, 248 }; // 0 - 23 (extra)

        public String ID, lastName, prefName;
        public List<String> college;
        // credits = EARNED + PROGRESS, earned = transfer + institution
        public int transfer, institution, earned, progress, credits, hours;
        public float GPA;
        public List<String> courseSubj;
        public List<String> courseNums;
        public List<String> titles; // course titles
        public List<String> creds;
        public List<String> grades;

        public Student()
        {
            ID = lastName = prefName = "";
            transfer = institution = earned = progress = credits = hours = 0;
            GPA = 0.0f;
            college = new List<String>();
            courseSubj = new List<String>();
            courseNums = new List<String>();
            titles = new List<String>();
        }

        public String collegeList()
        {
            if (college.Count == 0) return "";
            String output = "";
            for (int i = 0; i < college.Count - 1; i++) output += college[i] + ", ";
            output += college[college.Count - 1];
            return output;
        }

        // Get the course numbers taken for a specific subject
        public String getNumbers(String subject)
        {
            String output = "";

            for(int i = 0; i < courseSubj.Count; i++){
                if (courseSubj[i].Equals(subject) && !creds[i].Equals("0.00")) output += Form1.extractNumber(courseNums[i], false) + ",";
            }

            if (output.Length == 0) return "NONE";

            return output.TrimEnd(',');
        }

        public List<String> coursesToString(int area)
        {
            List<String> output = new List<String>();

            for (int i = 0; i < courseSubj.Count; i++)
            {
                for (int j = areaIndex[area]; j < areaIndex[area + 1]; j++)
                {
                    if (Form1.subjects[j].Equals(courseSubj[i]))  output.Add(courseToString(i));
                }
            }

            return output;
        }

        public String courseToString(int index)
        {
            String output = courseSubj[index] + " " + courseNums[index] + " ";
            while (TextRenderer.MeasureText(output, new Font("Times New Roman", 20.0f)).Width < 160) output += " ";
            output += titles[index];
            while (TextRenderer.MeasureText(output, new Font("Times New Roman", 20.0f)).Width < 540) output += " ";
            output += creds[index] + " " + grades[index];
            return output;
        }

        public static String collegeName(String collegeRead)
        {
            if (collegeRead.Contains("Ruben")) return "RSENR";
            if (collegeRead.Contains("Life")) return "CALS";
            if (collegeRead.Contains("Arts")) return "CAS";
            if (collegeRead.Contains("Engineer")) return "CEMS";
            if (collegeRead.Contains("Education")) return "CESS";
            if (collegeRead.Contains("Nursing")) return "CNHS";
            if (collegeRead.Contains("Business")) return "SBA";
            if (collegeRead.Contains("Graduate")) return "GC";
            return collegeRead;
        }
    }
}
