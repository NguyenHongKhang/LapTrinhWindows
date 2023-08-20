using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.BLL
{
    class Fuzzy
    {
        private static Fuzzy _Instance;
        public static Fuzzy Instance
        {
            get
            {

                if (_Instance == null)
                {
                    _Instance = new Fuzzy();
                }
                return _Instance;
            }
            private set
            {

            }
        }

        private Fuzzy()
        {

        }

        private int Levenshtein_Distance(String pat, String txt)
        {
            int l1 = pat.Length;
            int l2 = txt.Length;
            int[,] F = new int[l1 + 1, l2 + 1];
            int i, j;
            for (i = 0; i <= l1; i++) F[i, 0] = i;
            for (i = 0; i <= l2; i++) F[0, i] = i;
            for (i = 1; i <= l1; i++)
            {
                for (j = 1; j <= l2; j++)
                {
                    if (pat[i - 1] == txt[j - 1]) F[i, j] = F[i - 1, j - 1];
                    else F[i, j] = Math.Min(Math.Min(F[i, j - 1], F[i - 1, j]), F[i - 1, j - 1]) + 1;
                }
            }
            return F[l1, l2];
        }

        private string ConvertToUnSign2(string s)
        {
            string stFormD = s.Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder();
            for (int ich = 0; ich < stFormD.Length; ich++)
            {
                System.Globalization.UnicodeCategory uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(stFormD[ich]);
                if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(stFormD[ich]);
                }
            }
            sb = sb.Replace('Đ', 'D');
            sb = sb.Replace('đ', 'd');
            return (sb.ToString().Normalize(NormalizationForm.FormD));
        }

        public String Preprocessing(String str)
        {
            str = ConvertToUnSign2(str);
            String res = "";
            String[] words = str.Split(' ');
            foreach (String s in words)
            {
                if (!s.Equals(""))
                {
                    res += s + " ";
                }
            }
            res = res.ToLower();
            res = res.Trim();
            return res;
        }

        public double Normal_Ratio(String pat, String txt)
        {
            int l1 = pat.Length;
            int l2 = txt.Length;
            if (l1 == 0 && l2 == 0) return 1;
            double res;
            res = (l1 + l2) - Levenshtein_Distance(pat, txt);
            res = res / ((double)(l1 + l2));
            return res;
        }

        public double Partial_Ratio(String pat, String txt)
        {
            int l1 = pat.Length;
            int l2 = txt.Length;
            if (l2 > l1)
            {
                return 0;
            }

            String temp;
            double Ra = 0;
            int i;
            for (i = 0; i < (l1 - l2 + 1); i++)
            {
                temp = pat.Substring(i, l2);
                Ra = Math.Max(Ra, Normal_Ratio(temp, txt));
            }
            return Ra;
        }

        public double Token_Sort_Ratio(String pat, String txt)
        {
            String[] pats = pat.Split(' ');
            List<String> lspats = new List<String>();
            String[] txts = txt.Split(' ');
            List<String> lstxts = new List<String>();
            String newpat = "";
            String newtxt = "";

            foreach (String s in pats)
            {
                lspats.Add(s);
            }
            lspats.Sort();

            foreach (String s in lspats)
            {
                newpat += s + " ";
            }

            foreach (String s in txts)
            {
                lstxts.Add(s);
            }
            lstxts.Sort();

            foreach (String s in lstxts)
            {
                newtxt += s + " ";
            }

            if (newpat.Length > 0) newpat = newpat.Remove(newpat.Length - 1);
            if (newtxt.Length > 0) newtxt = newtxt.Remove(newtxt.Length - 1);

            //Console.WriteLine(newpat +"1" + "\n" + newtxt +"1");
            return Normal_Ratio(newpat, newtxt);
        }

        public double Token_Set_Ratio(String pat, String txt)
        {
            String[] pats = pat.Split(' ');
            List<String> lspats = new List<String>();
            String[] txts = txt.Split(' ');
            List<String> lstxts = new List<String>();

            foreach (String s in pats)
            {
                lspats.Add(s);
            }
            lspats.Sort();

            foreach (String s in txts)
            {
                lstxts.Add(s);
            }
            lstxts.Sort();

            int l1 = lspats.Count;
            int l2 = lstxts.Count;
            int i, j;
            i = j = 0;
            String intersection, restofpat, restoftxt;
            intersection = restofpat = restoftxt = "";

            while (i < l1 && j < l2)
            {
                if (String.Compare(lspats[i], lstxts[j]) == 0)
                {
                    intersection += lspats[i] + " ";
                    i++;
                    j++;
                }
                else if (String.Compare(lspats[i], lstxts[j]) < 0)
                {
                    restofpat += lspats[i] + " ";
                    i++;
                }
                else
                {
                    restoftxt += lstxts[j] + " ";
                    j++;
                }
            }

            for (; i < l1; i++) restofpat += lspats[i] + " ";
            for (; j < l2; j++) restoftxt += lstxts[j] + " ";
            String str1 = intersection + restofpat;
            String str2 = intersection + restoftxt;

            //Console.WriteLine(intersection.Length + " " + str1.Length + " " + str2.Length+ " ");

            if (intersection.Length > 0) intersection = intersection.Remove(intersection.Length - 1);
            if (str1.Length > 0) str1 = str1.Remove(str1.Length - 1);
            if (str2.Length > 0) str2 = str2.Remove(str2.Length - 1);

            //Console.WriteLine(intersection.Length + " " + str1.Length + " " + str2.Length+ " ");
            //Console.WriteLine(intersection + "\n" + str1 + "\n" + str2 );

            double Ra = 0;
            Ra = Math.Max(Normal_Ratio(intersection, str1), Ra);
            Ra = Math.Max(Normal_Ratio(intersection, str2), Ra);
            Ra = Math.Max(Normal_Ratio(str1, str2), Ra);

            return Ra;
        }

        public double Ratio(string pat, string txt)
        {
            pat = Preprocessing(pat);
            txt = Preprocessing(txt);
            double Ra = 0;
            Ra = Math.Max(Ra, Normal_Ratio(pat, txt));
            Ra = Math.Max(Ra, Partial_Ratio(pat, txt));
            Ra = Math.Max(Ra, Token_Sort_Ratio(pat, txt));
            Ra = Math.Max(Ra, Token_Set_Ratio(pat, txt));

            //Console.WriteLine("Normal_Ratio: {0}", Normal_Ratio(pat, txt));
            //Console.WriteLine("Partial_Ratio: {0}", Partial_Ratio(pat, txt));
            //Console.WriteLine("Token_Sort_Ratio: {0}",Token_Sort_Ratio(pat, txt));
            //Console.WriteLine("Token_Set_Ratio: {0}",Token_Set_Ratio(pat, txt));

            return Ra;
        }

    }
}
