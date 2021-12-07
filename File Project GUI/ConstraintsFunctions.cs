using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace File_Project_GUI
{
    class ConstraintsFunctions
    {
        public bool notNull(List<List<string>> rec, List<int> idx)
        {
            foreach (var k in idx)
            {
                for (int j = 0; j < rec.Count; j++)
                {
                    if (rec[j][k - 1] == null || rec[j][k - 1] == "") 
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public void removeNulls(List<List<string>> rec, SortedSet<int> rows)
        {
            int index = 0;
            for(int i = 0;i < rows.Count; i++)
            {
                rec.RemoveAt(rows.ElementAt(i) - index);
                index++;
            }
        }

        public List<List<string>> Default(List<List<string>> rec, List<int> idx, string arr)
        {
            string[] val = arr.Split(',');
            for (int i = 0; i < val.Length; i++)
            {
                for (int j = 0; j < rec.Count; j++)
                {
                    if (rec[j][idx[i] - 1] == null || rec[j][idx[i] - 1] == "")
                    {
                        rec[j][idx[i] - 1] = val[i];
                    }
                }
            }
            return rec;
        }

        public bool Unique(List<List<string>> rec, List<int> idx)
        {

            SortedSet<string> set = new SortedSet<string>();
            SortedSet<int> temp = new SortedSet<int>();

            foreach(var i in idx)
            {
                for(int j = 0;j < rec.Count; j++)
                {
                    if(rec[j][i - 1] == null || rec[j][i - 1] == "")
                    {
                        temp.Add(j);
                    }
                }
            }

            removeNulls(rec, temp);

            foreach (var k in idx)
            {
                for (int j = 0; j < rec.Count; j++)
                {
                    set.Add(rec[j][k - 1]);
                }

                if (set.Count != rec.Count)
                {
                    return false;
                }
                set.Clear();
            }

            return true;
        }

        public bool Condition(List<List<string>> rec, List<int> idx, string arr)
        {
            SortedSet<int> tmp = new SortedSet<int>();

            foreach (var i in idx)
            {
                for (int j = 0; j < rec.Count; j++)
                {
                    if (rec[j][i - 1] == null || rec[j][i - 1] == "")
                    {
                        tmp.Add(j);
                    }
                }
            }
            string[] con = arr.Split(',');

            removeNulls(rec, tmp);

            for (int i = 0; i < con.Length; i++)
            {
                int LessThanOrEqual, GreaterThanOrEqual, start = 0, Greater, Lesser, Equal, Xindx;
                string Range1 = "", Range2 = "", temp = "";

                GreaterThanOrEqual = con[i].IndexOf("<=", start);
                Xindx = con[i].IndexOf('X', start);
                if (GreaterThanOrEqual != -1)
                {
                    if (GreaterThanOrEqual < Xindx)
                    {
                        Range1 = con[i].Substring(start, Xindx);
                    }

                    else
                        Range2 = con[i].Substring(GreaterThanOrEqual, con[i].Length - Xindx - 1);
                }
                if (GreaterThanOrEqual < 0) GreaterThanOrEqual = 0;

                GreaterThanOrEqual = con[i].IndexOf("<=", GreaterThanOrEqual + 1);
                if (GreaterThanOrEqual > Xindx && Range2 == "")
                {
                    Range2 = con[i].Substring(GreaterThanOrEqual, con[i].Length - Xindx - 1);
                    GreaterThanOrEqual = 0;
                }
                else GreaterThanOrEqual = -1;

                Greater = con[i].IndexOf("<", start);

                if (GreaterThanOrEqual == -1 && Greater != -1)
                {
                    if (Greater < Xindx)
                    {
                        Range1 = con[i].Substring(start, Xindx);
                    }

                    else if (Greater >= Xindx && Range2 == "")
                        Range2 = con[i].Substring(Greater, con[i].Length - Xindx - 1);
                }
                if (Greater < 0) Greater = 0;
                Greater = con[i].IndexOf("<", Greater + 1);
                if (GreaterThanOrEqual == -1 && Greater > Xindx && Range2 == "")
                {
                    Range2 = con[i].Substring(Greater, con[i].Length - Xindx - 1);
                }

                LessThanOrEqual = con[i].IndexOf(">=", start);
                if (LessThanOrEqual != -1)
                {
                    if (LessThanOrEqual < Xindx && LessThanOrEqual != -1)
                        Range1 = con[i].Substring(start, Xindx);

                    else
                    {
                        Range2 = con[i].Substring(LessThanOrEqual, con[i].Length - Xindx - 1);
                    }
                }
                if (LessThanOrEqual < 0) LessThanOrEqual = 0;
                LessThanOrEqual = con[i].IndexOf(">=", LessThanOrEqual + 1);
                if (LessThanOrEqual > Xindx && Range2 == "")
                {
                    Range2 = con[i].Substring(LessThanOrEqual, con[i].Length - Xindx - 1);
                    LessThanOrEqual = 0;
                }
                else LessThanOrEqual = -1;
                Lesser = con[i].IndexOf(">", start);

                if (LessThanOrEqual == -1 && Lesser != -1)
                {

                    if (Lesser < Xindx)
                        Range1 = con[i].Substring(start, Xindx);

                    else
                        Range2 = con[i].Substring(Lesser, con[i].Length - Xindx - 1);
                }
                if (LessThanOrEqual == -1 && Lesser > Xindx && Range2 == "")
                {
                    Range2 = con[i].Substring(Lesser, con[i].Length - Xindx - 1);
                }
                if (Lesser < 0) Lesser = 0;
                Lesser = con[i].IndexOf(">", Lesser + 1);


                Equal = con[i].IndexOf("==", start);

                if (Equal != -1)
                {

                    if (Equal < Xindx)
                        Range1 = con[i].Substring(start, Xindx);

                    else
                        Range2 = con[i].Substring(Equal, con.Length - Xindx - 1);

                }
                if (Equal < 0) Equal = 0;
                Equal = con[i].IndexOf("==", Equal + 1);

                if (Equal != -1 && Equal > Xindx && Range2 == "")
                {
                    Range2 = con[i].Substring(Equal, con[i].Length - Xindx - 1);
                }

                if (Range1 == "") Range1 = "0";

                if (Range2 == "") Range2 = "0";
                if (Range1 == "0" && Range2 == "0")
                {
                        for (int j = 0; j < rec.Count; j++)
                        {
                            if (rec[j][idx[i]-1] != con[i]) return false;
                        }
                }
                else
                {
                        for (int j = 0; j < rec.Count; j++)
                        {
                            if (Range1 != "0")
                            {
                                if (Range1[Range1.Length - 2] == '<' && Range1[Range1.Length - 1] == '=')
                                {
                                    temp = Range1.Substring(0, Range1.Length - 2);
                                    if (int.Parse(temp) > int.Parse(rec[j][idx[i]-1])) return false;

                                }

                                else if (Range1[Range1.Length - 1] == '<')
                                {
                                    temp = Range1.Substring(0, Range1.Length - 1);
                                    if (int.Parse(temp) >= int.Parse(rec[j][idx[i]-1])) return false;

                                }

                                else if (Range1[Range1.Length - 1] == '>')
                                {
                                    temp = Range1.Substring(0, Range1.Length - 1);
                                    if (int.Parse(temp) <= int.Parse(rec[j][idx[i]-1])) return false;

                                }

                                else if (Range1[Range1.Length - 1] == '=' && Range1[Range1.Length - 2] == '=')
                                {
                                    temp = Range1.Substring(0, Range1.Length - 2);
                                    if (int.Parse(temp) != int.Parse(rec[j][idx[i]-1])) return false;
                                }
                            }

                            if (Range2 != "0")
                            {

                                if (Range2[0] == '>' && Range2[1] == '=')
                                {
                                    temp = Range2.Substring(2, Range2.Length - 2);
                                    if (int.Parse(temp) > int.Parse(rec[j][idx[i]-1])) return false;

                                }

                                else if (Range2[0] == '>' && LessThanOrEqual == -1)
                                {
                                    temp = Range2.Substring(1, Range2.Length - 1);
                                    if (int.Parse(temp) >= int.Parse(rec[j][idx[i]-1])) return false;
                                }

                                if (Range2[0] == '<' && Range2[1] == '=')
                                {
                                    temp = Range2.Substring(2, Range2.Length - 2);
                                    if (int.Parse(temp) < int.Parse(rec[j][idx[i]-1])) return false;
                                }

                                else if (Range2[0] == '<' && GreaterThanOrEqual == -1)
                                {
                                    temp = Range2.Substring(1, Range2.Length - 1);
                                    if (int.Parse(temp) <= int.Parse(rec[j][idx[i]-1])) return false;


                                }

                                else if (Range2[0] == '=' && Range2[1] == '=')
                                {
                                    temp = Range2.Substring(2, Range2.Length - 2);
                                    if (int.Parse(temp) != int.Parse(rec[j][idx[i]-1])) return false;
                                }
                            }
                        }

                }
            }
            return true;
        }

    }

}

