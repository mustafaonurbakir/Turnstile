using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace deneme
{
    public class Person
    {
        public string Name;
        public string Departman;
        public DateTime Date;
        public string lastmove;

        public Dictionary<DateTime, double> in_time_day = new Dictionary<DateTime, double>();
        public Dictionary<DateTime,double> in_time_mounth = new Dictionary<DateTime, double>();

        public Dictionary<string, int> girissayisi = new Dictionary<string, int>();
        public Dictionary<string, int> cikissayisi = new Dictionary<string, int>();
        public Dictionary<string, int> gecersiz_giris = new Dictionary<string, int>();
        public Dictionary<string, int> gecersiz_cikis = new Dictionary<string, int>();

        public int invalid_entry;
        public int invalid_out;
    }
}
