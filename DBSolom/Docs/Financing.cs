using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class Financing
    {
        public Financing()
        {
            Видалено = false;
            Створино = DateTime.Now;
            Проведено = DateTime.Now;
        }

        public long Id { get; set; }

        public bool Видалено { get; set; }

        public User Створив { get; set; }

        public DateTime Створино { get; set; }

        public User Змінив { get; set; }

        public DateTime Змінено { get; set; }

        public DateTime Проведено { get; set; }

        public Main_manager Головний_розпорядник { get; set; }

        public KFK КФК { get; set; }

        public MicroFoundation Мікрофонд { get; set; }

        public KDB КДБ { get; set; }

        public KEKB КЕКВ { get; set; }

        public double Сума { get; set; }

        public bool Підписано { get; set; }
    }
}
