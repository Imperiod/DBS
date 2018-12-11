using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class Correction
    {
        public Correction()
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

        [MaxLength(5)]
        public string Внутрішній_номер { get; set; }

        public string Підстава { get; set; }

        public DocStatus Статус { get; set; }

        public Main_manager Головний_розпорядник { get; set; }

        public KFK КФК { get; set; }

        public MicroFoundation Мікрофонд { get; set; }

        public KDB КДБ { get; set; }

        public KEKB КЕКВ { get; set; }

        public double Січень { get; set; }

        public double Лютий { get; set; }

        public double Березень { get; set; }

        public double Квітень { get; set; }

        public double Травень { get; set; }

        public double Червень { get; set; }

        public double Липень { get; set; }

        public double Серпень { get; set; }

        public double Вересень { get; set; }

        public double Жовтень { get; set; }

        public double Листопад { get; set; }

        public double Грудень { get; set; }
    }
}
