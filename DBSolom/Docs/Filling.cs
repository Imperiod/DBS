using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class Filling
    {
        public Filling()
        {
            Видалено = false;
            Створино = DateTime.Now;
            Змінено = DateTime.Now;
            Проведено = DateTime.Now;
            Підписано = false;
        }

        [Key]
        [Editable(false)]
        public long Id { get; set; }

        public bool Видалено { get; set; }

        [Required(AllowEmptyStrings = false)]
        public User Створив { get; set; }

        [Required(AllowEmptyStrings = false)]
        public DateTime Створино { get; set; }

        [Required(AllowEmptyStrings = false)]
        public User Змінив { get; set; }

        [Required(AllowEmptyStrings = false)]
        public DateTime Змінено { get; set; }

        [Required(AllowEmptyStrings = false)]
        public DateTime Проведено { get; set; }

        public bool Підписано { get; set; }

        [Required(AllowEmptyStrings = false)]
        public Main_manager Головний_розпорядник { get; set; }

        [Required(AllowEmptyStrings = false)]
        public KFK КФК { get; set; }

        [Required(AllowEmptyStrings = false)]
        public Foundation Фонд { get; set; }

        [Required(AllowEmptyStrings = false)]
        public KFB КФБ { get; set; }

        [Required(AllowEmptyStrings = false)]
        public KDB КДБ { get; set; }

        [Required(AllowEmptyStrings = false)]
        public KEKB КЕКВ {get; set;}

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
