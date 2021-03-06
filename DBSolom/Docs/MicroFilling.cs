﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class MicroFilling
    {
        [Key]
        [Editable(false)]
        public long Id { get; set; }

        public bool Видалено { get; set; } = false;

        [Required(AllowEmptyStrings = false)]
        public User Створив { get; set; }

        [Required(AllowEmptyStrings = false)]
        public DateTime Створино { get; set; } = DateTime.Now;

        [Required(AllowEmptyStrings = false)]
        public User Змінив { get; set; }

        [Required(AllowEmptyStrings = false)]
        public DateTime Змінено { get; set; } = DateTime.Now;

        [Required(AllowEmptyStrings = false)]
        public DateTime Проведено { get; set; } = DateTime.Now;

        public bool Підписано { get; set; } = false;

        [Required(AllowEmptyStrings = false)]
        public Main_manager Головний_розпорядник { get; set; }

        [Required(AllowEmptyStrings = false)]
        public KFK КФК { get; set; }

        [Required(AllowEmptyStrings = false)]
        public MicroFoundation Мікрофонд { get; set; }

        [Required(AllowEmptyStrings = false)]
        public KFB КФБ { get; set; }

        [Required(AllowEmptyStrings = false)]
        public KDB КДБ { get; set; }

        [Required(AllowEmptyStrings = false)]
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
