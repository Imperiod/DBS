﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class Main_manager
    {
        public Main_manager()
        {
            Видалено = false;
            Створино = DateTime.Now;
            Змінено = DateTime.Now;
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
        public string Найменування { get; set; }

        public string Повністю { get; set; }

        public int КПОЛ { get; set; }

        public int Код_ГУДКСУ { get; set; }

        public int Код_УДКСУ { get; set; }
    }
}
