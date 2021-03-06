﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class Manager
    {
        [Key]
        [Editable(false)]
        public long Id { get; set; }

        public bool Видалено { get; set; } = false;

        public User Створив { get; set; }

        public DateTime Створино { get; set; } = DateTime.Now;

        public User Змінив { get; set; }

        public DateTime Змінено { get; set; } = DateTime.Now;

        public Main_manager Головний_розпорядник { get; set; }

        public KFK КФК { get; set; }

        public string Найменування { get; set; }

        public string Повністю { get; set; }

        public int ЕГРПОУ { get; set; }

        public byte Рівень_розпорядника { get; set; }
    }
}
