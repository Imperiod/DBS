using System;
using System.Collections.Generic;
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
        }

        public long Id { get; set; }

        public bool Видалено { get; set; }

        public User Створив { get; set; }

        public DateTime Створино { get; set; }

        public User Змінив { get; set; }

        public DateTime Змінено { get; set; }

        public string Найменування { get; set; }

        public string Повністю { get; set; }

        public int КПОЛ { get; set; }

        public int Код_ГУДКСУ { get; set; }

        public int Код_УДКСУ { get; set; }
    }
}
