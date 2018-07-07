using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class Foundation
    {
        public Foundation()
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

        public MacroFoundation Макрофонд { get; set; } 

        public int Код { get; set; }

        public string Повністю { get; set; }
    }
}
