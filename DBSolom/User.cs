using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class User
    {
        public int Id { get; set; }

        public bool Видалено { get; set; } = false;

        public bool New { get; set; } = true;

        public string Контакти { get; set; }

        public string Логін { get; set; }
        
        public string Пароль { get; set; }
    }
}
