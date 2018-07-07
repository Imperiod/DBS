using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class User
    {
        public User()
        {
            Видалено = false;
            New = true;
        }

        public int Id { get; set; }

        public bool Видалено { get; set; }

        public bool New { get; set; }

        public string Контакти { get; set; }

        public string Логін { get; set; }
        
        public string Пароль { get; set; }
    }
}
