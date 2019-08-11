using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class MicroFoundation
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
        public Foundation Фонд { get; set; }

        [Required(AllowEmptyStrings = false)]
        public string Повністю { get; set; }
    }
}
