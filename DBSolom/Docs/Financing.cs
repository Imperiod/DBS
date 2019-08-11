using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class Financing
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

        [Required(AllowEmptyStrings = false)]
        public Main_manager Головний_розпорядник { get; set; }

        [Required(AllowEmptyStrings = false)]
        public KFK КФК { get; set; }

        [Required(AllowEmptyStrings = false)]
        public MicroFoundation Мікрофонд { get; set; }

        [Required(AllowEmptyStrings = false)]
        public KEKB КЕКВ { get; set; }

        public double Сума { get; set; }

        public bool Підписано { get; set; }
    }
}
