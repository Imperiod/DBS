using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class Low
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
        public User Правовласник { get; set; }

        #region "Tables"

        #region "Dictionary"

        public bool DocStatus { get; set; } = false;

        public bool Macrofoundation { get; set; } = false;

        public bool Foundation { get; set; } = false;

        public bool Microfoundation { get; set; } = false;

        public bool KFB { get; set; } = false;

        public bool KDB { get; set; } = false;

        public bool KEKB { get; set; } = false;

        public bool KFK { get; set; } = false;

        public bool Main_manager { get; set; } = false;

        public bool Manager { get; set; } = false;

        #endregion

        #region "Docs"

        public bool Correction { get; set; } = false;

        public bool Filling { get; set; } = false;

        public bool Microfilling { get; set; } = false;

        public bool Financing { get; set; } = false;

        #endregion

        public bool User { get; set; } = false;

        public bool Lowt { get; set; } = false;
        #endregion
    }
}
