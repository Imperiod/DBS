﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    public class Low
    {
        public Low()
        {
            Видалено = false;
            Створино = DateTime.Now;

            //9
            DocStatus = false;
            Macrofoundation = false;
            Foundation = false;
            Microfoundation = false;
            KDB = false;
            KEKB = false;
            KFK = false;
            Main_manager = false;
            Manager = false;

            //4
            Correction = false;
            Filling = false;
            Microfilling = false;
            Financing = false;

            //1
            Lowt = false;
        }

        public long Id { get; set; }

        public bool Видалено { get; set; }

        public User Створив { get; set; }

        public DateTime Створино { get; set; }

        public User Змінив { get; set; }

        public DateTime Змінено { get; set; }

        public User Правовласник { get; set; }

        #region "Tables"

        #region "Dictionary"

        public bool DocStatus { get; set; }

        public bool Macrofoundation { get; set; }

        public bool Foundation { get; set; }

        public bool Microfoundation { get; set; }

        public bool KDB { get; set; }

        public bool KEKB { get; set; }

        public bool KFK { get; set; }

        public bool Main_manager { get; set; }

        public bool Manager { get; set; }

        #endregion

        #region "Docs"

        public bool Correction { get; set; }

        public bool Filling { get; set; }

        public bool Microfilling { get; set; }

        public bool Financing { get; set; }

        #endregion

        public bool User { get; set; }

        public bool Lowt { get; set; }
        #endregion
    }
}
