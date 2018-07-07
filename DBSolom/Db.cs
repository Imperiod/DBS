using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.SqlServer;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace DBSolom
{
    class MyConfiguration : DbConfiguration
    {
        public MyConfiguration()
        {
            SetExecutionStrategy(
                "System.Data.SqlClient",
                () => new SqlAzureExecutionStrategy(10, TimeSpan.FromSeconds(10)));
        }
    }

    public class Db : DbContext
    {
        public List<string> list { get; set; }
        public List<string> names_months { get; set; }

    public Db(string conn) : base(conn)
        {
            DbConfiguration.SetConfiguration(new MyConfiguration());
            //Database.SetInitializer<Db>(new DropCreateDatabaseIfModelChanges<Db>());
            list = new List<string> { "==", "!=", ">", ">=", "<", "<=", "-", ">|<", "[,]" };
            names_months = new List<string>() { "Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень", "Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Грудень" };
        }

        //9
        #region "Dictionary"

        public DbSet<MacroFoundation> MacroFoundations { get; set; }

        public DbSet<Foundation> Foundations { get; set; }

        public DbSet<MicroFoundation> MicroFoundations { get; set; }

        public DbSet<KDB> KDBs { get; set; }

        public DbSet<KEKB> KEKBs { get; set; }

        public DbSet<KFK> KFKs { get; set; }

        public DbSet<Main_manager> Main_Managers { get; set; }

        public DbSet<Manager> Managers { get; set; }

        public DbSet<DocStatus> DocStatuses { get; set; }

        #endregion 

        //4
        #region "Docs"

        public DbSet<Filling> Fillings { get; set; }

        public DbSet<MicroFilling> Microfillings { get; set; }

        public DbSet<Financing> Financings { get; set; }

        public DbSet<Correction> Corrections { get; set; }

        #endregion

        public DbSet<Low> Lows { get; set; }

        public DbSet<User> Users { get; set; }
    }
}
