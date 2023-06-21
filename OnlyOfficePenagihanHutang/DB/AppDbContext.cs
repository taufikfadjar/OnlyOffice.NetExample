using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;
using System.Data.Entity.ModelConfiguration.Conventions;

namespace OnlyOfficePenagihanHutang.DB
{

    public class AppDbContext : DbContext
    {
        public AppDbContext() : base("DefaultConnection")
        {
        }

        public DbSet<Tagihan> Tagihan { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Conventions.Remove<PluralizingTableNameConvention>();
            base.OnModelCreating(modelBuilder);
        }

    }


    public class Tagihan
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Alamat { get; set; }
        public string NomerSuratHutang { get; set; }
        public string NomerFaktur { get; set; }
        public int Harga { get; set; }
    }

   
}