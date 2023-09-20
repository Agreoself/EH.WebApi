using EH.System.Models.Common;
using EH.System.Models.Entities;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using System.Collections.Generic;
using System.Net.Http;
using System.Reflection.Emit;

namespace EH.Repository.DataAccess
{
    public class MyAppDbContext : DbContext
    {
        public MyAppDbContext(DbContextOptions opt) : base(opt)
        {
        }

        public DbSet<Sys_Menus> Sys_Menus { get; set; }
        public DbSet<Sys_Users> Sys_Users { get; set; }
        public DbSet<Sys_Roles> Sys_Roles { get; set; }
        public DbSet<Sys_RoleMenu> Sys_RoleMenu { get; set; }
        public DbSet<Sys_UserRole> Sys_UserRole { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            base.OnConfiguring(optionsBuilder);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            //modelBuilder.Ignore<BaseEntity>();
            base.OnModelCreating(modelBuilder);
            modelBuilder.ApplyConfigurationsFromAssembly(this.GetType().Assembly);

            //modelBuilder.Entity<BaseEntity>().HasQueryFilter(entity => !entity.IsDeleted);//可以使用IgnoreQueryFilters()忽略
        }


        public override int SaveChanges()
        {
            //// 软删除的处理
            //ChangeTracker.DetectChanges();

            //var softDeletedEntities = ChangeTracker.Entries()
            //    .Where(x => x.State == EntityState.Deleted && x.Entity is BaseEntity)
            //    .ToList();

            //foreach (var deletedEntity in softDeletedEntities)
            //{
            //    deletedEntity.State = EntityState.Modified;
            //    deletedEntity.CurrentValues["IsDeleted"] = true; 
            //}

            ////真删除
             

            return base.SaveChanges();
        }

    }
}