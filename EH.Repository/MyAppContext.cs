using EH.System.Models.Common;
using EH.System.Models.Entities;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Diagnostics;
using System.Collections.Generic;
using System.Configuration;
using System.Linq.Expressions;
using System.Net.Http;
using System.Reflection;
using System.Reflection.Emit;

namespace EH.Repository.DataAccess
{
    public class MyAppDbContext : DbContext
    {
        public MyAppDbContext(DbContextOptions opt) : base(opt)
        {
        }
        #region System
        //public DbSet<Sys_Menus> Sys_Menus { get; set; }
        //public DbSet<Sys_Users> Sys_Users { get; set; }
        //public DbSet<Sys_Roles> Sys_Roles { get; set; }
        //public DbSet<Sys_RoleMenu> Sys_RoleMenu { get; set; }
        //public DbSet<Sys_UserRole> Sys_UserRole { get; set; }
        //public DbSet<Sys_EnumType> Sys_EnumType { get; set; }
        //public DbSet<Sys_EnumItem> Sys_EnumItem { get; set; }
        #endregion


        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            base.OnConfiguring(optionsBuilder);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            Assembly assembly = typeof(BaseEntity).Assembly;
            // 获取所有继承自 EntityBase 的非 abstract 类
            List<Type> entityTypes = assembly.GetTypes()
                .Where(t => t.IsSubclassOf(typeof(BaseEntity)) && !t.IsAbstract)
                .ToList();

            //// 注册实体
            //foreach (Type entityType in entityTypes)
            //{
            //    modelBuilder.Entity(entityType);
                
            //}

            //var s = modelBuilder.Model.GetEntityTypes();
            //var entityTypes= modelBuilder.Model.GetEntityTypes().Where(e => typeof(BaseEntity).IsAssignableFrom(e.ClrType));

            foreach (var entityType in entityTypes)
            {
                modelBuilder.Entity(entityType).Property<bool>(nameof(BaseEntity.IsDeleted));  //也可以直接填写 "IsDeleted"

                var parameter = Expression.Parameter(entityType, "e"); 
                var body = Expression.Equal(Expression.Call(typeof(EF), nameof(EF.Property), new[] { typeof(bool) }, parameter, Expression.Constant(nameof(BaseEntity.IsDeleted))), Expression.Constant(false));

                modelBuilder.Entity(entityType).HasQueryFilter(Expression.Lambda(body, parameter));
            }



            //var fixDecimalDatas = new List<Tuple<Type, Type, string>>();
            //foreach (var entityType in modelBuilder.Model.GetEntityTypes())
            //{
            //    foreach (var property in entityType.GetProperties())
            //    {
            //        if (Type.GetTypeCode(property.ClrType) == TypeCode.Decimal)
            //        {
            //            fixDecimalDatas.Add(new Tuple<Type, Type, string>(entityType.ClrType, property.ClrType, property.Name));
            //        }
            //    }
            //}

            //foreach (var item in fixDecimalDatas)
            //{
            //    modelBuilder.Entity(item.Item1).Property(item.Item2, item.Item3).HasPrecision(10, 2);
            //} 

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