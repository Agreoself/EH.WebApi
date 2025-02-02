﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using EH.Repository.DataAccess;
using EH.Repository.Interface;
using EH.System.Models.Common;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;

namespace EH.Repository.Implement
{
    public class RepositoryBase<T> : IRepositoryBase<T> where T : class
    {
        private readonly MyAppDbContext _dbContext;
        private readonly DbSet<T> _dbSet;
        private readonly string _connStr;

        public RepositoryBase(MyAppDbContext mydbcontext)
        {
            this._dbContext = mydbcontext as MyAppDbContext;
            this._dbSet = _dbContext.Set<T>();
            //this._connStr = "";//_dbContext.Database.GetDbConnection().ConnectionString;
        }

        public virtual void BeginTransaction(IsolationLevel isolationLevel = IsolationLevel.Unspecified)
        {
            if (this._dbContext.Database.CurrentTransaction == null)
            {
                this._dbContext.Database.BeginTransaction();
            }
        }

        public virtual void Commit()
        {
            var transaction = this._dbContext.Database.CurrentTransaction;
            if (transaction != null)
            {
                try
                {
                    transaction.Commit();
                }
                catch (Exception)
                {
                    transaction.Rollback();
                    throw;
                }
            }
        }

        public virtual void Rollback()
        {
            if (this._dbContext.Database.CurrentTransaction != null)
            {
                this._dbContext.Database.CurrentTransaction.Rollback();
            }
        }

        public virtual int SaveChanges()
        { 
            return this._dbContext.SaveChanges();
        }

    


        public virtual IQueryable<T> Entities
        {
            get { return this._dbSet.AsNoTracking(); }
        }

        public virtual IQueryable<T> TrackEntities
        {
            get { return this._dbSet; }
        }

        public virtual T Add(T entity, bool isSave = true)
        {
            this._dbSet.Add(entity);
            if (isSave)
            {
                this.SaveChanges();
            }
            return entity;
        }

        public virtual void AddRange(IEnumerable<T> entitys, bool isSave = true)
        {
            this._dbSet.AddRange(entitys);
            if (isSave)
            {
                this.SaveChanges();
            }
        }

        public virtual void Delete(T entity, bool isSave = true)
        {
            this._dbSet.Remove(entity);
            if (isSave)
            {
                this.SaveChanges();
            }
        }

        public virtual void DeleteRange(IEnumerable<T> entitys, bool isSave = true)
        {
            this._dbSet.RemoveRange(entitys);
            if (isSave)
            {
                this.SaveChanges();
            }
        }

        public virtual void Delete(bool isSave = true, params T[] entitys)
        {
            this._dbSet.RemoveRange(entitys);
            if (isSave)
            {
                this.SaveChanges();
            }
        }

        public virtual void Delete(object id, bool isSave = true)
        {
            this._dbSet.Remove(this._dbSet.Find(id));
            if (isSave)
            {
                this.SaveChanges();
            }
        }

        public virtual void Delete(Expression<Func<T, bool>> @where, bool isSave = true)
        {
            T[] entitys = this._dbSet.Where<T>(@where).ToArray();
            if (entitys.Length > 0)
            {
                this._dbSet.RemoveRange(entitys);
            }
            if (isSave)
            {
                this.SaveChanges();
            }
        }

        public void Update(T entity, bool isSave = true)
        {
            var entry = this._dbContext.Entry(entity);
            if (entry.State == EntityState.Detached)
            {
                entry.State = EntityState.Modified;
            }
            if (isSave)
            {
                this.SaveChanges();
            }
        }

        public void Update(bool isSave = true, params T[] entitys)
        {
            var entry = this._dbContext.Entry(entitys);
            if (entry.State == EntityState.Detached)
            {
                entry.State = EntityState.Modified;
            }
            if (isSave)
            {
                this.SaveChanges();
            }
        }

        public void UpdateRange(IEnumerable<T> entitys, bool isSave = true)
        {
            //foreach (var item in entitys)
            //{
            //    var entry = this._dbContext.Entry(item);
            //    if (entry.State == EntityState.Detached)
            //    {
            //        entry.State = EntityState.Modified;
            //    }
            //}
            this._dbSet.UpdateRange(entitys);
            if (isSave)
            {
                this.SaveChanges();
            }
        }

        public bool Any(Expression<Func<T, bool>> @where)
        {
            return this._dbSet.AsNoTracking().Any(@where);
        }

        public int Count()
        {
            return this._dbSet.AsNoTracking().Count();
        }

        public int Count(Expression<Func<T, bool>> @where)
        {
            return this._dbSet.AsNoTracking().Count(@where);
        }

        public T FirstOrDefault(Expression<Func<T, bool>> @where)
        {
            return this._dbSet.AsNoTracking().FirstOrDefault(@where);
        }

        public T FirstOrDefault<TOrder>(Expression<Func<T, bool>> @where, Expression<Func<T, TOrder>> order, bool isDesc = false)
        {
            if (isDesc)
            {
                return this._dbSet.AsNoTracking().OrderByDescending(order).FirstOrDefault(@where);
            }
            else
            {
                return this._dbSet.AsNoTracking().OrderBy(order).FirstOrDefault(@where);
            }
        }

        public IQueryable<T> Distinct(Expression<Func<T, bool>> @where)
        {
            return this._dbSet.AsNoTracking().Where(@where).Distinct();
        }

        public IQueryable<T> Where(Expression<Func<T, bool>> @where)
        {
            return this._dbSet.Where(@where);
        }

        public IQueryable<T> Where<TOrder>(Expression<Func<T, bool>> @where, Expression<Func<T, TOrder>> order, bool isDesc = false)
        {
            if (isDesc)
            {
                return this._dbSet.Where(@where).OrderByDescending(order);
            }
            else
            {
                return this._dbSet.Where(@where).OrderBy(order);
            }
        }

        public IEnumerable<T> Where<TOrder>(Func<T, bool> @where, Func<T, TOrder> order, int pageIndex, int pageSize, out int count, bool isDesc = false)
        {
            count = _dbSet.Where(@where).Count();
            if (isDesc)
            {
                return this._dbSet.Where(@where).OrderByDescending(order).Skip((pageIndex - 1) * pageSize).Take(pageSize);
            }
            else
            {
                return this._dbSet.Where(@where).OrderBy(order).Skip((pageIndex - 1) * pageSize).Take(pageSize);
            }
        }

        public IQueryable<T> Where<TOrder>(Expression<Func<T, bool>> @where, Expression<Func<T, TOrder>> order, int pageIndex, int pageSize, out int count, bool isDesc = false)
        {
            count = Count(@where);
            if (isDesc)
            {
                return this._dbSet.Where(@where).OrderByDescending(order).Skip((pageIndex - 1) * pageSize).Take(pageSize);
            }
            else
            {
                return this._dbSet.Where(@where).OrderBy(order).Skip((pageIndex - 1) * pageSize).Take(pageSize);
            }
        }

        public IQueryable<T> GetAll()
        {
            return this._dbSet.AsNoTracking();
        }

        public IQueryable<T> GetAll<TOrder>(Expression<Func<T, TOrder>> order, bool isDesc = false)
        {
            if (isDesc)
            {
                return this._dbSet.AsNoTracking().OrderByDescending(order);
            }
            else
            {
                return this._dbSet.AsNoTracking().OrderBy(order);
            }
        }

        public T GetById<Ttype>(Ttype id)
        {
            return this._dbSet.Find(id);
        }

        public Ttype Max<Ttype>(Expression<Func<T, Ttype>> column)
        {
            if (this._dbSet.AsNoTracking().Any())
            {
                return this._dbSet.AsNoTracking().Max<T, Ttype>(column);
            }
            return default(Ttype);
        }

        public Ttype Max<Ttype>(Expression<Func<T, Ttype>> column, Expression<Func<T, bool>> @where)
        {
            if (this._dbSet.AsNoTracking().Any(@where))
            {
                return this._dbSet.AsNoTracking().Where(@where).Max<T, Ttype>(column);
            }
            return default(Ttype);
        }

        public Ttype Min<Ttype>(Expression<Func<T, Ttype>> column)
        {
            if (this._dbSet.AsNoTracking().Any())
            {
                return this._dbSet.AsNoTracking().Min<T, Ttype>(column);
            }
            return default(Ttype);
        }

        public Ttype Min<Ttype>(Expression<Func<T, Ttype>> column, Expression<Func<T, bool>> @where)
        {
            if (this._dbSet.AsNoTracking().Any(@where))
            {
                return this._dbSet.AsNoTracking().Where(@where).Min<T, Ttype>(column);
            }
            return default(Ttype);
        }

        public TType Sum<TType>(Expression<Func<T, TType>> selector, Expression<Func<T, bool>> @where) where TType : new()
        {
            object result = 0;

            if (new TType().GetType() == typeof(decimal))
            {
                result = this._dbSet.AsNoTracking().Where(where).Sum(selector as Expression<Func<T, decimal>>);
            }
            if (new TType().GetType() == typeof(decimal?))
            {
                result = this._dbSet.AsNoTracking().Where(where).Sum(selector as Expression<Func<T, decimal?>>);
            }
            if (new TType().GetType() == typeof(double))
            {
                result = this._dbSet.AsNoTracking().Where(where).Sum(selector as Expression<Func<T, double>>);
            }
            if (new TType().GetType() == typeof(double?))
            {
                result = this._dbSet.AsNoTracking().Where(where).Sum(selector as Expression<Func<T, double?>>);
            }
            if (new TType().GetType() == typeof(float))
            {
                result = this._dbSet.AsNoTracking().Where(where).Sum(selector as Expression<Func<T, float>>);
            }
            if (new TType().GetType() == typeof(float?))
            {
                result = this._dbSet.AsNoTracking().Where(where).Sum(selector as Expression<Func<T, float?>>);
            }
            if (new TType().GetType() == typeof(int))
            {
                result = this._dbSet.AsNoTracking().Where(where).Sum(selector as Expression<Func<T, int>>);
            }
            if (new TType().GetType() == typeof(int?))
            {
                result = this._dbSet.AsNoTracking().Where(where).Sum(selector as Expression<Func<T, int?>>);
            }
            if (new TType().GetType() == typeof(long))
            {
                result = this._dbSet.AsNoTracking().Where(where).Sum(selector as Expression<Func<T, long>>);
            }
            if (new TType().GetType() == typeof(long?))
            {
                result = this._dbSet.AsNoTracking().Where(where).Sum(selector as Expression<Func<T, long?>>);
            }
            return (TType)result;
        }

        public void Dispose()
        {
            this._dbContext.Dispose();
        }


    }
}
