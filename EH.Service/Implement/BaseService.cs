using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.Implement;
using EH.Repository.Interface;
using EH.Service.Interface;
using NPOI.POIFS.FileSystem;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using EH.Repository.Interface.Sys;
using EH.System.Models.Dtos;
using EH.Repository.Implement.Sys;
using System.Reflection;

namespace EH.Service.Implement
{
    public class BaseService<T> : IBaseService<T> where T : class
    {
        private readonly JwtHelper jwtHelper;
        private readonly LogHelper logHelper;
        private readonly IRepositoryBase<T> repositoryBase;

        public BaseService()
        {

        }
        public BaseService(IRepositoryBase<T> repositoryBase,LogHelper logHelper)
        {
            this.repositoryBase = repositoryBase;
            this.logHelper = logHelper;
        }


        public string GenerateToken(Sys_Users user)
        {
            return jwtHelper.GenerateToken(user);
        }

        public bool ValidateToken(string token)
        {
            return jwtHelper.ValidateToken(token);
        }


        public TT GetClaimValue<TT>(string token, string type)
        {
            return jwtHelper.GetClaimValue<TT>(token, type);
        }

        public virtual T Insert(T entity,bool isSave=true)
        {
            try
            {
                var res = repositoryBase.Add(entity, isSave);
                return res;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public virtual bool Delete(T entity)
        {
            try
            {
                repositoryBase.Delete(entity);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public virtual bool DeleteRange(List<string> ids)
        {
           
            try
            {
                List<T> list = new List<T>();
                foreach (var id in ids)
                {
                    T entity = repositoryBase.GetById(Guid.Parse(id));
                    if (entity != null)
                    {
                        PropertyInfo propertyInfo = entity.GetType().GetProperty("IsDeleted");
                        if (propertyInfo != null && propertyInfo.PropertyType == typeof(bool))
                        {
                            propertyInfo.SetValue(entity, true); // 修改 IsDeleted 属性为 true，实现软删除
                            list.Add(entity);
                        }

                    }
                }

                repositoryBase.UpdateRange(list);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public virtual bool RealDeleteRange(List<string> ids)
        {

            try
            {
                List<T> list = new List<T>();
                foreach (var id in ids)
                {
                    T entity = repositoryBase.GetById<Guid>(Guid.Parse(id));
                    if (entity != null)
                        list.Add(entity);
                }

                repositoryBase.DeleteRange(list);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public virtual bool Update(T entity)
        {
            try
            {
                repositoryBase.Update(entity);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public virtual List<T> GetPageList(PageRequest<T> request, out int totalCount)
        {
            var whereCondision = request.GetWhere().Compile();
            var orderCondision = request.GetOrder().Compile();

            return repositoryBase.Where(whereCondision, orderCondision, request.PageIndex, request.PageSize, out totalCount, isDesc: request.isDesc).ToList().ToObject<List<T>>();
        }

        public virtual List<T> GetDeleteList(List<string> ids)
        {
            List<T> list = new List<T>();
            foreach (var id in ids)
            {
                T entity = repositoryBase.GetById<Guid>(Guid.Parse(id));
                if (entity != null)
                    list.Add(entity);
            }
            return list;
        }

        public virtual T GetEntityById(Guid id)
        {
            try
            {
                var res = repositoryBase.GetById(id);
                return res;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

    }
}
