﻿using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace EH.Service.Interface
{
    public interface IBaseService<T> where T : class
    {
        public string GenerateToken(Sys_Users user);
        public bool ValidateToken(string token);
        public TT GetClaimValue<TT>(string token, string type);

        public T Insert(T entity,bool isSave = false);
        public bool Update(T entity,bool isSave=false);
        public bool Delete(T entity, bool isSave = false);
        public bool DeleteRange(List<string> ids, bool isSave = false);
        public bool RealDeleteRange(List<string> ids, bool isSave = false);

        public List<T> GetPageList(PageRequest<T> request, out int totalCount);

        public List<T> GetDeleteList(List<string> ids);
        public T GetEntityById(Guid id);

        public bool UpdateRange(List<T> entitys, bool isSave = false);
        public bool InsertRange(List<T> entitys, bool isSave = false);
    }
}
