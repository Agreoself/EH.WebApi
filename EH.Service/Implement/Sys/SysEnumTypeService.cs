using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using EH.Repository.Implement.Sys;
using EH.Repository.Interface;
using EH.Repository.Interface.Sys;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using EH.Repository.Implement;
using NPOI.SS.Formula.Functions;
using System.Reflection;
using EH.Service.Interface.Sys;

namespace EH.Service.Implement.Sys
{
    public class SysEnumTypeService : BaseService<Sys_EnumType>, ISysEnumTypeService, ITransient
    {
        private readonly LogHelper logHelper;
        private readonly ISysEnumTypesRepository repository;
        private readonly ISysEnumItemsRepository itemRepository;
        public SysEnumTypeService(ISysEnumTypesRepository repository, ISysEnumItemsRepository itemRepository, LogHelper logHelper) : base(repository, logHelper)
        {
            this.repository = repository;
            this.itemRepository = itemRepository;
        }

        public override bool DeleteRange(List<string> ids, bool isSave = false)
        {
            try
            {
                List<Sys_EnumType> list = new List<Sys_EnumType>();
                List<Sys_EnumItem> itemList = new List<Sys_EnumItem>();
                foreach (var id in ids)
                {
                    Sys_EnumType entity = repository.GetById(Guid.Parse(id));
                    if (entity != null)
                    {
                        PropertyInfo propertyInfo = entity.GetType().GetProperty("IsDeleted");
                        if (propertyInfo != null && propertyInfo.PropertyType == typeof(bool))
                        {
                            propertyInfo.SetValue(entity, true); // 修改 IsDeleted 属性为 true，实现软删除
                            list.Add(entity);
                        }
                        itemList.AddRange(itemRepository.Entities.Where(i => i.EnumTypeId == id).ToList());
                        itemList.ForEach(i => i.IsDeleted = true);
                    }
                }
                itemRepository.UpdateRange(itemList);
                repository.UpdateRange(list);
                return true;
            }
            catch (Exception ex)
            {
                logHelper.LogError(ex.ToString());
                return false;
            }

        }

        public Dictionary<string, object> GetAllDic()
        {
            Dictionary<string, object> enums=new Dictionary<string, object>();
            var enumType = repository.Entities.ToList();
            foreach (var e in enumType)
            {
                var eItem = itemRepository.Entities.Where(i => i.EnumTypeId == e.ID.ToString()).ToList().Select(i => new {label=i.Text,value=i.Value });
                enums.Add(e.EnumCode, eItem);
            }
            return enums;
        }

    }
}
