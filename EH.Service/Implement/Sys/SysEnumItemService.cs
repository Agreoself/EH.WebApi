using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.Interface.Sys;
using EH.Service.Interface.Sys;

namespace EH.Service.Implement.Sys
{
    public class SysEnumItemService : BaseService<Sys_EnumItem>, ISysEnumItemService, ITransient
    {
        private readonly LogHelper logHelper; 
        private readonly ISysEnumItemsRepository repository; 
        public SysEnumItemService(ISysEnumItemsRepository repository, LogHelper logHelper) : base(repository, logHelper)
        {
            this.repository = repository; 
        }

    }
}
