using EH.System.Models.Common;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;
using Microsoft.EntityFrameworkCore.Diagnostics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Commons
{
    public class DbContextHelper
    {
       
    }

    public class CreateByInterceptor : SaveChangesInterceptor
    {
        private readonly IHttpContextAccessor _httpContextAccessor;

        public CreateByInterceptor(IHttpContextAccessor httpContextAccessor)
        {
            _httpContextAccessor = httpContextAccessor;
        }

        public override InterceptionResult<int> SavingChanges(DbContextEventData eventData, InterceptionResult<int> result)
        {
            var entities = eventData.Context.ChangeTracker.Entries()
                .Where(e => e.State == EntityState.Added || e.State == EntityState.Modified);

            foreach (var entityEntry in entities)
            {
                var currentLoggedInUser = _httpContextAccessor.HttpContext.User.Identity.Name;
                if (currentLoggedInUser != null)
                {
                    currentLoggedInUser = currentLoggedInUser.Split('\\')[1];
                }
                var entity = (BaseEntity)entityEntry.Entity;

                if (entityEntry.State == EntityState.Added)
                {
                    entity.CreateDate = DateTime.Now;
                    entity.CreateBy = currentLoggedInUser;
                }

                entity.ModifyDate = DateTime.Now;
                entity.ModifyBy = currentLoggedInUser;
            } 
              
            return base.SavingChanges(eventData, result);
        }
    }
}
