using EH.Repository.Interface.Attendance;
using EH.Repository.Interface.Sys;
using EH.Service.Implement;
using EH.Service.Interface.Attendance;
using EH.System.Commons;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using StackExchange.Redis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Service.Interface.Sys
{
    public class AucActivityService : BaseService<Auc_Activity>, IAucActivityService, ITransient
    {
        private readonly IAucActivityRepository repository;
        private readonly LogHelper logHelper;
        private readonly IConnectionMultiplexer redis;
        private IDatabase Database => redis.GetDatabase();
        public AucActivityService(IAucActivityRepository repository, IConnectionMultiplexer redis, LogHelper logHelper) : base(repository, logHelper)
        {
            this.repository = repository;
            this.logHelper = logHelper;
            this.redis = redis;
        }

        public override Auc_Activity Insert(Auc_Activity entity, bool isSave = false)
        {
            try
            {
                repository.BeginTransaction();
                entity.SetLifecycle();
                var res = base.Insert(entity, isSave);
                if (res == null)
                {
                    repository.Rollback();
                    return null;
                }
                UpdateRedisCache(res, false);

                repository.SaveChanges();
                repository.Commit(); 
                return res;
            }
            catch (Exception ex)
            {
                repository.Rollback();
                logHelper.LogError("insert activity error :" + ex.ToString());
                return null;
            }
        }

        public override bool Update(Auc_Activity entity, bool isSave = false)
        {
            try
            {
                repository.BeginTransaction();
                entity.SetLifecycle();
                var res = base.Update(entity);
                var updateEntity = repository.FirstOrDefault(i => i.ID == entity.ID);
                if (res)
                {
                    UpdateRedisCache(updateEntity, true);
                    repository.SaveChanges();
                    repository.Commit();
                }
                else
                {
                    repository.Rollback();
                }
                return res;
            }
            catch (Exception ex)
            {
                repository.Rollback();
                logHelper.LogError("update activity error :" + ex.ToString());
                return false;
            }
        }

        public override bool Delete(Auc_Activity entity, bool isSave = false)
        {
            try
            {
                repository.BeginTransaction();
                var res = base.Delete(entity);
                if (res)
                {
                    var endTime = entity.EndTime;
                    var redisKey = $"activity:{entity.ID}:endtime";
                    Database.KeyDelete(redisKey);
                    repository.SaveChanges();
                    repository.Commit();
                }
                else
                {
                    repository.Rollback();
                }
                return res;
            }
            catch (Exception ex)
            {
                repository.Rollback();
                logHelper.LogError("delete activity error :" + ex.ToString());
                return false;
            }
        }


        // 更新Redis缓存的通用方法
        private void UpdateRedisCache(Auc_Activity activity, bool isUpdate)
        {
            try
            {
                if (activity == null) return;

                var now = DateTime.Now;

                // 只在生命周期为 "0" 时更新开始时间缓存
                if (activity.Lifecycle == 0)
                {
                    var startTimeKey = $"activity:{activity.ID}:starttime";
                    if (activity.StartTime > now)
                    {
                        var timeToLive = activity.StartTime - now;
                        Database.StringSet(startTimeKey, activity.StartTime.ToString("o"), timeToLive);
                    }
                    else
                    {
                        Database.KeyDelete(startTimeKey);
                    }
                }

                var endTimeKey = $"activity:{activity.ID}:endtime"; 
                var expiration = activity.EndTime > now ? (activity.EndTime - now) : TimeSpan.Zero;

                // 更新结束时间缓存
                if (expiration > TimeSpan.Zero)
                { 
                    Database.StringSet(endTimeKey, activity.EndTime.ToString("o"), expiration); 
                }
                else
                {
                    Database.KeyDelete(endTimeKey); 
                }
            }
            catch (Exception ex)
            {
                logHelper.LogError("UpdateRedisCache :" + ex.ToString());
                throw;
            }

        }

        public override bool InsertRange(List<Auc_Activity> entitys, bool isSave = false)
        {
            try
            {

                repository.BeginTransaction();
                entitys.ForEach(e => e.SetLifecycle());
                var res = base.InsertRange(entitys, isSave);
                if (res)
                {
                    var now = DateTime.Now;
                    foreach (var entity in entitys)
                    {
                        if (entity.Lifecycle == 0)
                        {
                            var startTimeKey = $"activity:{entity.ID}:starttime";
                            Database.StringSet(startTimeKey, entity.StartTime.ToString("o"), entity.StartTime - now);
                        }
                        var endTime = entity.EndTime;
                        var redisKey = $"activity:{entity.ID}:endtime"; 
                        Database.StringSet(redisKey, endTime.ToString("o"), endTime - now); 
                    }
                    repository.SaveChanges();
                    repository.Commit();
                }
                else
                {
                    repository.Rollback();
                }
                return res;
            }
            catch (Exception ex)
            {
                repository.Rollback();
                logHelper.LogError("batch insert activity error :" + ex.ToString());
                return false;
            }
        }

        public override bool UpdateRange(List<Auc_Activity> entitys, bool isSave = false)
        {
            try
            {
                repository.BeginTransaction();
                entitys.ForEach(e => e.SetLifecycle());
                var res = base.UpdateRange(entitys, isSave);
                if (res)
                {
                    foreach (var entity in entitys)
                    {
                        if (entity.Lifecycle == 0)
                        {
                            var startTimeKey = $"activity:{entity.ID}:starttime";
                            Database.StringSet(startTimeKey, entity.StartTime.ToString("o"), entity.StartTime - DateTime.Now);
                        }
                        var endTime = entity.EndTime;
                        var redisKey = $"activity:{entity.ID}:endtime";
                        Database.StringSet(redisKey, endTime.ToString("o"), endTime - DateTime.Now);
                    }
                    repository.SaveChanges();
                    repository.Commit();
                }
                else
                {
                    repository.Rollback();
                }
                return res;
            }
            catch (Exception ex)
            {
                repository.Rollback();
                logHelper.LogError("batch update activity error :" + ex.ToString());
                return false;
            }
        }

        public override bool DeleteRange(List<string> ids, bool isSave = false)
        {
            try
            {
                repository.BeginTransaction();
                var res = base.DeleteRange(ids, isSave);
                if (res)
                {
                    var entitys = repository.Where(i => ids.Contains(i.ID.ToString()));
                    foreach (var entity in entitys)
                    {
                        var endTime = entity.EndTime;
                        var redisKey = $"activity:{entity.ID}:endtime";
                        Database.KeyDelete(redisKey); 
                    }
                    repository.SaveChanges();
                    repository.Commit();
                }
                else
                {
                    repository.Rollback();
                }
                return res;
            }
            catch (Exception ex)
            {
                repository.Rollback();
                logHelper.LogError("batch delete activity error :" + ex.ToString());
                return false;
            }

        }
    }
}
