using EH.Repository.Interface.Sys;
using EH.Service.Implement;
using EH.System.Commons;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore.Metadata;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using StackExchange.Redis;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace EH.Service.Interface.Sys
{
    public class AucProductService : BaseService<Auc_Product>, IAucProductService, ITransient
    {
        private readonly IAucProductRepository repository;
        private readonly IAucRecordRepository recordRepository;
        private readonly LogHelper logHelper;

        private readonly IConnectionMultiplexer redis;
        private IDatabase Database => redis.GetDatabase();
        public AucProductService(IAucProductRepository repository, IAucRecordRepository recordRepository, IConnectionMultiplexer redis, LogHelper logHelper) : base(repository, logHelper)
        {
            this.repository = repository;
            this.recordRepository = recordRepository;
            this.logHelper = logHelper;
            this.redis = redis;
        }

        public List<Auc_Product> Search(Auc_ProductSeachDto dto, out int totalCount)
        {
            totalCount = 0;
            var where = dto.pageRequest.GetWhere().Compile();
            var order = dto.pageRequest.GetOrder().Compile();
            var res = repository.GetAll().ToList();
            if (!string.IsNullOrEmpty(dto.date))
            {
                var date = dto.date.Split('-');
                res = res.Where(i => i.StartTime >= Convert.ToDateTime(date[0]) && i.EndTime <= Convert.ToDateTime(date[1])).ToList();
            }
            if (!string.IsNullOrEmpty(dto.nameOrSummary))
            {
                res = res.Where(i => i.SkuName.Contains(dto.nameOrSummary) || i.SkuCode.Contains(dto.nameOrSummary) || i.Summary.Contains(dto.nameOrSummary)).ToList();
            }
            res = res.Where(where).ToList();
            res = dto.pageRequest.isDesc ? res.OrderByDescending(order).ToList() : res.OrderBy(order).ToList();
            totalCount = res.Count();
            res = res.Skip((dto.pageRequest.PageIndex - 1) * dto.pageRequest.PageSize).Take(dto.pageRequest.PageSize).ToList();
            return res;
        }
        public async Task<string> UploadAttachments(IFormFileCollection files, string productId,string? originImgs)
        {
            var path=Directory.GetCurrentDirectory();
            var root = path+"\\Auction";

            var folder = root + "\\"+productId + "\\";
            try
            {
                List<string> fileNames = new();
                foreach (var file in files)
                {
                    if (file.Length == 0)
                        continue;
                    var filePath = folder;
                    if (!Directory.Exists(filePath))
                    {
                        Directory.CreateDirectory(filePath);
                    }
                    // = Path.Combine(folder, file.FileName);
                    filePath = folder + file.FileName;

                    using var stream = new FileStream(filePath, FileMode.Create);
                    await file.CopyToAsync(stream);
                    fileNames.Add(filePath);
                }
                var images=string.Join(",", fileNames);
                //if (originImgs!= "null")
                if (!string.IsNullOrEmpty(originImgs))
                {
                    var originList= originImgs.Split(',');
                    foreach (var origin in originList)
                    {
                        var originUrl=root + origin+",";
                        images += originUrl;
                    }
                    images = images.Substring(0, images.Length - 1);
                }

                return images;
            }
            catch (Exception ex)
            {
                logHelper.LogError("product UploadAttachments exception:" + ex.ToString());
                return "Exception";
            }

        }

        public List<Auc_MyAuctionDto> GetMyAuctions(Auc_MyAuctionRequestDto req, out int total)
        {
            string userId = req.userId;
            if (req.isEnd)
            {
                var result = repository.Entities
              .Join(recordRepository.Entities, p => p.ID.ToString(), r => r.ProductId,
                  (p, r) => new { Product = p, Record = r })
              .Where(pr => pr.Product.Lifecycle == -1 && pr.Record.UserId == userId && pr.Record.Lifecycle == 1)
              .Select(pr => new Auc_MyAuctionDto { ID = pr.Product.ID.ToString(), Code = pr.Product.SkuCode, Name = pr.Product.SkuName, Time = pr.Record.BidTime,Price=pr.Product.CurrentPrice })
              .Distinct()
              .ToList();
                total = result.Count;
                return result.Skip((req.pageIndex - 1) * req.pageSize).Take(req.pageSize).ToList();
            }
            else
            {
                //    var result = recordRepository.Entities
                //.Join(repository.Entities, r => r.ProductId, p => p.ID.ToString(),
                //    (r, p) => new { Record = r, Product = p })
                //.OrderByDescending(i=>i.Record.BidPrice)
                //.Where(rp => rp.Record.UserId == userId)
                //.Select(rp => new Auc_MyAuctionDto { ID = rp.Product.ID.ToString(), Code = rp.Product.SkuCode, Name = rp.Product.SkuName, CurrentPrice = rp.Product.CurrentPrice, Bid = rp.Record.BidPrice, ProductLifecycle = rp.Product.Lifecycle, BidLifecycle = rp.Record.Lifecycle })
                //.ToList();
                var result = recordRepository.Entities
    .Join(repository.Entities,
        r => r.ProductId,
        p => p.ID.ToString(),
        (r, p) => new { Record = r, Product = p })
    .Where(rp => rp.Record.UserId == userId)
    .AsEnumerable()  // Load data into memory for further processing
    .GroupBy(rp => rp.Product.ID)  // Group by Product ID
    .Select(g => g.OrderByDescending(rp => rp.Record.BidPrice).FirstOrDefault())  // Get the highest bid for each product
    .Select(rp => new Auc_MyAuctionDto
    {
        ID = rp.Product.ID.ToString(),
        Code = rp.Product.SkuCode,
        Name = rp.Product.SkuName,
        CurrentPrice = rp.Product.CurrentPrice,
        Bid = rp.Record.BidPrice,
        ProductLifecycle = rp.Product.Lifecycle==1?"正在拍卖": "拍卖结束",
        BidLifecycle = rp.Record.Lifecycle==1?"领先":"出局",
    })
    .OrderByDescending(dto => dto.Bid)  // Order by bid price
    .ToList();
                total = result.Count;
                return result.Skip((req.pageIndex - 1) *req.pageSize).Take(req.pageSize).ToList();
            }
             
        } 
        public override Auc_Product Insert(Auc_Product entity, bool isSave = false)
        {
            try
            {
                repository.BeginTransaction();
                entity.SetLifecycle();
                entity.CurrentPrice = entity.BasePrice;
                var res = base.Insert(entity, isSave);
                if (res == null)
                {
                    repository.Rollback();
                    return null;
                }
                UpdateRedisCache(res, false);

                repository.SaveChanges();
                repository.Commit();

                //if (res.Lifecycle == "0")
                //{
                //    var redisKey = $"product:{entity.ID}:starttime";
                //    Database.StringSet(redisKey, res.StartTime.ToString("o"), res.StartTime - DateTime.Now);
                //}
                //if (res != null)
                //{
                //    var endTime = res.EndTime;
                //    var redisKey = $"product:{res.ID}:endtime";
                //    var bidKey = $"bid:{res.ID}";
                //    Database.StringSet(redisKey, endTime.ToString("o"), endTime - DateTime.Now);
                //    Database.StringSet(bidKey, res.BasePrice.ToString(), endTime - DateTime.Now);
                //    repository.SaveChanges();
                //    repository.Commit();
                //}
                //else
                //{
                //    repository.Rollback();
                //}
                return res;
            }
            catch (Exception ex)
            {
                repository.Rollback();
                logHelper.LogError("insert product error :" + ex.ToString());
                return null;
            }
        }

        public override bool Update(Auc_Product entity, bool isSave = false)
        {
            try
            {
                repository.BeginTransaction();
                entity.SetLifecycle();
                var bidKey = $"bid:{entity.ID}";
                if(!Database.StringGet(bidKey).IsNullOrEmpty)
                    entity.CurrentPrice = (decimal)Database.StringGet(bidKey);
                else
                {
                    var currentBidEntity = recordRepository.Where(i => i.ProductId == entity.ID.ToString()).OrderByDescending(o => o.BidPrice).FirstOrDefault();
                    entity.CurrentPrice=currentBidEntity == null?entity.CurrentPrice:currentBidEntity.BidPrice;
                }

                var res = base.Update(entity);
                var updateEntity = repository.FirstOrDefault(i => i.ID == entity.ID);
                if (res)
                {
                    UpdateRedisCache(updateEntity, true);
                    //var now=DateTime.Now;
                    //if (entity.Lifecycle == "0")
                    //{
                    //    var startTimeKey = $"product:{entity.ID}:starttime";
                    //    Database.StringSet(startTimeKey, entity.StartTime.ToString("o"), entity.StartTime - now);
                    //}
                    //var endTime = entity.EndTime;
                    //var redisKey = $"product:{entity.ID}:endtime";
                    //Database.StringSet(redisKey, endTime.ToString("o"), endTime - now);
                    //if (updateEntity != null)
                    //{
                    //    var bidKey = $"bid:{updateEntity.ID}";
                    //    Database.StringSet(bidKey, updateEntity.CurrentPrice.ToString(), endTime - now);
                    //}
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
                logHelper.LogError("update product error :" + ex.ToString());
                return false;
            }
        }

        public override bool Delete(Auc_Product entity, bool isSave = false)
        {
            try
            {
                repository.BeginTransaction();
                var res = base.Delete(entity);
                if (res)
                {
                    var endTime = entity.EndTime;
                    var redisKey = $"product:{entity.ID}:endtime";
                    Database.KeyDelete(redisKey);
                    var bidKey = $"bid:{entity.ID}";
                    Database.KeyDelete(bidKey);
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
                logHelper.LogError("delete product error :" + ex.ToString());
                return false;
            }
        }


        // 更新Redis缓存的通用方法
        private void UpdateRedisCache(Auc_Product product, bool isUpdate)
        {
            try
            {
                if (product == null) return;

                var now = DateTime.Now;

                // 只在生命周期为 "0" 时更新开始时间缓存
                if (product.Lifecycle == 0)
                {
                    var startTimeKey = $"product:{product.ID}:starttime";
                    if (product.StartTime > now)
                    {
                        var timeToLive = product.StartTime - now;
                        Database.StringSet(startTimeKey, product.StartTime.ToString("o"), timeToLive);
                    }
                    else
                    {
                        Database.KeyDelete(startTimeKey);
                    }
                }

                var endTimeKey = $"product:{product.ID}:endtime";
                var bidKey = $"bid:{product.ID}";
                var expiration = product.EndTime > now ? (product.EndTime - now) : TimeSpan.Zero;

                // 更新结束时间缓存
                if (expiration > TimeSpan.Zero)
                {
                    var bidPrice = isUpdate ? product.CurrentPrice.ToString() : product.BasePrice.ToString();
                    Database.StringSet(endTimeKey, product.EndTime.ToString("o"), expiration);
                    Database.StringSet(bidKey, bidPrice, expiration);
                }
                else
                {
                    Database.KeyDelete(endTimeKey);
                    Database.KeyDelete(bidKey);
                }
            }
            catch (Exception ex)
            {
                logHelper.LogError("UpdateRedisCache :" + ex.ToString());
                throw;
            }

        }

        public override bool InsertRange(List<Auc_Product> entitys, bool isSave = false)
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
                            var startTimeKey = $"product:{entity.ID}:starttime";
                            Database.StringSet(startTimeKey, entity.StartTime.ToString("o"), entity.StartTime - now);
                        }
                        var endTime = entity.EndTime;
                        var redisKey = $"product:{entity.ID}:endtime";
                        var bidKey = $"bid:{entity.ID}";
                        Database.StringSet(redisKey, endTime.ToString("o"), endTime - now);
                        Database.StringSet(bidKey, entity.BasePrice.ToString(), endTime - now);
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
                logHelper.LogError("batch insert product error :" + ex.ToString());
                return false;
            }
        }

        public override bool UpdateRange(List<Auc_Product> entitys, bool isSave = false)
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
                            var startTimeKey = $"product:{entity.ID}:starttime";
                            Database.StringSet(startTimeKey, entity.StartTime.ToString("o"), entity.StartTime - DateTime.Now);
                        }
                        var endTime = entity.EndTime;
                        var redisKey = $"product:{entity.ID}:endtime";
                        Database.StringSet(redisKey, endTime.ToString("o"), endTime - DateTime.Now);

                        var bidKey = $"bid:{entity.ID}";
                        Database.StringSet(bidKey, entity.CurrentPrice.ToString(), endTime - DateTime.Now);

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
                logHelper.LogError("batch update product error :" + ex.ToString());
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
                        var redisKey = $"product:{entity.ID}:endtime";
                        Database.KeyDelete(redisKey);
                        var bidKey = $"bid:{entity.ID}";
                        Database.KeyDelete(bidKey);
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
                logHelper.LogError("batch update product error :" + ex.ToString());
                return false;
            }

        }

        public Auc_Product DiyInsert(Auc_ProductDto Dto)
        {
            Auc_Product entity = Dto.ToObject<Auc_Product>();
            var images = UploadAttachments(Dto.ImageFile, "","").Result;
            entity.Images = images;
            return base.Insert(entity);
        }

        public bool Import(IFormFile file, string activityId)
        {
            var data = OfficeHelper.ImportFromExcel<Auc_Product>(file);
            data.ForEach(i => { i.ActivityId = activityId; });
            try
            {
                InsertRange(data);
                return true;
            }
            catch (Exception ex)
            {
                logHelper.LogError("import product fail :" + ex.ToString());
                return false;
            }

        }

        public byte[] Export(string activityId)
        {
            var data = repository.Where(i => i.ActivityId == activityId).ToList();
            var res = OfficeHelper.ExportToExcel(data);
            return res;
        }

    }
}
