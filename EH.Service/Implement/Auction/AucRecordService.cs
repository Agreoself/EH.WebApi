using EH.Repository.Interface.Sys;
using EH.Service.Implement;
using EH.System.Commons;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StackExchange.Redis;
using Microsoft.EntityFrameworkCore;
using NPOI.HSSF.Record;
using System.Threading.Channels;

namespace EH.Service.Interface.Sys
{
    public class AucRecordService : BaseService<Auc_Record>, IAucRecordService, ITransient
    {
        private readonly IAucRecordRepository repository;
        private readonly IAucProductRepository productRepository;
        private readonly LogHelper logHelper;

        private readonly IConnectionMultiplexer redis;

        private readonly BidHub bidHub;
        private readonly MyAuctionsHub myAuctionsHub;
        public AucRecordService(IAucRecordRepository repository, IAucProductRepository productRepository, LogHelper logHelper, IConnectionMultiplexer redis, BidHub bidHub, MyAuctionsHub myAuctionsHub) : base(repository, logHelper)
        {
            this.repository = repository;
            this.logHelper = logHelper;
            this.redis = redis;
            this.productRepository = productRepository;
            this.bidHub = bidHub;
            this.myAuctionsHub = myAuctionsHub;
        }

        private IDatabase Database => redis.GetDatabase();

        public byte[] Export(string productId)
        {
            var data = repository.Where(i => i.ProductId == productId).ToList();
            var res = OfficeHelper.ExportToExcel(data);
            return res;
        }

        public async Task<bool> PlaceBidAsync(Auc_Record entity)
        {
            var productKey = $"product:{entity.ProductId}";
            var bidKey = $"bid:{entity.ProductId}";

            // Check auction end time
            var endTimeStr = await Database.StringGetAsync($"{productKey}:endtime");

            if (endTimeStr.IsNullOrEmpty)
            {
                // 如果 Redis 中没有存储过期时间，视为未过期或返回特定结果
                return false; // 或者其他适当的处理方式
            }

            if (DateTime.TryParse(endTimeStr, out var endDateTime) && DateTime.Now > endDateTime)
            {
                return false; // Auction has ended
            }
            try
            {
                if (DateTime.TryParse(endTimeStr, out var endtime))
                {
                    var bidTime = entity.BidTime;
                    var currentPriceStr = await Database.StringGetAsync(bidKey);
                    //logHelper.LogInfo("currentPriceStr:" + currentPriceStr);

                    if (currentPriceStr.IsNullOrEmpty)
                    {
                        return false;
                    }
                    //logHelper.LogInfo("Convert.ToDecimal(currentPriceStr):" + Convert.ToDecimal(currentPriceStr));

                    if (Convert.ToDecimal(currentPriceStr) >= entity.BidPrice)
                    {
                        return false;
                    }
                    //logHelper.LogInfo("entity.BidPrice："+ entity.BidPrice);

                    // 比较当前时间与过期时间
                    if (bidTime > endtime)
                    {
                        return false;

                    }
                    else
                    {
                        var res = await UpdateDatabaseAsync(entity);
                        if (res)
                        {
                            await Database.StringSetAsync(bidKey, entity.BidPrice.ToString(), endtime - DateTime.Now
                           );
                            await bidHub.SendBid(entity);
                            await myAuctionsHub.MyAuctionReceiveBid(entity);

                            //await redis.GetDatabase().ListRightPushAsync("myAuctionsReceive", entity.ToJson());
                        }

                        return res;
                    }
                }
                else
                {
                    // 如果解析失败，视为未过期或返回特定结果
                    return false; // 或者其他适当的处理方式
                }
            }
            catch (Exception ex)
            {
                logHelper.LogError("placeBid error:" + ex.ToString());
                return false;
            }



            #region script
            //            var script = @"-- 获取 endTime 并转换为数字
            //local endTimeStr = redis.call('GET', KEYS[1])
            //local endTime = tonumber(endTimeStr)

            //-- 检查 endTime 是否为有效数字
            //if endTimeStr == false or endTime == nil then
            //    return ""Error: endTime is not a number or is nil""
            //end

            //-- 检查 endTime 是否小于传入的参数值
            //if endTime < tonumber(ARGV[1]) then
            //    return 0
            //end

            //-- 获取 currentBid 并转换为数字
            //local currentBidStr = redis.call('GET', KEYS[2])
            //local currentBid = tonumber(currentBidStr)

            //-- 检查 currentBid 是否为有效数字
            //if currentBidStr == false or currentBid == nil then
            //    return ""Error: currentBid is not a number or is nil""
            //end

            //-- 更新值并返回结果
            //if currentBid < tonumber(ARGV[2]) then
            //    redis.call('SET', KEYS[2], ARGV[2])
            //    redis.call('SET', KEYS[3], ARGV[3])
            //    return 1
            //else
            //    return 0
            //end";

            //    var endTimeTimestamp = ((DateTimeOffset)DateTime.UtcNow).ToUnixTimeSeconds();
            //    var result = await Database.ScriptEvaluateAsync(script,
            //new RedisKey[] { $"{productKey}:endTime", bidKey, $"{bidKey}:user" },
            //new RedisValue[] { (RedisValue)endTimeTimestamp, (RedisValue)entity.BidPrice.ToString(), (RedisValue)entity.UserId });

            //if ((long)result == 1)
            //{
            //    // Update the database
            //    return await UpdateDatabaseAsync(entity);
            //}

            #endregion
        }

        private async Task<bool> UpdateDatabaseAsync(Auc_Record entity)
        {
            repository.BeginTransaction();
            try
            {
                // Update the highest bid and its status
                var currentBidEntity = repository.Where(i => i.ProductId == entity.ProductId).OrderByDescending(o => o.BidPrice).FirstOrDefault();
                var currentHighestBid = currentBidEntity == null ? 0 : currentBidEntity.BidPrice;

                var bidStatus = entity.BidPrice >= currentHighestBid ? 1 : -1;
                var resEntity = repository.Add(entity, false);
                if (currentBidEntity != null)
                { 
                    var otherEntity = repository.Where(i => i.ID != resEntity.ID && i.ProductId == entity.ProductId);

                    await otherEntity.ForEachAsync(e =>
                    {
                        e.Lifecycle = -1;
                    });
                    repository.UpdateRange(otherEntity, false);
                }
                // Update the current price of the auction item if the bid is the highest
                if (bidStatus == 1)
                {
                    var product = productRepository.FirstOrDefault(i => i.ID.ToString() == entity.ProductId);
                    product.CurrentPrice = entity.BidPrice;
                    productRepository.Update(product);
                }
                repository.SaveChanges();
                repository.Commit();
                return true;
            }
            catch (Exception ex)
            {
                logHelper.LogError("UpdateDatabaseAsync error :" + ex.ToString());
                repository.Rollback();
                return false;
            }
        }

        public Decimal GetCurrentPrice(string productId)
        {
            decimal currentPrice = 0;
            var endTimeStr = Database.StringGet($"product:{productId}:endtime");
            var bidKey = $"bid:{productId}";
            currentPrice = (decimal)Database.StringGet(bidKey);
            if (currentPrice == 0)
            {
                currentPrice = productRepository.FirstOrDefault(f => f.ID.ToString() == productId).CurrentPrice;
            }
            return currentPrice;
        }

        public override List<Auc_Record> GetPageList(PageRequest<Auc_Record> request, out int totalCount)
        {
            var userId = "";
            if (!string.IsNullOrEmpty(request.where))
                userId = request.where.Split('=')[1];
            request.where = null;
            var res = base.GetPageList(request, out totalCount);
            res.ForEach(e =>
            {
                var productKey = $"product:{e.ProductId}";
                var endTimeStr = Database.StringGet($"{productKey}:endtime");

                if (endTimeStr.IsNullOrEmpty)
                {
                    // 如果 Redis 中没有存储过期时间，视为未过期或返回特定结果
                    //return false; // 或者其他适当的处理方式
                    return;
                }

                if (DateTime.TryParse(endTimeStr, out var endDateTime) && DateTime.Now > endDateTime)
                {
                    return;
                }

                e.UserId = (e.UserId == userId || userId == "admin") ? e.UserId : "*******";
                e.CreateBy = "*******";
                e.ModifyBy = "*******";
            });
            return res;
        }
    }
}
