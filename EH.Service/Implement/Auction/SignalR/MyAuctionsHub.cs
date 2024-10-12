using EH.Service.Interface.Sys;
using EH.System.Commons;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using Microsoft.AspNetCore.SignalR;
using Newtonsoft.Json;
using StackExchange.Redis;
using System.Security.Claims;
using System.Threading.Channels;
using System.Threading.Tasks;

public class MyAuctionsHub : Hub, ISingleton
{
    private static readonly Dictionary<string, List<string>> _messageCache = new();

    private readonly IConnectionMultiplexer redis;

    private bool isConneted;
    //private readonly IAucProductService service;
    public MyAuctionsHub(IConnectionMultiplexer redis)
    {
        this.redis = redis;
    }
    public async Task MyAuctionReceiveBid(Auc_Record entity)
    {
        // 广播出价到所有连接的客户端
        if (Clients != null&&isConneted)
        {
            var s = Context;
            var userId = Context.User?.Identity.Name;
            if (userId != null)
            {
                //Auc_MyAuctionRequestDto dto = new()
                //{
                //    userId = userId.Split("\\")[1],
                //    pageSize = 1000,
                //    pageIndex = 0,
                //    isEnd = false,
                //};
                //var products = service.GetMyAuctions(dto, out int total);
                //if (products.Select(i => i.ID.ToString()).Contains(entity.ProductId))
                    await Clients.Group(userId).SendAsync("MyAuctionReceiveBid", entity);
            }

        }

        //if (_messageCache.ContainsKey(Context.ConnectionId))
        //{
        //    _messageCache[Context.ConnectionId].Add(entity.ToJson());
        //}
        //await Clients.All.SendAsync("ReceiveMessage", entity.ToJson());

        //if (_messageCache.ContainsKey(Context.ConnectionId))
        //{
        //    _messageCache[Context.ConnectionId].Add(entity.ToJson());
        //}
        //else
        //{
        //    _messageCache[Context.ConnectionId] = new List<string> { entity.ToJson() };
        //}

        //await Clients.All.SendAsync("MyAuctionReceiveBid", entity);
    }

    public override async Task OnConnectedAsync()
    {
        isConneted=true;
        var userId = Context.User?.Identity.Name;
        if (!string.IsNullOrEmpty(userId))
        {
            // 将用户 ID 作为组名
            await Groups.AddToGroupAsync(Context.ConnectionId, userId);
        }
        await base.OnConnectedAsync();
    }

    public override async Task OnDisconnectedAsync(Exception exception)
    {
        isConneted=false;
        var userId = Context.User?.FindFirst(ClaimTypes.NameIdentifier)?.Value;
        if (!string.IsNullOrEmpty(userId))
        {
            await Groups.RemoveFromGroupAsync(Context.ConnectionId, userId);
        }
        await base.OnDisconnectedAsync(exception);
    }


    //public override async Task OnConnectedAsync()
    //{
    //    try
    //    {
    //        if (_messageCache.TryGetValue(Context.ConnectionId, out var messages))
    //        {
    //            foreach (var message in messages)
    //            {
    //                await Clients.Caller.SendAsync("MyAuctionReceiveBid", message);
    //            }
    //            // Clear cached messages after sending
    //            _messageCache.Remove(Context.ConnectionId);
    //        }
    //        else
    //        {
    //            _messageCache[Context.ConnectionId] = new List<string>();
    //        }

    //        //var messages = await redis.GetDatabase().ListRangeAsync("myAuctionsReceive");
    //        //foreach (var message in messages)
    //        //{
    //        //    await Clients.Caller.SendAsync("MyAuctionReceiveBid",JsonConvert.DeserializeObject< Auc_Record >(message.ToString()));
    //        //}

    //        //// Clear the queue after sending messages
    //        //await redis.GetDatabase().KeyDeleteAsync("myAuctionsReceive");

    //        await base.OnConnectedAsync();
    //    }
    //    catch (Exception ex)
    //    {
    //        throw;
    //    }
    //    // Fetch messages from Redis queue and send to client

    //}

    //public override Task OnDisconnectedAsync(Exception? exception)
    //{
    //    if (_messageCache.ContainsKey(Context.ConnectionId))
    //    {
    //        _messageCache.Remove(Context.ConnectionId);
    //    }

    //    return base.OnDisconnectedAsync(exception);
    //}
}
