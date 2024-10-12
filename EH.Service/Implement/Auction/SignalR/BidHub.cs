using EH.System.Commons;
using EH.System.Models.Entities; 
using Microsoft.AspNetCore.SignalR;
using StackExchange.Redis;
using System.Security.Claims;
using System.Threading.Channels;
using System.Threading.Tasks; 

public class BidHub : Hub,ISingleton
{
    private bool isConneted;
    public async Task SendBid(Auc_Record entity)
    {
        if(Clients != null && isConneted)
        // 广播出价到所有连接的客户端
        await Clients.All.SendAsync("ReceiveBid", entity);
    }

    public override async Task OnConnectedAsync()
    {
        isConneted = true;
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
        isConneted = false;
        var userId = Context.User?.FindFirst(ClaimTypes.NameIdentifier)?.Value;
        if (!string.IsNullOrEmpty(userId))
        {
            await Groups.RemoveFromGroupAsync(Context.ConnectionId, userId);
        }
        await base.OnDisconnectedAsync(exception);
    }

}
