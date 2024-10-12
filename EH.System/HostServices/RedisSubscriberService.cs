using EH.Repository.Interface.Sys;
using EH.Service.Interface.Attendance;
using EH.Service.Interface.Sys;
using StackExchange.Redis;

namespace EH.System.HostServices
{
    public class RedisSubscriberService : IHostedService
    {
        private readonly IConnectionMultiplexer redis;
        private ISubscriber _subscriber;
        IServiceScopeFactory serviceScopeFactory;

        public RedisSubscriberService(IConnectionMultiplexer redis, IServiceScopeFactory serviceScopeFactory)
        {
            this.redis = redis;
            this.serviceScopeFactory = serviceScopeFactory;
        }
        public Task StartAsync(CancellationToken cancellationToken)
        {
            _subscriber = redis.GetSubscriber();
            _subscriber.Subscribe("__keyevent@0__:expired", (channel, message) =>
            {
                Console.WriteLine(111);
                using var scope = serviceScopeFactory.CreateScope();
                var key = message.ToString();
                if (key.StartsWith("product"))
                {
                    var productRepository = scope.ServiceProvider.GetRequiredService<IAucProductRepository>();
                    var id = Guid.Parse(key.Split(':')[1]);
                    var entity = productRepository.GetById(id);
                    entity.Lifecycle = key.EndsWith("endtime") ? -1 : 1;
                    productRepository.Update(entity);
                }

                if (key.StartsWith("activity"))
                {
                    var activityRepository = scope.ServiceProvider.GetRequiredService<IAucActivityRepository>();
                    var id = Guid.Parse(key.Split(':')[1]);
                    var entity = activityRepository.GetById(id);
                    entity.Lifecycle = key.EndsWith("endtime") ? -1 : 1;
                    activityRepository.Update(entity);
                }

            });
            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            _subscriber?.UnsubscribeAll();
            return Task.CompletedTask;
        }
    }

}
