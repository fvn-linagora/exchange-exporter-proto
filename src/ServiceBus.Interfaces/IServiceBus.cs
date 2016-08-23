using System;
using System.Reactive.Concurrency;
using System.Threading.Tasks;

namespace ServiceBus.Interfaces
{
    public interface IMessageEmitter
    {
        Task Send(ICommand command);

        Task Publish(IEvent ev);
    }

    public interface IServiceBusClient
    {
        IDisposable Subscribe(object subscriber, IScheduler scheduler = null);
    }

    public interface IServiceBus : IServiceBusClient, IMessageEmitter
    {
    }

    public interface IMessage
    {
    }

    public interface ICommand
    {
    }
    public interface IEvent
    {
    }
}
