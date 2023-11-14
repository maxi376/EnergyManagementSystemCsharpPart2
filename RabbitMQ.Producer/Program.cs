﻿using RabbitMQ.Client;
using System;

namespace RabbitMQ.Producer
{
    static class Program//producer1
    {
        static void Main(string[] args)
        {
            var factory = new ConnectionFactory
            {
                Uri = new Uri("amqp://guest:guest@localhost:5672")
            };
            using var connection = factory.CreateConnection();
            using var channel = connection.CreateModel();
            QueueProducer.Publish(channel);
        }
    }
}
