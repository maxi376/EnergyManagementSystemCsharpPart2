﻿using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using RabbitMQ.Client;
using System;
using System.Text;
using System.Threading;
using System.IO;
using System.Reflection;

namespace RabbitMQ.Producer
{
    public static class QueueProducer//producer2
    {
        public static void Publish(RabbitMQ.Client.IModel channel)//producer1
        {
            channel.QueueDeclare("demo-queue",
                durable: true,
                exclusive: false,
                autoDelete: false,
                arguments: null);
            var count = 0;
            var i = 1;
            var device_id = 2;
            DateTime now = DateTime.Now;

            string filePath = ".\\sensor.csv";


            filePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\sensor.csv";
            //wb = excel.Workbooks.Open(filePath);
            //wb.Close(filePath);
            //Excel.Range xlRange = xlWorksheet.UsedRange;

            //int rowCount = xlRange.Rows.Count;
            //int colCount = xlRange.Columns.Count;
            var reader = new StreamReader(filePath);

            string line;
            string[] columns = null;
            while ((line = reader.ReadLine()) != null)
            {
                columns = line.Split(',');
                //now columns array has a ll data of column in a row!
                //like:
                foreach (string s in columns)
                {
                    now = DateTime.Now;
                    var message = new
                    {
                        timestamp = $"Hello! timestamp: {now}",
                        device_id = $" device_id: {device_id}",
                        measurement_value = $" measurement_value: {s}"
                    };
                    var body = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(message));

                    channel.BasicPublish("", "demo-queue", null, body);
                    Thread.Sleep(2000);
                }
            }
            reader.Close();


        }
    }
}
