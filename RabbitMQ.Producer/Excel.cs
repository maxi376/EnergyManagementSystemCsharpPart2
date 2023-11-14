using Microsoft.Office.Interop.Excel;
using RabbitMQ.Client;
using System;
using System.Collections.Generic;
using System.Text;
using IModel = RabbitMQ.Client.IModel;

namespace RabbitMQ.Producer
{
    public static class Excel
    {


        public static void ReadCells1(IModel channel)
        {
                string filePath = ".\\sensor.csv";
                Microsoft.Office.Interop.Excel.Application excel = new
                Microsoft.Office.Interop.Excel.Application();
                Workbook wb;
                Worksheet ws;

                wb = excel.Workbooks.Open(filePath);
                ws = wb.Worksheets[1];

            Microsoft.Office.Interop.Excel.Range cell = ws.Range["A1:A1"];
                foreach (string Result in cell.Value)
                {
                    //MessageBox.Show(Result);
                }

        }
    }
}
