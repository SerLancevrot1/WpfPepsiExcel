using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.Text;

namespace WpfPepsiExcel
{
   static class MongoConnect
    {
        // подключение к ДБ
        internal static IMongoDatabase ConElectr()
        {
            IMongoCollection<MongoNode> Parametrs;
            string connectionString = "mongodb://172.17.0.8:27017";
            MongoClient client = new MongoClient(connectionString);

           IMongoDatabase database = client.GetDatabase("Electricity");
           return  database;       
        }

        internal static IMongoDatabase ConWater()
        {
            IMongoCollection<MongoNode> Parametrs;
            string connectionString = "mongodb://172.17.0.8:27017";
            MongoClient client = new MongoClient(connectionString);

            IMongoDatabase database = client.GetDatabase("Water");
            return database;
        }

        internal static IMongoDatabase ConGas()
        {
            IMongoCollection<MongoNode> Parametrs;
            string connectionString = "mongodb://172.17.0.8:27017";
            MongoClient client = new MongoClient(connectionString);

            IMongoDatabase database = client.GetDatabase("Gas");
            return database;
        }
    }
}
