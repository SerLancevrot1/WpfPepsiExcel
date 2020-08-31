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
            IMongoCollection<MongoNodeElectricity> Parametrs;
            string connectionString = "mongodb://172.17.0.8:27017";
            MongoClient client = new MongoClient(connectionString);

           IMongoDatabase database = client.GetDatabase("Test");
           return  database;       
        }

        internal static IMongoDatabase ConWater()
        {
            IMongoCollection<MongoNodeElectricity> Parametrs;
            string connectionString = "mongodb://172.17.0.8:27017";
            MongoClient client = new MongoClient(connectionString);

            IMongoDatabase database = client.GetDatabase("Water");
            return database;
        }

        internal static IMongoDatabase ConGas()
        {
            IMongoCollection<MongoNodeElectricity> Parametrs;
            string connectionString = "mongodb://172.17.0.8:27017";
            MongoClient client = new MongoClient(connectionString);

            IMongoDatabase database = client.GetDatabase("Test");
            return database;
        }
    }
}
