using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.Text;

namespace WpfPepsiExcel
{
   static class MongoConnect
    {

        internal static IMongoDatabase Con()
        {
            IMongoCollection<MongoNode> Parametrs;
            string connectionString = "mongodb://172.17.0.8:27017";
            MongoClient client = new MongoClient(connectionString);

           IMongoDatabase database = client.GetDatabase("Test");
           return  database;       
        }
    }
}
