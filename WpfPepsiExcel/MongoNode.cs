using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace WpfPepsiExcel
{
    // Классы для получения данных из ДБ
    class MongoNode
    {
        [BsonId]
        public ObjectId _id { get; set; }
        public int ID { get; set; }
        public DateTime dateTime { get; set; }
        [BsonRepresentation(BsonType.Int64, AllowTruncation = true)]
        public float wP_in { get; set; }
        [BsonRepresentation(BsonType.Int64, AllowTruncation = true)]
        public float WP_out { get; set; }
        [BsonRepresentation(BsonType.Int64, AllowTruncation = true)]
        public float WQ_in { get; set; }
        [BsonRepresentation(BsonType.Int64, AllowTruncation = true)]
        public float WQ_oup { get; set; }
        [BsonRepresentation(BsonType.Int64, AllowTruncation = true)]
        public float WQ { get; set; }
    }

    class MongoNodeWater
    {
        [BsonId]
        public ObjectId _id { get; set; }
        public int ID { get; set; }
        public string name { get; set; }
        public float value { get; set; }
        [BsonDateTimeOptions]
        public DateTime dateTime { get; set; }
    }

    class MongoNodeGas
    {
        [BsonId]
        public ObjectId _id { get; set; }
        public int ID { get; set; }
        public string name { get; set; }
        public float value { get; set; }
        [BsonDateTimeOptions]
        public DateTime dateTime { get; set; }
        public bool IsWork { get; set; }


    }
}
