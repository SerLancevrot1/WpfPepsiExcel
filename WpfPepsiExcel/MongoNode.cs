using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace WpfPepsiExcel
{
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
}
