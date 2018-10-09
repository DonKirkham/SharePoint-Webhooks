using Microsoft.WindowsAzure.Storage.Table;
using Newtonsoft.Json;
using System;

namespace SharePointDemo.Models
{
    public class LastChangeEntity : TableEntity
{
        public LastChangeEntity(string category, string listId)
                : base(category, listId) { }

        public LastChangeEntity() { }
        [JsonProperty(PropertyName = "ChangeToken")]
        public String LastChangeToken { get; set; }
}
}
