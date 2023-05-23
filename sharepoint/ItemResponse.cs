using System.Text.Json.Serialization;

namespace sharepoint
{
    public class ItemResponse
    {
        [JsonPropertyName("@odata.context")]
        public string? ODataContext { get; set; }

        [JsonPropertyName("value")]
        public List<Item>? Value { get; set; }
    }

    public class Item
    {
        [JsonPropertyName("@odata.etag")]
        public string? ODataEtag { get; set; }

        [JsonPropertyName("createdDateTime")]
        public DateTime? CreatedDateTime { get; set; }

        [JsonPropertyName("eTag")]
        public string? ETag { get; set; }

        [JsonPropertyName("id")]
        public string? Id { get; set; }

        [JsonPropertyName("lastModifiedDateTime")]
        public DateTime? LastModifiedDateTime { get; set; }

        [JsonPropertyName("webUrl")]
        public string? WebUrl { get; set; }

        [JsonPropertyName("fields@odata.context")]
        public string? FieldsODataContext { get; set; }

        [JsonPropertyName("fields")]
        public Fields? Fields { get; set; }

        [JsonPropertyName("driveItem@odata.context")]
        public string? DriveItemODataContext { get; set; }

        [JsonPropertyName("driveItem")]
        public DriveItem? DriveItem { get; set; }
    }

    public class Fields
    {
        // expanded fields properties
    }

    public class DriveItem
    {
        [JsonPropertyName("id")]
        public string? Id { get; set; }

        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("createdBy")]
        public IdentitySet? CreatedBy { get; set; }

        [JsonPropertyName("createdDateTime")]
        public DateTime? CreatedDateTime { get; set; }

        [JsonPropertyName("description")]
        public string? Description { get; set; }

        [JsonPropertyName("eTag")]
        public string? ETag { get; set; }

        [JsonPropertyName("lastModifiedBy")]
        public IdentitySet? LastModifiedBy { get; set; }

        [JsonPropertyName("lastModifiedDateTime")]
        public DateTime? LastModifiedDateTime { get; set; }

        [JsonPropertyName("size")]
        public long? Size { get; set; }

        [JsonPropertyName("webUrl")]
        public string? WebUrl { get; set; }

        // other DriveItem properties
    }

    public class IdentitySet
    {
        [JsonPropertyName("application")]
        public Identity? Application { get; set; }

        [JsonPropertyName("device")]
        public Identity? Device { get; set; }

        [JsonPropertyName("user")]
        public Identity? User { get; set; }
    }

    public class Identity
    {
        [JsonPropertyName("displayName")]
        public string? DisplayName { get; set; }

        [JsonPropertyName("id")]
        public string? Id { get; set; }
    }
}
