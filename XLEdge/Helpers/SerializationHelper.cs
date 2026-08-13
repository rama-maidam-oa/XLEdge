using System.Text.Json;

namespace XLEdge.Helpers
{
    public static class SerializationHelper
    {

        public static string SerializeToJson<T>(T obj)
        {
            var json = JsonSerializer.Serialize(obj, JsonGlobals.Options);

            return json;
        }
    }
}
