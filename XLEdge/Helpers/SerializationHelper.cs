using System.Text.Json;
using XLEdge.Utilities;

namespace XLEdge.Helpers
{
    public static class SerializationHelper
    {
        //private static int _callCount = 0;

        public static string SerializeToJson<T>(T obj)
        {
            //_callCount++;
            //LogUtility.LogWarn($"[SerializationHelper] Call #{_callCount}: Serializing {typeof(T).Name}");

            var json = JsonSerializer.Serialize(obj, JsonGlobals.Options);
            //LogUtility.LogWarn($"[SerializationHelper] Output: {json}");

            return json;
        }
    }
}
