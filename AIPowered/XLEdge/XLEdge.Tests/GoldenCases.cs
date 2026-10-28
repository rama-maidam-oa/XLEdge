using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;

namespace XLEdge.Tests
{
    /// <summary>
    /// The inputs the golden values are taken for, and how a result is written down. Works only by reflection over the
    /// assembly it is given, so the same cases run against XLEdge 11.1.2's XLEdge.dll (to produce GoldenValues) and against the
    /// core (to check it). Formatting cases run under several current cultures: the code under test falls back to it.
    /// </summary>
    public static class GoldenCases
    {
        private static readonly string[] Cultures = { "en-US", "de-DE", "th-TH" };

        private sealed class Case
        {
            public string Id;
            public string Type;
            public string Member;
            public object[] Args;
            public bool PerCulture;
        }

        private static IEnumerable<Case> All()
        {
            const string F = "XLEdge.Helpers.XLEdgeValueFormatter";
            foreach (string v in new[] { "10/28/2026", "2026-10-28", "28-Oct-2026", "20261028", "2026-10-28T14:05:00", "2026-10-28T14:05:00.123Z",
                                         "28/10/2026", "13/13/2026", "Oct 28, 2026", "28 October 2026", "", "not a date" })
                yield return C(F, "FormatDateValue", true, v);
            foreach (var (v, t) in new (object, string)[] { ("  abc ", "STRING"), ("2026-10-28", "DATE"), ("2026-10-28 10:00:00", "datetime"), ("x", null),
                                                            (null, "STRING"), (12.5, "NUMBER"), ("", "NUMBER") })
                yield return C(F, "FormatValue", true, v, t);
            foreach (var (v, t) in new (object, string)[] { ("1,234.50", "NUMBER"), ("1.234,50", "NUMBER"), ("2026-10-28", "DATE"), ("28.10.2026", "DATE"),
                                                            ("10/28/2026 14:05:00", "DATETIME"), ("12:30", "TIME"), ("abc", "NUMBER"), ("007", "STRING"),
                                                            ("-5", "NUMBER"), ("1E3", "NUMBER"), ("", "NUMBER"), (42, "NUMBER") })
                yield return C(F, "ToExcelCellValue", true, v, t);
            foreach (object v in new object[] { "2026-10-28", "123", "12.5", "abc", "", null, "28/10/2026", "TRUE", "1,234" })
                yield return C(F, "InferDrilldownDataType", true, v);
            foreach (var (v, t) in new (object, string)[] { ("2026-10-28", "DATE"), ("2026-10-28T10:00:00", "DATETIME"), ("1,234", "NUMBER"), ("abc", "STRING"),
                                                            (null, "STRING"), ("12.50", "NUMBER"), ("28/10/2026", "DATE") })
                yield return C(F, "FormatDrilldownValue", true, v, t);
            foreach (string v in new[] { "12", "12.5", "1,234", "-3", "abc", "", "1e5", " 7 " })
                yield return C(F, "IsNumeric", true, v);
            foreach (string v in new[] { "=abc", "==x", "a=b", "", null })
                yield return C(F, "RemoveEquaSymbol", false, v);

            const string E = "XLEdge.Helpers.ApiErrorMessageExtractor";
            foreach (string body in new[] { "{\"message\":\"Bad request\"}", "{\"error\":\"invalid_grant\",\"error_description\":\"Token expired\"}",
                                            "{\"errors\":[{\"message\":\"m1\"},{\"message\":\"m2\"}]}", "{\"status\":500}",
                                            "<html><head><title>500 Internal Server Error</title></head><body><h1>Oops</h1><p>Something failed</p></body></html>",
                                            "<!DOCTYPE html><html><body>Login required</body></html>", "plain text failure", "", null })
            {
                yield return C(E, "ExtractErrorMessage", false, body);
                yield return C(E, "IsHtmlResponse", false, body);
            }

            const string A = "XLEdge.Helpers.AttachmentLinkHelper";
            foreach (string cell in new[] { "<a href='https://host/rest/secure/attachment?id=1'>file.pdf</a>", "<a href=\"https://h/x\">doc.xlsx</a>",
                                            "<a href=''></a>", "<a href=\"\"></a>", "<a href='u'></a>", "<a></a>", "plain value", "", null })
            {
                yield return C(A, "TryParseAttachmentLink", false, cell, null, null);
                yield return C(A, "IsEmptyAttachmentAnchor", false, cell);
            }
            foreach (var (tip, url) in new[] { ("/rest/secure/attachment?id=1", "https://host"), ("rest/x", "https://host/"), ("https://abs/x", "https://host"),
                                               ("", "https://host"), ("/x", ""), (null, null) })
                yield return C(A, "BuildDownloadUrl", false, tip, url);

            foreach (string v in new[] { "Edge|1|2|Name|XLSX|a|b", "Process|9", "Logs", "", null })
                yield return C("XLEdge.Helpers.EdgeRequestParser", "Parse", false, v);

            foreach (Exception ex in new Exception[] { new TimeoutException("t"), new System.Net.Http.HttpRequestException("h"), new InvalidOperationException("i"),
                                                       new System.Threading.Tasks.TaskCanceledException(), new System.Net.WebException("w"), new UnauthorizedAccessException("u") })
                yield return C("XLEdge.Helpers.ExceptionHelper", "GetFriendlyErrorMessage", false, ex);

            yield return C("XLEdge.Helpers.XLEdgeOperatorMappings", "Map", false);
            yield return C("JSON", "RoundTrip", true,
                "{\"a\":1,\"b\":1.50,\"c\":\"2\",\"d\":null,\"e\":true,\"f\":12345678901234567890,\"g\":[1,2.5],\"h\":{\"x\":\"y\"},\"i\":-0.0001}");
        }

        private static Case C(string type, string member, bool perCulture, params object[] args) =>
            new Case { Type = type, Member = member, Args = args, PerCulture = perCulture, Id = type.Split('.').Last() + "." + member + "(" + string.Join(",", args.Select(Show)) + ")" };

        /// <summary>Every case's "id = result" line, in a fixed order.</summary>
        public static string[] Run(Assembly xledge)
        {
            var lines = new List<string>();
            CultureInfo saved = Thread.CurrentThread.CurrentCulture;
            try
            {
                foreach (Case c in All())
                {
                    foreach (string culture in c.PerCulture ? Cultures : new[] { "en-US" })
                    {
                        Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);
                        lines.Add((c.PerCulture ? culture + " " : "") + c.Id + " = " + Evaluate(xledge, c));
                    }
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = saved;
            }
            return lines.ToArray();
        }

        private static string Evaluate(Assembly xledge, Case c)
        {
            try
            {
                if (c.Type == "JSON") return RoundTrip(xledge, (string)c.Args[0]);
                Type type = xledge.GetType(c.Type, throwOnError: true);
                FieldInfo field = type.GetField(c.Member, BindingFlags.Public | BindingFlags.Static);
                if (field != null) return Describe(field.GetValue(null));
                MethodInfo method = type.GetMethods(BindingFlags.Public | BindingFlags.Static)
                    .Single(m => m.Name == c.Member && m.GetParameters().Length == c.Args.Length);
                object[] args = (object[])c.Args.Clone();
                object result = method.Invoke(null, args);
                // Out parameters are part of the result.
                var outs = method.GetParameters().Where(p => p.IsOut).Select(p => p.Name + "=" + Describe(args[p.Position])).ToList();
                return outs.Count == 0 ? Describe(result) : Describe(result) + " " + string.Join(" ", outs);
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                return "throws " + ex.InnerException.GetType().FullName + ": " + ex.InnerException.Message;
            }
        }

        /// <summary>System.Text.Json through XLEdge's JsonGlobals.Options (with its NumericJsonConverter), read then written back.</summary>
        private static string RoundTrip(Assembly xledge, string json)
        {
            object options = xledge.GetType("XLEdge.Helpers.JsonGlobals", true).GetProperty("Options").GetValue(null);
            Type serializer = options.GetType().Assembly.GetType("System.Text.Json.JsonSerializer", true);
            MethodInfo deserialize = serializer.GetMethods().Single(m => m.Name == "Deserialize" && !m.IsGenericMethod &&
                m.GetParameters().Select(p => p.ParameterType).SequenceEqual(new[] { typeof(string), typeof(Type), options.GetType() }));
            MethodInfo serialize = serializer.GetMethods().Single(m => m.Name == "Serialize" && !m.IsGenericMethod &&
                m.GetParameters().Select(p => p.ParameterType).SequenceEqual(new[] { typeof(object), typeof(Type), options.GetType() }));
            object value = deserialize.Invoke(null, new object[] { json, typeof(Dictionary<string, object>), options });
            return Describe(value) + " -> " + (string)serialize.Invoke(null, new object[] { value, typeof(Dictionary<string, object>), options });
        }

        private static string Show(object value) => value == null ? "null" : value is string s ? "\"" + s + "\"" : value is Exception e ? e.GetType().Name : Convert.ToString(value, CultureInfo.InvariantCulture);

        /// <summary>A culture-independent, type-revealing text for a result.</summary>
        public static string Describe(object value)
        {
            switch (value)
            {
                case null: return "null";
                case string s: return "s:" + s.Replace("\r", "\\r").Replace("\n", "\\n");
                case bool b: return "b:" + b;
                case double d: return "d:" + d.ToString("R", CultureInfo.InvariantCulture);
                case decimal m: return "m:" + m.ToString(CultureInfo.InvariantCulture);
                case int i: return "i:" + i.ToString(CultureInfo.InvariantCulture);
                case long l: return "l:" + l.ToString(CultureInfo.InvariantCulture);
                case DateTime dt: return "t:" + dt.ToString("o", CultureInfo.InvariantCulture);
                case IDictionary dict:
                    return "{" + string.Join(";", dict.Keys.Cast<object>().Select(k => Convert.ToString(k, CultureInfo.InvariantCulture) + ":" + Describe(dict[k]))
                        .OrderBy(x => x, StringComparer.Ordinal)) + "}";
                case IEnumerable list when !(value is string):
                    return "[" + string.Join(",", list.Cast<object>().Select(Describe)) + "]";
            }
            Type type = value.GetType();
            if (type.IsPrimitive || type.IsEnum) return type.Name + ":" + Convert.ToString(value, CultureInfo.InvariantCulture);
            if (type.FullName == "System.Text.Json.JsonElement") return "json:" + value;
            var props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(p => p.GetIndexParameters().Length == 0)
                .OrderBy(p => p.Name, StringComparer.Ordinal).Select(p => p.Name + "=" + Describe(p.GetValue(value)));
            return type.Name + "{" + string.Join(",", props) + "}";
        }
    }
}
