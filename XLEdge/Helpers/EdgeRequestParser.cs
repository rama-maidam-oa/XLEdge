using System;
using XLEdge.Models;

namespace XLEdge.Helpers
{
    public static class EdgeRequestParser
    {
        public static EdgeRequest Parse(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                throw new ArgumentException("Input cannot be null or empty.", nameof(input));

            string[] parts = input.Split('|');

            return new EdgeRequest
            {
                ReportType = parts.Length > 0 ? parts[0] : string.Empty,
                ReportId = parts.Length > 1 ? parts[1] : string.Empty,
                ReportRunId = parts.Length > 2 ? parts[2] : string.Empty,
                ReportName = parts.Length > 3 ? parts[3] : string.Empty,
                Extension = parts.Length > 4 ? parts[4] : string.Empty,
                Extra1 = parts.Length > 5 ? parts[5] : string.Empty,
                Extra2 = parts.Length > 6 ? parts[6] : string.Empty
            };
        }
    }
}
