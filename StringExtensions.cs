using System.Text;

namespace Community.PowerToys.Run.Plugin.OfficeSearch
{
    public static class StringExtensions
    {
        public static string EllipsifyInterpolatedQuery(this string plaintext, string query, int lookahead = 3, bool bold = false)
        {
            const int ellipsisSize = 3;

            StringBuilder result = new StringBuilder();
            int index = plaintext.IndexOf(query);
            while (index != -1)
            {
                if (index - lookahead < ellipsisSize)
                    result.Append(plaintext.Substring(0, index - lookahead));
                else
                    result.Append(new string('.', ellipsisSize));

                if (bold)
                    result.Append("<b>");

                result.Append(plaintext.Substring(
                    Math.Max(index - lookahead, 0),
                    Math.Min(index + query.Length + lookahead, plaintext.Length - 1)));

                if (bold)
                    result.Append("</b>");

                if (index + query.Length + lookahead + ellipsisSize > plaintext.Length)
                    result.Append(plaintext.Substring(index + query.Length + lookahead + 1));
                else
                    result.Append(new string('.', ellipsisSize));

                index = plaintext.IndexOf(query, index + 1);
            }

            if (result.Length == 0)
                result.Append("ERROR: no matches found in source text");

            return result.ToString();
        }
    }
}
