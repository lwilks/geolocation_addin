using System;
using System.Collections.Generic;
using System.Linq;
using GeolocationAddin.Models;

namespace GeolocationAddin.Helpers
{
    public static class FuzzyMatcher
    {
        public static FuzzyMatchResult FindBestMatch(
            string input,
            IEnumerable<string> candidates,
            double tokenThreshold = 0.5,
            double levThreshold = 0.4)
        {
            FuzzyMatchResult best = null;

            foreach (var candidate in candidates)
            {
                var tokenScore = TokenOverlapScore(input, candidate);
                var levScore = NormalizedLevenshtein(input, candidate);

                if (tokenScore < tokenThreshold || levScore > levThreshold)
                    continue;

                if (best == null ||
                    tokenScore > best.TokenOverlapScore ||
                    (Math.Abs(tokenScore - best.TokenOverlapScore) < 0.001 && levScore < best.LevenshteinScore))
                {
                    best = new FuzzyMatchResult
                    {
                        MatchedKey = candidate,
                        TokenOverlapScore = tokenScore,
                        LevenshteinScore = levScore,
                        IsConfident = true
                    };
                }
            }

            return best;
        }

        public static double TokenOverlapScore(string a, string b)
        {
            var tokensA = Tokenize(a);
            var tokensB = Tokenize(b);

            if (tokensA.Length == 0 && tokensB.Length == 0)
                return 1.0;
            if (tokensA.Length == 0 || tokensB.Length == 0)
                return 0.0;

            var setB = new HashSet<string>(tokensB, StringComparer.OrdinalIgnoreCase);
            int matching = tokensA.Count(t => setB.Contains(t));

            return (double)matching / Math.Max(tokensA.Length, tokensB.Length);
        }

        public static int LevenshteinDistance(string a, string b)
        {
            if (string.IsNullOrEmpty(a)) return b?.Length ?? 0;
            if (string.IsNullOrEmpty(b)) return a.Length;

            a = a.ToLowerInvariant();
            b = b.ToLowerInvariant();

            var prev = new int[b.Length + 1];
            var curr = new int[b.Length + 1];

            for (int j = 0; j <= b.Length; j++)
                prev[j] = j;

            for (int i = 1; i <= a.Length; i++)
            {
                curr[0] = i;
                for (int j = 1; j <= b.Length; j++)
                {
                    int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    curr[j] = Math.Min(
                        Math.Min(curr[j - 1] + 1, prev[j] + 1),
                        prev[j - 1] + cost);
                }

                var temp = prev;
                prev = curr;
                curr = temp;
            }

            return prev[b.Length];
        }

        public static double NormalizedLevenshtein(string a, string b)
        {
            if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b))
                return 0.0;

            int maxLen = Math.Max(a?.Length ?? 0, b?.Length ?? 0);
            return (double)LevenshteinDistance(a, b) / maxLen;
        }

        private static string[] Tokenize(string input)
        {
            if (string.IsNullOrEmpty(input))
                return Array.Empty<string>();

            return input
                .ToLowerInvariant()
                .Split(new[] { ' ', '-', '_', ':', '.', '(', ')', '[', ']' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(t => t.Length > 1)
                .ToArray();
        }
    }
}
