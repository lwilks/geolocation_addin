namespace GeolocationAddin.Models
{
    public class FuzzyMatchResult
    {
        public string MatchedKey { get; set; }
        public double TokenOverlapScore { get; set; }
        public double LevenshteinScore { get; set; }
        public bool IsConfident { get; set; }
    }
}
