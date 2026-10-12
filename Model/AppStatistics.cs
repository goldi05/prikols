using System;

namespace Model
{
    public class AppStatistics
    {
        public int TotalMovies { get; set; }
        public int WatchedCount { get; set; }
        public double AverageRating { get; set; }
        public DateTime LastUpdated { get; set; }
    }
}