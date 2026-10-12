using System;

namespace Model
{
    public class Review
    {
        public int Id { get; set; }
        public int MovieId { get; set; }
        public Movie? Movie { get; set; }
        public double Rating { get; set; }
        public string Comment { get; set; } = string.Empty;
        public DateTime DateWatched { get; set; }
    }
}
