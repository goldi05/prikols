using System.Collections.Generic;

namespace Model
{
    public interface IRepository
    {
        List<Movie> GetAllMovies();
        List<Genre> GetAllGenres();
        List<Director> GetAllDirectors();
        List<Review> GetAllReviews();

        void AddMovie(Movie movie);
        void UpdateMovie(Movie movie);
        void DeleteMovie(Movie movie);
        void SaveReview(Review review);

        void ClearAll();
        void ClearGeneratedData();
        void GenerateTestData();

        AppStatistics GetStatistics();
        void SaveStatistics(AppStatistics stats);
    }
}
