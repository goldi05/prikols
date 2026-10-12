using Model;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace ViewModel
{
    public class MoviesDataStore : INotifyPropertyChanged
    {
        private readonly IRepository _repository;
        private Movie? _selectedMovie;
        private AppStatistics _statistics = new() { LastUpdated = DateTime.Now };

        public MoviesDataStore(IRepository repository)
        {
            _repository = repository ?? throw new ArgumentNullException(nameof(repository));
            Movies = new ObservableCollection<Movie>();
            Genres = new ObservableCollection<Genre>();
            Directors = new ObservableCollection<Director>();
            Reviews = new ObservableCollection<Review>();
            Refresh();
        }

        public ObservableCollection<Movie> Movies { get; }
        public ObservableCollection<Genre> Genres { get; }
        public ObservableCollection<Director> Directors { get; }
        public ObservableCollection<Review> Reviews { get; }

        public Movie? SelectedMovie
        {
            get => _selectedMovie;
            set
            {
                if (ReferenceEquals(_selectedMovie, value))
                    return;

                _selectedMovie = value;
                OnPropertyChanged();
            }
        }

        public AppStatistics Statistics
        {
            get => _statistics;
            private set
            {
                _statistics = value;
                OnPropertyChanged();
            }
        }

        public void Refresh()
        {
            var selectedId = SelectedMovie?.Id;

            Movies.Clear();
            foreach (var movie in _repository.GetAllMovies())
                Movies.Add(movie);

            Genres.Clear();
            foreach (var genre in _repository.GetAllGenres())
                Genres.Add(genre);

            Directors.Clear();
            foreach (var director in _repository.GetAllDirectors())
                Directors.Add(director);

            Reviews.Clear();
            foreach (var review in _repository.GetAllReviews())
                Reviews.Add(review);

            SelectedMovie = selectedId.HasValue
                ? Movies.FirstOrDefault(m => m.Id == selectedId.Value)
                : null;

            SaveStatisticsToDb();
        }

        public void AddMovie(Movie movie)
        {
            _repository.AddMovie(movie);
            Refresh();
        }

        public void DeleteSelectedMovie()
        {
            if (SelectedMovie == null)
                return;

            _repository.DeleteMovie(SelectedMovie);
            SelectedMovie = null;
            Refresh();
        }

        public void SaveReview(Review review)
        {
            _repository.SaveReview(review);
            Refresh();
        }

        public void GenerateTestData()
        {
            _repository.GenerateTestData();
            Refresh();
        }

        public void ClearGeneratedData()
        {
            _repository.ClearGeneratedData();
            Refresh();
        }

        public void ClearAll()
        {
            _repository.ClearAll();
            SelectedMovie = null;
            Refresh();
        }

        private void SaveStatisticsToDb()
        {
            var ratings = Movies
                .Where(m => m.Review != null && m.Review.Rating > 0)
                .Select(m => m.Review!.Rating)
                .ToList();

            var stats = new AppStatistics
            {
                TotalMovies = Movies.Count,
                WatchedCount = Movies.Count(m => m.Review != null),
                AverageRating = ratings.Count == 0 ? 0 : ratings.Average(),
                LastUpdated = DateTime.Now
            };

            _repository.SaveStatistics(stats);
            Statistics = _repository.GetStatistics();
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}
