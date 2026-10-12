using Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace ViewModel
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private readonly ViewModelNavigator _navigator;

        public MainWindowViewModel()
            : this(DesignerProperties.GetIsInDesignMode(new DependencyObject())
                ? new NullRepository()
                : new ConcreteRepository())
        {
        }

        public MainWindowViewModel(IRepository repository)
        {
            ArgumentNullException.ThrowIfNull(repository);

            var dataStore = new MoviesDataStore(repository);

            MainMenuContent = new MainMenuViewModel();
            MoviesContent = new MoviesViewModel(dataStore);
            AddMovieContent = new AddMovieViewModel(dataStore);
            MovieDetailsContent = new MovieDetailsViewModel(dataStore);
            StatisticsContent = new StatisticsViewModel(dataStore);

            _navigator = new ViewModelNavigator(new INavigableViewModel[]
            {
                MainMenuContent,
                MoviesContent,
                AddMovieContent,
                MovieDetailsContent,
                StatisticsContent
            });

            MainMenuBtnCommand = new RelayCommand(GoToMainMenu);
            MoviesBtnCommand = new RelayCommand(GoToMovies);
            AddMovieBtnCommand = new RelayCommand(GoToAddMovie);
            MovieDetailsBtnCommand = new RelayCommand(GoToMovieDetails);
            StatisticsBtnCommand = new RelayCommand(GoToStatistics);

            GoToMainMenu();
        }

        public MainMenuViewModel MainMenuContent { get; }
        public MoviesViewModel MoviesContent { get; }
        public AddMovieViewModel AddMovieContent { get; }
        public MovieDetailsViewModel MovieDetailsContent { get; }
        public StatisticsViewModel StatisticsContent { get; }

        public ICommand MainMenuBtnCommand { get; }
        public ICommand MoviesBtnCommand { get; }
        public ICommand AddMovieBtnCommand { get; }
        public ICommand MovieDetailsBtnCommand { get; }
        public ICommand StatisticsBtnCommand { get; }

        private void GoToMainMenu()
        {
            _navigator.NavigateTo(MainMenuContent);
        }

        private void GoToMovies()
        {
            _navigator.NavigateTo(MoviesContent);
        }

        private void GoToAddMovie()
        {
            _navigator.NavigateTo(AddMovieContent);
        }

        private void GoToMovieDetails()
        {
            _navigator.NavigateTo(MovieDetailsContent);
        }

        private void GoToStatistics()
        {
            _navigator.NavigateTo(StatisticsContent);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        private sealed class NullRepository : IRepository
        {
            public List<Movie> GetAllMovies() => new();
            public List<Genre> GetAllGenres() => new();
            public List<Director> GetAllDirectors() => new();
            public List<Review> GetAllReviews() => new();
            public void AddMovie(Movie movie) { }
            public void UpdateMovie(Movie movie) { }
            public void DeleteMovie(Movie movie) { }
            public void SaveReview(Review review) { }
            public void ClearAll() { }
            public void ClearGeneratedData() { }
            public void GenerateTestData() { }
            public AppStatistics GetStatistics() => new() { LastUpdated = DateTime.Now };
            public void SaveStatistics(AppStatistics stats) { }
        }
    }
}
