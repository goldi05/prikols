using Model;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace ViewModel
{
    public class MoviesViewModel : INotifyPropertyChanged, INavigableViewModel
    {
        private readonly MoviesDataStore _dataStore;
        private bool _isVisible;

        public MoviesViewModel(MoviesDataStore dataStore)
        {
            _dataStore = dataStore ?? throw new ArgumentNullException(nameof(dataStore));
            _dataStore.PropertyChanged += OnDataStorePropertyChanged;

            DeleteMovieCommand = new RelayCommand(DeleteMovie, CanDeleteMovie);
        }

        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                _isVisible = value;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<Movie> Movies => _dataStore.Movies;

        public Movie? SelectedMovie
        {
            get => _dataStore.SelectedMovie;
            set => _dataStore.SelectedMovie = value;
        }

        public ICommand DeleteMovieCommand { get; }

        private void DeleteMovie()
        {
            _dataStore.DeleteSelectedMovie();
        }

        private bool CanDeleteMovie()
        {
            return SelectedMovie != null;
        }

        private void OnDataStorePropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(MoviesDataStore.SelectedMovie))
            {
                OnPropertyChanged(nameof(SelectedMovie));
                ((RelayCommand)DeleteMovieCommand).RaiseCanExecuteChanged();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}
