using Model;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace ViewModel
{
    public class AddMovieViewModel : INotifyPropertyChanged, INavigableViewModel
    {
        private readonly MoviesDataStore _dataStore;
        private bool _isVisible;
        private string _newTitle = string.Empty;
        private int _newYear = DateTime.Now.Year;
        private string _newGenre = string.Empty;
        private string _newDirector = string.Empty;
        private string _newDescription = string.Empty;

        public AddMovieViewModel(MoviesDataStore dataStore)
        {
            _dataStore = dataStore ?? throw new ArgumentNullException(nameof(dataStore));
            AddNewMovieCommand = new RelayCommand(AddNewMovie, CanAddNewMovie);
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

        public string NewTitle
        {
            get => _newTitle;
            set
            {
                _newTitle = value;
                OnPropertyChanged();
                ((RelayCommand)AddNewMovieCommand).RaiseCanExecuteChanged();
            }
        }

        public int NewYear
        {
            get => _newYear;
            set
            {
                _newYear = value;
                OnPropertyChanged();
            }
        }

        public string NewGenre
        {
            get => _newGenre;
            set
            {
                _newGenre = value;
                OnPropertyChanged();
            }
        }

        public string NewDirector
        {
            get => _newDirector;
            set
            {
                _newDirector = value;
                OnPropertyChanged();
            }
        }

        public string NewDescription
        {
            get => _newDescription;
            set
            {
                _newDescription = value;
                OnPropertyChanged();
            }
        }

        public ICommand AddNewMovieCommand { get; }

        private void AddNewMovie()
        {
            var movie = new Movie
            {
                Title = NewTitle,
                ReleaseYear = NewYear,
                Description = NewDescription,
                Genre = string.IsNullOrWhiteSpace(NewGenre) ? null : new Genre { Name = NewGenre.Trim() },
                Director = string.IsNullOrWhiteSpace(NewDirector) ? null : new Director { FullName = NewDirector.Trim() }
            };

            _dataStore.AddMovie(movie);

            NewTitle = string.Empty;
            NewGenre = string.Empty;
            NewDirector = string.Empty;
            NewDescription = string.Empty;
            NewYear = DateTime.Now.Year;
        }

        private bool CanAddNewMovie()
        {
            return !string.IsNullOrWhiteSpace(NewTitle);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}
