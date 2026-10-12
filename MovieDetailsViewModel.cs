using Model;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace ViewModel
{
    public class MovieDetailsViewModel : INotifyPropertyChanged, INavigableViewModel
    {
        private readonly MoviesDataStore _dataStore;
        private bool _isVisible;
        private double _editRating;
        private string _editReview = string.Empty;

        public MovieDetailsViewModel(MoviesDataStore dataStore)
        {
            _dataStore = dataStore ?? throw new ArgumentNullException(nameof(dataStore));
            _dataStore.PropertyChanged += OnDataStorePropertyChanged;

            SaveReviewCommand = new RelayCommand(SaveReview, CanSaveReview);
            SyncEditFieldsFromSelection();
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

        public Movie? SelectedMovie => _dataStore.SelectedMovie;

        public double EditRating
        {
            get => _editRating;
            set
            {
                _editRating = value;
                OnPropertyChanged();
            }
        }

        public string EditReview
        {
            get => _editReview;
            set
            {
                _editReview = value;
                OnPropertyChanged();
            }
        }

        public ICommand SaveReviewCommand { get; }

        private void SaveReview()
        {
            if (SelectedMovie == null)
                return;

            var review = SelectedMovie.Review ?? new Review { MovieId = SelectedMovie.Id };
            review.MovieId = SelectedMovie.Id;
            review.Rating = EditRating;
            review.Comment = EditReview ?? string.Empty;
            review.DateWatched = DateTime.Now;
            review.Movie = SelectedMovie;

            _dataStore.SaveReview(review);
        }

        private bool CanSaveReview()
        {
            return SelectedMovie != null;
        }

        private void SyncEditFieldsFromSelection()
        {
            if (SelectedMovie?.Review != null)
            {
                EditRating = SelectedMovie.Review.Rating;
                EditReview = SelectedMovie.Review.Comment;
            }
            else
            {
                EditRating = 0;
                EditReview = string.Empty;
            }
        }

        private void OnDataStorePropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(MoviesDataStore.SelectedMovie))
            {
                OnPropertyChanged(nameof(SelectedMovie));
                SyncEditFieldsFromSelection();
                ((RelayCommand)SaveReviewCommand).RaiseCanExecuteChanged();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}
