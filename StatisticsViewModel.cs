using Model;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace ViewModel
{
    public class StatisticsViewModel : INotifyPropertyChanged, INavigableViewModel
    {
        private readonly MoviesDataStore _dataStore;
        private bool _isVisible;

        public StatisticsViewModel(MoviesDataStore dataStore)
        {
            _dataStore = dataStore ?? throw new ArgumentNullException(nameof(dataStore));
            _dataStore.PropertyChanged += OnDataStorePropertyChanged;

            GenerateDataCommand = new RelayCommand(GenerateData);
            ClearGeneratedDataCommand = new RelayCommand(ClearGeneratedData);
            ClearAllDataCommand = new RelayCommand(ClearAllData);
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

        public AppStatistics Statistics => _dataStore.Statistics;
        public ObservableCollection<Genre> Genres => _dataStore.Genres;
        public ObservableCollection<Director> Directors => _dataStore.Directors;
        public ObservableCollection<Review> Reviews => _dataStore.Reviews;

        public ICommand GenerateDataCommand { get; }
        public ICommand ClearGeneratedDataCommand { get; }
        public ICommand ClearAllDataCommand { get; }

        private void GenerateData()
        {
            _dataStore.GenerateTestData();
        }

        private void ClearGeneratedData()
        {
            _dataStore.ClearGeneratedData();
        }

        private void ClearAllData()
        {
            _dataStore.ClearAll();
        }

        private void OnDataStorePropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(MoviesDataStore.Statistics))
                OnPropertyChanged(nameof(Statistics));
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}
