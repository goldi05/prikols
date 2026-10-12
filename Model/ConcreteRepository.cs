using System;
using System.Collections.Generic;
using Microsoft.Data.Sqlite;

namespace Model
{
    public class ConcreteRepository : IRepository
    {
        private readonly string _connectionString;

        public ConcreteRepository(string dbPath = "moviediary.db")
        {
            _connectionString = $"Data Source={dbPath}";
            InitializeDatabase();
        }

        private SqliteConnection OpenConnection()
        {
            var connection = new SqliteConnection(_connectionString);
            connection.Open();
            using var pragma = new SqliteCommand("PRAGMA foreign_keys = ON;", connection);
            pragma.ExecuteNonQuery();
            return connection;
        }

        private void InitializeDatabase()
        {
            using var connection = OpenConnection();
            DropLegacyMovieTable(connection);

            var createDirectors = @"
                CREATE TABLE IF NOT EXISTS Directors (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    FullName TEXT NOT NULL UNIQUE
                );";

            var createGenres = @"
                CREATE TABLE IF NOT EXISTS Genres (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Name TEXT NOT NULL UNIQUE
                );";

            var createMovies = @"
                CREATE TABLE IF NOT EXISTS Movies (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Title TEXT NOT NULL,
                    ReleaseYear INTEGER NOT NULL,
                    Description TEXT,
                    DirectorId INTEGER,
                    GenreId INTEGER,
                    IsTestData INTEGER NOT NULL DEFAULT 0,
                    FOREIGN KEY (DirectorId) REFERENCES Directors(Id),
                    FOREIGN KEY (GenreId) REFERENCES Genres(Id)
                );";

            var createReviews = @"
                CREATE TABLE IF NOT EXISTS Reviews (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    MovieId INTEGER NOT NULL UNIQUE,
                    Rating REAL NOT NULL,
                    Comment TEXT,
                    DateWatched TEXT NOT NULL,
                    FOREIGN KEY (MovieId) REFERENCES Movies(Id) ON DELETE CASCADE
                );";

            var createStats = @"
                CREATE TABLE IF NOT EXISTS Statistics (
                    Id INTEGER PRIMARY KEY CHECK (Id = 1),
                    TotalMovies INTEGER NOT NULL,
                    WatchedCount INTEGER NOT NULL,
                    AverageRating REAL NOT NULL,
                    LastUpdated TEXT NOT NULL
                );";

            using (var command = new SqliteCommand(createDirectors, connection))
                command.ExecuteNonQuery();
            using (var command = new SqliteCommand(createGenres, connection))
                command.ExecuteNonQuery();
            using (var command = new SqliteCommand(createMovies, connection))
                command.ExecuteNonQuery();
            using (var command = new SqliteCommand(createReviews, connection))
                command.ExecuteNonQuery();
            using (var command = new SqliteCommand(createStats, connection))
                command.ExecuteNonQuery();

            using var countCommand = new SqliteCommand("SELECT COUNT(*) FROM Movies;", connection);
            var movieCount = Convert.ToInt32(countCommand.ExecuteScalar());
            if (movieCount == 0)
            {
                connection.Close();
                GenerateTestData();
            }
        }

        private static void DropLegacyMovieTable(SqliteConnection connection)
        {
            using var exists = new SqliteCommand(
                "SELECT name FROM sqlite_master WHERE type='table' AND name='Movies';",
                connection);
            if (exists.ExecuteScalar() == null)
                return;

            var isLegacy = false;
            using (var info = new SqliteCommand("PRAGMA table_info(Movies);", connection))
            using (var reader = info.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (string.Equals(reader.GetString(1), "Genre", StringComparison.OrdinalIgnoreCase))
                    {
                        isLegacy = true;
                        break;
                    }
                }
            }

            if (!isLegacy)
                return;

            using var drop = new SqliteCommand("DROP TABLE IF EXISTS Movies;", connection);
            drop.ExecuteNonQuery();
        }

        public List<Movie> GetAllMovies()
        {
            var movies = new List<Movie>();
            using var connection = OpenConnection();
            var query = @"
                SELECT m.Id, m.Title, m.ReleaseYear, m.Description, m.IsTestData,
                       d.Id, d.FullName,
                       g.Id, g.Name,
                       r.Id, r.Rating, r.Comment, r.DateWatched
                FROM Movies m
                LEFT JOIN Directors d ON m.DirectorId = d.Id
                LEFT JOIN Genres g ON m.GenreId = g.Id
                LEFT JOIN Reviews r ON r.MovieId = m.Id
                ORDER BY m.Id;";

            using var command = new SqliteCommand(query, connection);
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                var movie = new Movie
                {
                    Id = reader.GetInt32(0),
                    Title = reader.GetString(1),
                    ReleaseYear = reader.GetInt32(2),
                    Description = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                    IsTestData = reader.GetInt32(4) == 1
                };

                if (!reader.IsDBNull(5))
                {
                    movie.DirectorId = reader.GetInt32(5);
                    movie.Director = new Director
                    {
                        Id = reader.GetInt32(5),
                        FullName = reader.GetString(6)
                    };
                }

                if (!reader.IsDBNull(7))
                {
                    movie.GenreId = reader.GetInt32(7);
                    movie.Genre = new Genre
                    {
                        Id = reader.GetInt32(7),
                        Name = reader.GetString(8)
                    };
                }

                if (!reader.IsDBNull(9))
                {
                    movie.Review = new Review
                    {
                        Id = reader.GetInt32(9),
                        MovieId = movie.Id,
                        Rating = reader.GetDouble(10),
                        Comment = reader.IsDBNull(11) ? string.Empty : reader.GetString(11),
                        DateWatched = DateTime.Parse(reader.GetString(12)),
                        Movie = movie
                    };
                }

                movies.Add(movie);
            }

            return movies;
        }

        public List<Genre> GetAllGenres()
        {
            var genres = new List<Genre>();
            using var connection = OpenConnection();
            using var command = new SqliteCommand("SELECT Id, Name FROM Genres ORDER BY Name;", connection);
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                genres.Add(new Genre
                {
                    Id = reader.GetInt32(0),
                    Name = reader.GetString(1)
                });
            }

            return genres;
        }

        public List<Director> GetAllDirectors()
        {
            var directors = new List<Director>();
            using var connection = OpenConnection();
            using var command = new SqliteCommand("SELECT Id, FullName FROM Directors ORDER BY FullName;", connection);
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                directors.Add(new Director
                {
                    Id = reader.GetInt32(0),
                    FullName = reader.GetString(1)
                });
            }

            return directors;
        }

        public List<Review> GetAllReviews()
        {
            var reviews = new List<Review>();
            using var connection = OpenConnection();
            using var command = new SqliteCommand(
                "SELECT Id, MovieId, Rating, Comment, DateWatched FROM Reviews ORDER BY DateWatched;",
                connection);
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                reviews.Add(new Review
                {
                    Id = reader.GetInt32(0),
                    MovieId = reader.GetInt32(1),
                    Rating = reader.GetDouble(2),
                    Comment = reader.IsDBNull(3) ? string.Empty : reader.GetString(3),
                    DateWatched = DateTime.Parse(reader.GetString(4))
                });
            }

            return reviews;
        }

        public void AddMovie(Movie movie)
        {
            using var connection = OpenConnection();
            using var transaction = connection.BeginTransaction();

            movie.DirectorId = EnsureDirector(connection, transaction, movie.Director);
            movie.GenreId = EnsureGenre(connection, transaction, movie.Genre);

            var query = @"
                INSERT INTO Movies (Title, ReleaseYear, Description, DirectorId, GenreId, IsTestData)
                VALUES (@Title, @ReleaseYear, @Description, @DirectorId, @GenreId, @IsTestData);
                SELECT last_insert_rowid();";

            using (var command = new SqliteCommand(query, connection, transaction))
            {
                command.Parameters.AddWithValue("@Title", movie.Title);
                command.Parameters.AddWithValue("@ReleaseYear", movie.ReleaseYear);
                command.Parameters.AddWithValue("@Description", (object?)movie.Description ?? DBNull.Value);
                command.Parameters.AddWithValue("@DirectorId", (object?)movie.DirectorId ?? DBNull.Value);
                command.Parameters.AddWithValue("@GenreId", (object?)movie.GenreId ?? DBNull.Value);
                command.Parameters.AddWithValue("@IsTestData", movie.IsTestData ? 1 : 0);
                movie.Id = Convert.ToInt32(command.ExecuteScalar());
            }

            if (movie.Review != null)
            {
                movie.Review.MovieId = movie.Id;
                SaveReview(connection, transaction, movie.Review);
            }

            transaction.Commit();
        }

        public void UpdateMovie(Movie movie)
        {
            using var connection = OpenConnection();
            using var transaction = connection.BeginTransaction();

            movie.DirectorId = EnsureDirector(connection, transaction, movie.Director);
            movie.GenreId = EnsureGenre(connection, transaction, movie.Genre);

            var query = @"
                UPDATE Movies
                SET Title = @Title,
                    ReleaseYear = @ReleaseYear,
                    Description = @Description,
                    DirectorId = @DirectorId,
                    GenreId = @GenreId,
                    IsTestData = @IsTestData
                WHERE Id = @Id;";

            using (var command = new SqliteCommand(query, connection, transaction))
            {
                command.Parameters.AddWithValue("@Id", movie.Id);
                command.Parameters.AddWithValue("@Title", movie.Title);
                command.Parameters.AddWithValue("@ReleaseYear", movie.ReleaseYear);
                command.Parameters.AddWithValue("@Description", (object?)movie.Description ?? DBNull.Value);
                command.Parameters.AddWithValue("@DirectorId", (object?)movie.DirectorId ?? DBNull.Value);
                command.Parameters.AddWithValue("@GenreId", (object?)movie.GenreId ?? DBNull.Value);
                command.Parameters.AddWithValue("@IsTestData", movie.IsTestData ? 1 : 0);
                command.ExecuteNonQuery();
            }

            transaction.Commit();
        }

        public void DeleteMovie(Movie movie)
        {
            using var connection = OpenConnection();
            using var command = new SqliteCommand("DELETE FROM Movies WHERE Id = @Id;", connection);
            command.Parameters.AddWithValue("@Id", movie.Id);
            command.ExecuteNonQuery();
        }

        public void SaveReview(Review review)
        {
            using var connection = OpenConnection();
            SaveReview(connection, null, review);
        }

        private static void SaveReview(SqliteConnection connection, SqliteTransaction? transaction, Review review)
        {
            var upsert = @"
                INSERT INTO Reviews (MovieId, Rating, Comment, DateWatched)
                VALUES (@MovieId, @Rating, @Comment, @DateWatched)
                ON CONFLICT(MovieId) DO UPDATE SET
                    Rating = excluded.Rating,
                    Comment = excluded.Comment,
                    DateWatched = excluded.DateWatched;";

            using (var command = new SqliteCommand(upsert, connection))
            {
                if (transaction != null)
                    command.Transaction = transaction;
                command.Parameters.AddWithValue("@MovieId", review.MovieId);
                command.Parameters.AddWithValue("@Rating", review.Rating);
                command.Parameters.AddWithValue("@Comment", (object?)review.Comment ?? DBNull.Value);
                command.Parameters.AddWithValue("@DateWatched", review.DateWatched.ToString("o"));
                command.ExecuteNonQuery();
            }

            using (var command = new SqliteCommand("SELECT Id FROM Reviews WHERE MovieId = @MovieId;", connection))
            {
                if (transaction != null)
                    command.Transaction = transaction;
                command.Parameters.AddWithValue("@MovieId", review.MovieId);
                review.Id = Convert.ToInt32(command.ExecuteScalar());
            }
        }

        public void ClearAll()
        {
            using var connection = OpenConnection();
            using (var command = new SqliteCommand("DELETE FROM Reviews;", connection))
                command.ExecuteNonQuery();
            using (var command = new SqliteCommand("DELETE FROM Movies;", connection))
                command.ExecuteNonQuery();
            using (var command = new SqliteCommand("DELETE FROM Directors;", connection))
                command.ExecuteNonQuery();
            using (var command = new SqliteCommand("DELETE FROM Genres;", connection))
                command.ExecuteNonQuery();
            using (var command = new SqliteCommand("DELETE FROM Statistics;", connection))
                command.ExecuteNonQuery();
        }

        public void ClearGeneratedData()
        {
            using var connection = OpenConnection();
            using var command = new SqliteCommand("DELETE FROM Movies WHERE IsTestData = 1;", connection);
            command.ExecuteNonQuery();
        }

        public AppStatistics GetStatistics()
        {
            using var connection = OpenConnection();
            var query = "SELECT TotalMovies, WatchedCount, AverageRating, LastUpdated FROM Statistics WHERE Id = 1";

            using var command = new SqliteCommand(query, connection);
            using var reader = command.ExecuteReader();
            if (reader.Read())
            {
                return new AppStatistics
                {
                    TotalMovies = reader.GetInt32(0),
                    WatchedCount = reader.GetInt32(1),
                    AverageRating = reader.GetDouble(2),
                    LastUpdated = DateTime.Parse(reader.GetString(3))
                };
            }

            return new AppStatistics { LastUpdated = DateTime.Now };
        }

        public void SaveStatistics(AppStatistics stats)
        {
            using var connection = OpenConnection();
            var query = @"
                INSERT INTO Statistics (Id, TotalMovies, WatchedCount, AverageRating, LastUpdated)
                VALUES (1, @TotalMovies, @WatchedCount, @AverageRating, @LastUpdated)
                ON CONFLICT(Id) DO UPDATE SET
                    TotalMovies = excluded.TotalMovies,
                    WatchedCount = excluded.WatchedCount,
                    AverageRating = excluded.AverageRating,
                    LastUpdated = excluded.LastUpdated;";

            using var command = new SqliteCommand(query, connection);
            command.Parameters.AddWithValue("@TotalMovies", stats.TotalMovies);
            command.Parameters.AddWithValue("@WatchedCount", stats.WatchedCount);
            command.Parameters.AddWithValue("@AverageRating", stats.AverageRating);
            command.Parameters.AddWithValue("@LastUpdated", stats.LastUpdated.ToString("o"));
            command.ExecuteNonQuery();
        }

        public void GenerateTestData()
        {
            var nolan = new Director { FullName = "Кристофер Нолан" };
            var darabont = new Director { FullName = "Фрэнк Дарабонт" };
            var wachowski = new Director { FullName = "Вачовски" };

            var drama = new Genre { Name = "Драма" };
            var scifi = new Genre { Name = "Фантастика" };

            var testMovies = new List<Movie>
            {
                new Movie
                {
                    Title = "Побег из Шоушенка",
                    ReleaseYear = 1994,
                    Description = "Успешный банкир Энди Дюфрейн обвинен в убийстве...",
                    Director = darabont,
                    Genre = drama,
                    IsTestData = true,
                    Review = new Review
                    {
                        Rating = 9.5,
                        Comment = "Замечательный фильм.",
                        DateWatched = new DateTime(2024, 1, 10)
                    }
                },
                new Movie
                {
                    Title = "Начало",
                    ReleaseYear = 2010,
                    Description = "Кобб – вор, умеющий извлекать секреты...",
                    Director = nolan,
                    Genre = scifi,
                    IsTestData = true,
                    Review = new Review
                    {
                        Rating = 8.8,
                        Comment = "Запутанный сюжет.",
                        DateWatched = new DateTime(2024, 3, 2)
                    }
                },
                new Movie
                {
                    Title = "Интерстеллар",
                    ReleaseYear = 2014,
                    Description = "Исследователи отправляются сквозь червоточину...",
                    Director = nolan,
                    Genre = scifi,
                    IsTestData = true
                },
                new Movie
                {
                    Title = "Матрица",
                    ReleaseYear = 1999,
                    Description = "Хакер Нео узнаёт правду о реальности.",
                    Director = wachowski,
                    Genre = scifi,
                    IsTestData = true,
                    Review = new Review
                    {
                        Rating = 9.0,
                        Comment = "Классика фантастики.",
                        DateWatched = new DateTime(2024, 5, 20)
                    }
                }
            };

            var existingTitles = new HashSet<string>(
                GetAllMovies().Select(m => m.Title),
                StringComparer.OrdinalIgnoreCase);

            foreach (var movie in testMovies)
            {
                if (!existingTitles.Contains(movie.Title))
                    AddMovie(movie);
            }
        }

        private static int? EnsureGenre(SqliteConnection connection, SqliteTransaction? transaction, Genre? genre)
        {
            if (genre == null || string.IsNullOrWhiteSpace(genre.Name))
                return null;

            var name = genre.Name.Trim();
            using (var select = new SqliteCommand("SELECT Id FROM Genres WHERE Name = @Name COLLATE NOCASE;", connection))
            {
                if (transaction != null)
                    select.Transaction = transaction;
                select.Parameters.AddWithValue("@Name", name);
                var existing = select.ExecuteScalar();
                if (existing != null)
                {
                    genre.Id = Convert.ToInt32(existing);
                    genre.Name = name;
                    return genre.Id;
                }
            }

            using (var insert = new SqliteCommand(
                "INSERT INTO Genres (Name) VALUES (@Name); SELECT last_insert_rowid();",
                connection))
            {
                if (transaction != null)
                    insert.Transaction = transaction;
                insert.Parameters.AddWithValue("@Name", name);
                genre.Id = Convert.ToInt32(insert.ExecuteScalar());
                genre.Name = name;
                return genre.Id;
            }
        }

        private static int? EnsureDirector(SqliteConnection connection, SqliteTransaction? transaction, Director? director)
        {
            if (director == null || string.IsNullOrWhiteSpace(director.FullName))
                return null;

            var fullName = director.FullName.Trim();
            using (var select = new SqliteCommand(
                "SELECT Id FROM Directors WHERE FullName = @FullName COLLATE NOCASE;",
                connection))
            {
                if (transaction != null)
                    select.Transaction = transaction;
                select.Parameters.AddWithValue("@FullName", fullName);
                var existing = select.ExecuteScalar();
                if (existing != null)
                {
                    director.Id = Convert.ToInt32(existing);
                    director.FullName = fullName;
                    return director.Id;
                }
            }

            using (var insert = new SqliteCommand(
                "INSERT INTO Directors (FullName) VALUES (@FullName); SELECT last_insert_rowid();",
                connection))
            {
                if (transaction != null)
                    insert.Transaction = transaction;
                insert.Parameters.AddWithValue("@FullName", fullName);
                director.Id = Convert.ToInt32(insert.ExecuteScalar());
                director.FullName = fullName;
                return director.Id;
            }
        }
    }
}
