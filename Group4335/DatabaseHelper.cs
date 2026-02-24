using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

namespace Group4335
{
    public class DatabaseHelper
    {
        private string connectionString = "Data Source=orders.db;Version=3;";

        public void SaveOrdersToDatabase(List<Order> orders)
        {
            string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "orders.db");
            string connectionString = $"Data Source={dbPath};Version=3;";

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                string createTableQuery = @"
                    CREATE TABLE IF NOT EXISTS Orders (
                        Id INTEGER PRIMARY KEY,
                        OrderCode TEXT,
                        CreationDate DATETIME,
                        ClientCode TEXT,
                        Services TEXT,
                        Status TEXT
                    )";

                using (var command = new SQLiteCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }

                foreach (var order in orders)
                {
                    string insertQuery = @"
                        INSERT OR REPLACE INTO Orders (Id, OrderCode, CreationDate, ClientCode, Services, Status)
                        VALUES (@Id, @OrderCode, @CreationDate, @ClientCode, @Services, @Status)";

                    using (var command = new SQLiteCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@Id", order.Id);
                        command.Parameters.AddWithValue("@OrderCode", order.OrderCode ?? "");
                        command.Parameters.AddWithValue("@CreationDate", order.CreationDate);
                        command.Parameters.AddWithValue("@ClientCode", order.ClientCode ?? "");
                        command.Parameters.AddWithValue("@Services", order.Services ?? "");
                        command.Parameters.AddWithValue("@Status", order.Status ?? "");

                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        public List<Order> GetOrdersFromDatabase()
        {
            var orders = new List<Order>();
            string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "orders.db");
            string connectionString = $"Data Source={dbPath};Version=3;";

            if (!File.Exists(dbPath))
            {
                return orders;
            }

            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT * FROM Orders";

                using (var command = new SQLiteCommand(query, connection))
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        orders.Add(new Order
                        {
                            Id = Convert.ToInt32(reader["Id"]),
                            OrderCode = reader["OrderCode"]?.ToString() ?? "",
                            CreationDate = Convert.ToDateTime(reader["CreationDate"]),
                            ClientCode = reader["ClientCode"]?.ToString() ?? "",
                            Services = reader["Services"]?.ToString() ?? "",
                            Status = reader["Status"]?.ToString() ?? ""
                        });
                    }
                }
            }

            return orders;
        }
    }
}