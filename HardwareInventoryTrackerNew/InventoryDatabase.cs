using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

namespace HardwareInventoryTrackerNew
{
    public class InventoryDatabase : IDisposable
    {
        private readonly SQLiteConnection connection;

        public InventoryDatabase()
        {
            string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "inventory.db");
            connection = new SQLiteConnection($"Data Source={dbPath};Version=3;");
            connection.Open();
            SetupDatabase();
        }

        public void SetupDatabase()
        {
            string createInventoryTable = @"
                CREATE TABLE IF NOT EXISTS inventory (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    asset_tag TEXT,
                    description TEXT,
                    serial_number TEXT NOT NULL,
                    transfer_sheet TEXT,
                    notes TEXT,
                    date TEXT,
                    time TEXT,
                    location TEXT,
                    transferred_by TEXT,
                    received_by TEXT,
                    color TEXT
                )";
            string createKnownInventoryTable = @"
                CREATE TABLE IF NOT EXISTS known_inventory (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    asset_tag TEXT,
                    description TEXT,
                    serial_number TEXT
                )";
            string createSettingsTable = @"
                CREATE TABLE IF NOT EXISTS settings (
                    id INTEGER PRIMARY KEY,
                    theme TEXT
                )";

            ExecuteNonQuery(createInventoryTable, new Dictionary<string, object>());
            ExecuteNonQuery(createKnownInventoryTable, new Dictionary<string, object>());
            ExecuteNonQuery(createSettingsTable, new Dictionary<string, object>());
        }

        public List<object[]> ExecuteQuery(string query, Dictionary<string, object>? parameters = null)
        {
            using (var cmd = new SQLiteCommand(query, connection))
            {
                if (parameters != null)
                {
                    foreach (var param in parameters)
                        cmd.Parameters.AddWithValue(param.Key, param.Value);
                }

                using (var reader = cmd.ExecuteReader())
                {
                    var results = new List<object[]>();
                    while (reader.Read())
                    {
                        var row = new object[reader.FieldCount];
                        for (int i = 0; i < reader.FieldCount; i++)
                            row[i] = reader[i];
                        results.Add(row);
                    }
                    return results;
                }
            }
        }

        public void ExecuteNonQuery(string query, Dictionary<string, object>? parameters = null)
        {
            using (var cmd = new SQLiteCommand(query, connection))
            {
                if (parameters != null)
                {
                    foreach (var param in parameters)
                        cmd.Parameters.AddWithValue(param.Key, param.Value);
                }
                cmd.ExecuteNonQuery();
            }
        }

        public void BulkInsert(List<Dictionary<string, string>> entries)
        {
            using (var transaction = connection.BeginTransaction())
            {
                foreach (var entry in entries)
                {
                    string columns = string.Join(", ", entry.Keys);
                    string values = string.Join(", ", entry.Keys.Select(k => $"@{k}"));
                    string query = $"INSERT INTO inventory ({columns}) VALUES ({values})";
                    using (var cmd = new SQLiteCommand(query, connection))
                    {
                        foreach (var param in entry)
                            cmd.Parameters.AddWithValue($"@{param.Key}", param.Value);
                        cmd.ExecuteNonQuery();
                    }
                }
                transaction.Commit();
            }
        }

        public void Dispose()
        {
            connection?.Close();
            connection?.Dispose();
        }
    }
}