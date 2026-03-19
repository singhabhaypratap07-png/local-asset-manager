using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Data.Sqlite;
using System.Diagnostics;

var builder = WebApplication.Create();
var app = builder.Build();

app.UseDefaultFiles();
app.UseStaticFiles();

// Database Setup
string connectionString = "Data Source=assets.db";
using (var connection = new SqliteConnection(connectionString)) {
    connection.Open();
    var cmd = connection.CreateCommand();
    cmd.CommandText = "CREATE TABLE IF NOT EXISTS assets (id TEXT PRIMARY KEY, name TEXT, serial TEXT);";
    cmd.ExecuteNonQuery();
}

app.MapGet("/api/data", () => new { message = "Hello! Data is stored in assets.db on your PC." });

// Auto-open browser
Task.Run(() => { Thread.Sleep(2000); Process.Start(new ProcessStartInfo("http://localhost:5000") { UseShellExecute = true }); });

app.Run("http://localhost:5000");
