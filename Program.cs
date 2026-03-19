using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Data.Sqlite;
using System.Diagnostics;

// यहाँ सुधार किया गया है: CreateBuilder(args) का इस्तेमाल करें
var builder = WebApplication.CreateBuilder(args);
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
_ = Task.Run(() => { 
    Thread.Sleep(2000); 
    try {
        Process.Start(new ProcessStartInfo("http://localhost:5000") { UseShellExecute = true }); 
    } catch {
        // अगर ब्राउज़र नहीं खुलता तो कोई बात नहीं, यूजर मैन्युअली खोल सकता है
    }
});

app.Run("http://localhost:5000");
