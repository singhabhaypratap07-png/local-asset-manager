using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Data.Sqlite;
using System.Diagnostics;

try {
    Console.WriteLine("--- Starting Asset Manager ---");
    
    var builder = WebApplication.CreateBuilder(args);
    var app = builder.Build();

    app.UseDefaultFiles();
    app.UseStaticFiles();

    // Database Setup
    string connectionString = "Data Source=assets.db";
    Console.WriteLine("Checking Database...");
    using (var connection = new SqliteConnection(connectionString)) {
        connection.Open();
        var cmd = connection.CreateCommand();
        cmd.CommandText = "CREATE TABLE IF NOT EXISTS assets (id TEXT PRIMARY KEY, name TEXT, serial TEXT);";
        cmd.ExecuteNonQuery();
    }
    Console.WriteLine("Database is OK.");

    app.MapGet("/api/data", () => new { message = "Success! Server is running." });

    Console.WriteLine("Server starting on http://127.0.0.1:8080");
    
    // Auto-open browser
    _ = Task.Run(() => { 
        Thread.Sleep(2000); 
        try {
            Process.Start(new ProcessStartInfo("http://127.0.0.1:8080") { UseShellExecute = true }); 
        } catch { }
    });

    app.Run("http://127.0.0.1:8080");

} catch (Exception ex) {
    Console.WriteLine("\n!!! CRITICAL ERROR !!!");
    Console.WriteLine(ex.ToString());
    Console.WriteLine("\nPress ENTER to close this window...");
    Console.ReadLine(); // यह विंडो को बंद होने से रोकेगा
}
