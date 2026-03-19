// AssetManager.cs
// C# 5, .NET Framework 4.x, WinForms + System.Data.SQLite
// ZERO-INSTALL .xlsx importer/exporter (Zip/XML) — no ACE OLEDB, no Office.
// Dates: dd-MM-yyyy
// Features: Login (PBKDF2), RBAC, User Management, Audit Logs, Assets (dynamic tabs), Import XLSX/CSV (Admin), Reports CSV/XLSX/PDF (pretty table).

using System;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;      // for ZipArchive (.xlsx is ZIP)
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace AssetManager
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try { Db.Initialize(); }
            catch (Exception ex)
            {
                MessageBox.Show("Database init failed:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            User loggedIn = null;
            using (var login = new LoginForm())
            {
                if (login.ShowDialog() == DialogResult.OK) loggedIn = login.LoggedInUser;
                else return;
            }

            string selectedUnit = null;
            using (var uf = new UnitSelectForm(loggedIn))
            {
                if (uf.ShowDialog() == DialogResult.OK) selectedUnit = uf.SelectedUnit;
                else return;
            }

            Application.Run(new MainForm(loggedIn, selectedUnit));
        }
    }

    // ===== Models =====
    public class User
    {
        public long Id;
        public string Username;
        public string Password;   // runtime only (not stored)
        public string Unit;       // ALL/Sugar/Distillery
        public string Role;       // Admin/User
        public bool CanEdit;
        public bool CanDel;
        public bool CanRepo;
    }

    public class AssetRow
    {
        public long Id;
        public string Unit;
        public string Category;
        public string SerialNo;
        public string Make;
        public string Model;
        public string Location;
        public string AssetNo;
        public string InvoiceDate;   // dd-MM-yyyy
        public string InvoiceNumber;
        public string PONumber;
        public string WarrantyUpTo;  // dd-MM-yyyy
        public string ExtraJson;
        public string CreatedBy;
    }

    // ===== DB + Security + Utilities =====
    public static class Db
    {
        public static readonly string AppDir = AppDomain.CurrentDomain.BaseDirectory;
        public static readonly string DbPath = Path.Combine(AppDir, "AssetManager.db");
        public static readonly string ConnStr = "Data Source=" + DbPath + ";Version=3;foreign keys=true;";

        // Password policy (tweak as you like)
        public static class PasswordPolicy
        {
            public static int MinLength = 8;
            public static bool RequireUpper = true;
            public static bool RequireLower = true;
            public static bool RequireDigit = true;
            public static bool RequireSpecial = true;

            public static bool Validate(string username, string password, out string message)
            {
                message = "";
                if (string.IsNullOrEmpty(password) || password.Length < MinLength)
                { message = "Password must be at least " + MinLength + " characters."; return false; }
                if (RequireUpper && !password.Any(char.IsUpper)) { message = "Password must contain at least one uppercase letter."; return false; }
                if (RequireLower && !password.Any(char.IsLower)) { message = "Password must contain at least one lowercase letter."; return false; }
                if (RequireDigit && !password.Any(char.IsDigit)) { message = "Password must contain at least one digit."; return false; }
                if (RequireSpecial && !password.Any(ch => "!@#$%^&*()_-+=[]{}|;:'\",.<>?/`~\\".Contains(ch))) { message = "Password must contain at least one special character."; return false; }
                if (!string.IsNullOrEmpty(username) && password.IndexOf(username, StringComparison.OrdinalIgnoreCase) >= 0)
                { message = "Password must not contain the username."; return false; }
                return true;
            }
        }

        public static void Initialize()
        {
            if (!File.Exists(DbPath)) SQLiteConnection.CreateFile(DbPath);

            using (var con = new SQLiteConnection(ConnStr))
            {
                con.Open();
                using (var cmd = con.CreateCommand())
                {
                    cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS Users(
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Username TEXT UNIQUE NOT NULL,
    Password TEXT NOT NULL,
    Unit TEXT NOT NULL DEFAULT 'ALL',
    Role TEXT NOT NULL DEFAULT 'User',
    CanEdit INTEGER NOT NULL DEFAULT 0,
    CanDel  INTEGER NOT NULL DEFAULT 0,
    CanRepo INTEGER NOT NULL DEFAULT 0,
    PwdHash TEXT,
    PwdSalt TEXT,
    PwdIter INTEGER DEFAULT 10000
);
CREATE TABLE IF NOT EXISTS Assets(
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Unit TEXT NOT NULL,
    Category TEXT NOT NULL,
    SerialNo TEXT,
    Make TEXT,
    Model TEXT,
    Location TEXT,
    AssetNo TEXT,
    InvoiceDate TEXT,     -- dd-MM-yyyy
    InvoiceNumber TEXT,
    PONumber TEXT,
    WarrantyUpTo TEXT,    -- dd-MM-yyyy
    ExtraJson TEXT,
    CreatedBy TEXT
);
CREATE INDEX IF NOT EXISTS IX_Assets_UnitCat ON Assets(Unit, Category);
CREATE INDEX IF NOT EXISTS IX_Assets_SN ON Assets(SerialNo);

CREATE TABLE IF NOT EXISTS ActivityLogs(
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Activity TEXT,
    Username TEXT,
    At TEXT
);";
                    cmd.ExecuteNonQuery();
                }

                // Users columns migration (safe-guard)
                try { using (var m = con.CreateCommand()) { m.CommandText = "ALTER TABLE Users ADD COLUMN PwdHash TEXT"; m.ExecuteNonQuery(); } } catch { }
                try { using (var m = con.CreateCommand()) { m.CommandText = "ALTER TABLE Users ADD COLUMN PwdSalt TEXT"; m.ExecuteNonQuery(); } } catch { }
                try { using (var m = con.CreateCommand()) { m.CommandText = "ALTER TABLE Users ADD COLUMN PwdIter INTEGER DEFAULT 10000"; m.ExecuteNonQuery(); } } catch { }

                // Seed default admin
                using (var cmd = con.CreateCommand())
                {
                    cmd.CommandText = "SELECT COUNT(*) FROM Users WHERE Username='admin'";
                    var cnt = Convert.ToInt32(cmd.ExecuteScalar());
                    if (cnt == 0)
                    {
                        // Seed with hashed password "123"
                        string s, h; int it; CreatePasswordHash("123", out s, out h, out it);
                        using (var ins = con.CreateCommand())
                        {
                            ins.CommandText = @"INSERT INTO Users(Username,Password,Unit,Role,CanEdit,CanDel,CanRepo,PwdHash,PwdSalt,PwdIter)
                                                VALUES('admin','', 'ALL','Admin',1,1,1,@h,@s,@i)";
                            ins.Parameters.AddWithValue("@h", h);
                            ins.Parameters.AddWithValue("@s", s);
                            ins.Parameters.AddWithValue("@i", it);
                            ins.ExecuteNonQuery();
                        }
                    }
                }
            }

            // Daily DB backup
            try
            {
                string backupName = "Backup_" + DateTime.Now.ToString("yyyyMMdd") + ".db";
                string backupPath = Path.Combine(AppDir, backupName);
                if (!File.Exists(backupPath)) File.Copy(DbPath, backupPath, false);
            }
            catch { }
        }

        // ---- Audit ----
        public static void Log(string activity, string username)
        {
            try
            {
                using (var con = new SQLiteConnection(ConnStr))
                using (var cmd = con.CreateCommand())
                {
                    con.Open();
                    cmd.CommandText = "INSERT INTO ActivityLogs(Activity,Username,At) VALUES(@a,@u,@t)";
                    cmd.Parameters.AddWithValue("@a", activity ?? "");
                    cmd.Parameters.AddWithValue("@u", username ?? "");
                    cmd.Parameters.AddWithValue("@t", DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"));
                    cmd.ExecuteNonQuery();
                }
            }
            catch { /* non-fatal */ }
        }

        public static DataTable GetLogs(DateTime? from, DateTime? to, string contains)
        {
            using (var con = new SQLiteConnection(ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                var sb = new StringBuilder("SELECT Id,Activity,Username,At FROM ActivityLogs WHERE 1=1");
                if (from.HasValue) { sb.Append(" AND date(substr(At,7,4)||'-'||substr(At,4,2)||'-'||substr(At,1,2)) >= date(@f)"); cmd.Parameters.AddWithValue("@f", from.Value.ToString("yyyy-MM-dd")); }
                if (to.HasValue)   { sb.Append(" AND date(substr(At,7,4)||'-'||substr(At,4,2)||'-'||substr(At,1,2)) <= date(@t)"); cmd.Parameters.AddWithValue("@t", to.Value.ToString("yyyy-MM-dd")); }
                if (!string.IsNullOrWhiteSpace(contains)) { sb.Append(" AND Activity LIKE @c"); cmd.Parameters.AddWithValue("@c", "%" + contains + "%"); }
                sb.Append(" ORDER BY Id DESC");
                cmd.CommandText = sb.ToString();
                using (var r = cmd.ExecuteReader())
                {
                    var dt = new DataTable(); dt.Load(r); return dt;
                }
            }
        }

        // ---- Password hashing (PBKDF2) ----
        public static void CreatePasswordHash(string password, out string saltB64, out string hashB64, out int iter)
        {
            iter = 10000; // iterations
            using (var rng = new RNGCryptoServiceProvider())
            {
                byte[] salt = new byte[16];
                rng.GetBytes(salt);
                var pbkdf2 = new Rfc2898DeriveBytes(password ?? "", salt, iter);
                var hash = pbkdf2.GetBytes(32); // 256-bit
                saltB64 = Convert.ToBase64String(salt);
                hashB64 = Convert.ToBase64String(hash);
            }
        }

        public static bool VerifyPassword(string password, string saltB64, string hashB64, int iter)
        {
            if (string.IsNullOrEmpty(saltB64) || string.IsNullOrEmpty(hashB64)) return false;
            byte[] salt = Convert.FromBase64String(saltB64);
            var pbkdf2 = new Rfc2898DeriveBytes(password ?? "", salt, iter > 0 ? iter : 10000);
            var hash = pbkdf2.GetBytes(32);
            var given = Convert.ToBase64String(hash);
            return string.Equals(given, hashB64, StringComparison.Ordinal);
        }

        public static User Authenticate(string username, string password)
        {
            using (var con = new SQLiteConnection(ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                cmd.CommandText = "SELECT Id,Username,Password,Unit,Role,CanEdit,CanDel,CanRepo,PwdHash,PwdSalt,PwdIter FROM Users WHERE Username=@u";
                cmd.Parameters.AddWithValue("@u", username);
                using (var r = cmd.ExecuteReader())
                {
                    if (r.Read())
                    {
                        long id = r.GetInt64(0);
                        string dbUser = r.GetString(1);
                        string plain = r.IsDBNull(2) ? "" : r.GetString(2);
                        string unit = r.GetString(3);
                        string role = r.GetString(4);
                        bool ce = r.GetInt32(5) == 1, cd = r.GetInt32(6) == 1, cr = r.GetInt32(7) == 1;
                        string pHash = r.IsDBNull(8) ? "" : r.GetString(8);
                        string pSalt = r.IsDBNull(9) ? "" : r.GetString(9);
                        int pIter = r.IsDBNull(10) ? 10000 : r.GetInt32(10);

                        bool ok = false;

                        if (!string.IsNullOrEmpty(pHash) && !string.IsNullOrEmpty(pSalt))
                        {
                            ok = VerifyPassword(password, pSalt, pHash, pIter);
                        }
                        else
                        {
                            // legacy plain
                            if (string.Equals(plain, password))
                            {
                                ok = true;
                                // migrate to hash
                                try
                                {
                                    string s, h; int it;
                                    CreatePasswordHash(password, out s, out h, out it);
                                    using (var up = con.CreateCommand())
                                    {
                                        up.CommandText = "UPDATE Users SET PwdHash=@h, PwdSalt=@s, PwdIter=@i, Password='' WHERE Id=@id";
                                        up.Parameters.AddWithValue("@h", h);
                                        up.Parameters.AddWithValue("@s", s);
                                        up.Parameters.AddWithValue("@i", it);
                                        up.Parameters.AddWithValue("@id", id);
                                        up.ExecuteNonQuery();
                                    }
                                } catch { }
                            }
                        }

                        if (ok)
                        {
                            return new User { Id = id, Username = dbUser, Password = "", Unit = unit, Role = role, CanEdit = ce, CanDel = cd, CanRepo = cr };
                        }
                    }
                }
            }
            return null;
        }

        // ---- Assets ----
        public static DataTable GetExpiringWithinDays(string unit, int days)
        {
            var dt = new DataTable();
            dt.Columns.Add("Unit"); dt.Columns.Add("Category"); dt.Columns.Add("SerialNo");
            dt.Columns.Add("Model"); dt.Columns.Add("WarrantyUpTo");

            var now = DateTime.Now.Date;
            var limit = now.AddDays(days);

            using (var con = new SQLiteConnection(ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                cmd.CommandText = string.Equals(unit, "ALL", StringComparison.OrdinalIgnoreCase)
                    ? "SELECT Unit,Category,SerialNo,Model,WarrantyUpTo FROM Assets"
                    : "SELECT Unit,Category,SerialNo,Model,WarrantyUpTo FROM Assets WHERE Unit=@u";
                if (!string.Equals(unit, "ALL", StringComparison.OrdinalIgnoreCase))
                    cmd.Parameters.AddWithValue("@u", unit);

                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        var w = r.IsDBNull(4) ? "" : r.GetString(4);
                        DateTime wd;
                        if (TryParseDateAny(w, out wd))
                        {
                            if (wd >= now && wd <= limit)
                            {
                                var row = dt.NewRow();
                                row["Unit"] = r.GetString(0);
                                row["Category"] = r.GetString(1);
                                row["SerialNo"] = r.IsDBNull(2) ? "" : r.GetString(2);
                                row["Model"] = r.IsDBNull(3) ? "" : r.GetString(3);
                                row["WarrantyUpTo"] = wd.ToString("dd-MM-yyyy");
                                dt.Rows.Add(row);
                            }
                        }
                    }
                }
            }
            return dt;
        }

        public static DataTable GetAssetsByUnitAndCategory(string unit, string category)
        {
            using (var con = new SQLiteConnection(ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                cmd.CommandText = "SELECT Id,SerialNo,Make,Model,Location,AssetNo,InvoiceDate,InvoiceNumber,PONumber,WarrantyUpTo,ExtraJson FROM Assets WHERE Unit=@u AND Category=@c ORDER BY Id DESC";
                cmd.Parameters.AddWithValue("@u", unit);
                cmd.Parameters.AddWithValue("@c", category);
                using (var r = cmd.ExecuteReader())
                {
                    var dt = new DataTable();
                    dt.Load(r);
                    return dt;
                }
            }
        }

        public static string[] GetCategoriesForUnit(string unit)
        {
            using (var con = new SQLiteConnection(ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                cmd.CommandText = "SELECT DISTINCT Category FROM Assets WHERE Unit=@u ORDER BY Category";
                cmd.Parameters.AddWithValue("@u", unit);
                var list = new System.Collections.Generic.List<string>();
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read()) list.Add(r.GetString(0));
                }
                return list.ToArray();
            }
        }

        public static void UpsertAsset(AssetRow a)
        {
            using (var con = new SQLiteConnection(ConnStr))
            {
                con.Open();
                using (var tx = con.BeginTransaction())
                {
                    using (var ins = con.CreateCommand())
                    {
                        ins.CommandText = @"
INSERT INTO Assets(Unit,Category,SerialNo,Make,Model,Location,AssetNo,InvoiceDate,InvoiceNumber,PONumber,WarrantyUpTo,ExtraJson,CreatedBy)
VALUES(@u,@cat,@sn,@make,@model,@loc,@asset,@invdt,@invno,@po,@wr,@x,@by)";
                        ins.Parameters.AddWithValue("@u", a.Unit ?? "");
                        ins.Parameters.AddWithValue("@cat", a.Category ?? "");
                        ins.Parameters.AddWithValue("@sn", a.SerialNo ?? "");
                        ins.Parameters.AddWithValue("@make", a.Make ?? "");
                        ins.Parameters.AddWithValue("@model", a.Model ?? "");
                        ins.Parameters.AddWithValue("@loc", a.Location ?? "");
                        ins.Parameters.AddWithValue("@asset", a.AssetNo ?? "");
                        ins.Parameters.AddWithValue("@invdt", a.InvoiceDate ?? "");
                        ins.Parameters.AddWithValue("@invno", a.InvoiceNumber ?? "");
                        ins.Parameters.AddWithValue("@po", a.PONumber ?? "");
                        ins.Parameters.AddWithValue("@wr", a.WarrantyUpTo ?? "");
                        ins.Parameters.AddWithValue("@x", a.ExtraJson ?? "");
                        ins.Parameters.AddWithValue("@by", a.CreatedBy ?? "");
                        ins.ExecuteNonQuery();
                    }
                    Log("UPSERT Asset: " + (a.Category ?? "") + " SN=" + (a.SerialNo ?? ""), a.CreatedBy);
                    tx.Commit();
                }
            }
        }

        public static bool DeleteAsset(long id, string byUser)
        {
            using (var con = new SQLiteConnection(ConnStr))
            {
                con.Open();
                int affected = 0;
                using (var cmd = con.CreateCommand())
                {
                    cmd.CommandText = "DELETE FROM Assets WHERE Id=@id";
                    cmd.Parameters.AddWithValue("@id", id);
                    affected = cmd.ExecuteNonQuery();
                }
                if (affected > 0)
                {
                    Log("DELETE Asset Id=" + id, byUser);
                }
                return affected > 0;
            }
        }

        public static void DataTableToCsv(DataTable dt, string path)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                sb.Append(dt.Columns[i].ColumnName);
                if (i < dt.Columns.Count - 1) sb.Append(",");
            }
            sb.AppendLine();
            foreach (DataRow row in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    var v = row[i] == DBNull.Value ? "" : Convert.ToString(row[i]);
                    if (v != null && (v.Contains(",") || v.Contains("\"")))
                        v = "\"" + v.Replace("\"", "\"\"") + "\"";
                    sb.Append(v);
                    if (i < dt.Columns.Count - 1) sb.Append(",");
                }
                sb.AppendLine();
            }
            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        // ---- Users: CRUD & guards ----
        public static DataTable GetAllUsers()
        {
            using (var con = new SQLiteConnection(ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                cmd.CommandText = "SELECT Id, Username, Unit, Role, CanEdit, CanDel, CanRepo FROM Users ORDER BY Username";
                using (var r = cmd.ExecuteReader())
                {
                    var dt = new DataTable();
                    dt.Load(r);
                    return dt;
                }
            }
        }

        public static bool UsernameExists(string username, long excludeId = 0)
        {
            using (var con = new SQLiteConnection(ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                if (excludeId > 0)
                {
                    cmd.CommandText = "SELECT COUNT(*) FROM Users WHERE Username=@u AND Id<>@id";
                    cmd.Parameters.AddWithValue("@u", username);
                    cmd.Parameters.AddWithValue("@id", excludeId);
                }
                else
                {
                    cmd.CommandText = "SELECT COUNT(*) FROM Users WHERE Username=@u";
                    cmd.Parameters.AddWithValue("@u", username);
                }
                return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
            }
        }

        public static int AdminCount()
        {
            using (var con = new SQLiteConnection(ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                cmd.CommandText = "SELECT COUNT(*) FROM Users WHERE Role='Admin'";
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

        public static long InsertUser(User u)
        {
            using (var con = new SQLiteConnection(ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                string s, h; int it; CreatePasswordHash(u.Password ?? "", out s, out h, out it);
                cmd.CommandText = @"INSERT INTO Users(Username,Password,Unit,Role,CanEdit,CanDel,CanRepo,PwdHash,PwdSalt,PwdIter)
                                    VALUES(@un,'',@unit,@role,@e,@d,@r,@h,@s,@i);
                                    SELECT last_insert_rowid();";
                cmd.Parameters.AddWithValue("@un", u.Username ?? "");
                cmd.Parameters.AddWithValue("@unit", string.IsNullOrEmpty(u.Unit) ? "ALL" : u.Unit);
                cmd.Parameters.AddWithValue("@role", string.IsNullOrEmpty(u.Role) ? "User" : u.Role);
                cmd.Parameters.AddWithValue("@e", u.CanEdit ? 1 : 0);
                cmd.Parameters.AddWithValue("@d", u.CanDel ? 1 : 0);
                cmd.Parameters.AddWithValue("@r", u.CanRepo ? 1 : 0);
                cmd.Parameters.AddWithValue("@h", h);
                cmd.Parameters.AddWithValue("@s", s);
                cmd.Parameters.AddWithValue("@i", it);
                var id = Convert.ToInt64(cmd.ExecuteScalar());
                return id;
            }
        }

        public static void UpdateUser(User u)
        {
            using (var con = new SQLiteConnection(ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                cmd.CommandText = @"UPDATE Users SET Username=@un, Unit=@unit, Role=@role, 
                                    CanEdit=@e, CanDel=@d, CanRepo=@r WHERE Id=@id";
                cmd.Parameters.AddWithValue("@un", u.Username ?? "");
                cmd.Parameters.AddWithValue("@unit", string.IsNullOrEmpty(u.Unit) ? "ALL" : u.Unit);
                cmd.Parameters.AddWithValue("@role", string.IsNullOrEmpty(u.Role) ? "User" : u.Role);
                cmd.Parameters.AddWithValue("@e", u.CanEdit ? 1 : 0);
                cmd.Parameters.AddWithValue("@d", u.CanDel ? 1 : 0);
                cmd.Parameters.AddWithValue("@r", u.CanRepo ? 1 : 0);
                cmd.Parameters.AddWithValue("@id", u.Id);
                cmd.ExecuteNonQuery();
            }
        }

        public static void UpdatePassword(long id, string newPass)
        {
            using (var con = new SQLiteConnection(ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                string s, h; int it; CreatePasswordHash(newPass ?? "", out s, out h, out it);
                cmd.CommandText = "UPDATE Users SET Password='', PwdHash=@h, PwdSalt=@s, PwdIter=@i WHERE Id=@id";
                cmd.Parameters.AddWithValue("@h", h);
                cmd.Parameters.AddWithValue("@s", s);
                cmd.Parameters.AddWithValue("@i", it);
                cmd.Parameters.AddWithValue("@id", id);
                cmd.ExecuteNonQuery();
            }
        }

        public static bool DeleteUser(long id)
        {
            using (var con = new SQLiteConnection(ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                cmd.CommandText = "DELETE FROM Users WHERE Id=@id";
                cmd.Parameters.AddWithValue("@id", id);
                return cmd.ExecuteNonQuery() > 0;
            }
        }

        // ---- Date helpers ----
        public static bool TryParseDateAny(string input, out DateTime dt)
        {
            dt = DateTime.MinValue;
            if (string.IsNullOrWhiteSpace(input)) return false;

            // handle range: "22-03-2024 to 19-06-2025" => take end date
            var s = input.Trim();
            var toIdx = s.IndexOf("to", StringComparison.OrdinalIgnoreCase);
            if (toIdx > 0)
            {
                var right = s.Substring(toIdx + 2).Trim(new[] { ' ', '-', '\t' });
                s = right;
            }

            string[] fmts = new[]
            {
                "dd-MM-yyyy","d-M-yyyy","dd/MM/yyyy","M/d/yyyy","MM/dd/yyyy","yyyy-MM-dd","dd.MM.yyyy",
                "MM-dd-yyyy","M-d-yyyy","d/M/yyyy","dd MMM yyyy"
            };
            return DateTime.TryParseExact(s, fmts, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out dt)
                || DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt);
        }

        public static string ToDdMMyyyy(DateTime dt) => dt.ToString("dd-MM-yyyy");
    }

    // ===== Forms =====
    public class LoginForm : Form
    {
        private TextBox txtU = new TextBox();
        private TextBox txtP = new TextBox();
        private Button btnLogin = new Button();
        private Label lblInfo = new Label();
        public User LoggedInUser { get; private set; }

        public LoginForm()
        {
            this.Text = "Login";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.ClientSize = new Size(320, 180);

            var lblU = new Label { Text = "Username:", Location = new Point(20, 30), AutoSize = true };
            var lblP = new Label { Text = "Password:", Location = new Point(20, 70), AutoSize = true };

            txtU.Location = new Point(110, 28); txtU.Width = 170;
            txtP.Location = new Point(110, 68); txtP.Width = 170; txtP.PasswordChar = '*';

            btnLogin.Text = "Login"; btnLogin.Location = new Point(110, 110); btnLogin.Width = 100;
            btnLogin.Click += BtnLogin_Click;

            lblInfo.Text = "Default: admin / 123";
            lblInfo.ForeColor = Color.Gray; lblInfo.Location = new Point(110, 140); lblInfo.AutoSize = true;

            this.Controls.AddRange(new Control[] { lblU, txtU, lblP, txtP, btnLogin, lblInfo });
            this.AcceptButton = btnLogin;
        }

        void BtnLogin_Click(object sender, EventArgs e)
        {
            var u = txtU.Text.Trim();
            var p = txtP.Text.Trim();
            if (u.Length == 0 || p.Length == 0) { MessageBox.Show("Please enter username and password."); return; }
            var auth = Db.Authenticate(u, p);
            if (auth != null)
            {
                Db.Log("LOGIN SUCCESS", u);
                this.LoggedInUser = auth;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                Db.Log("LOGIN FAIL", u);
                MessageBox.Show("Wrong credentials.");
            }
        }
    }

    public class UnitSelectForm : Form
    {
        private ComboBox cmb = new ComboBox();
        private Button btn = new Button();
        private User user;
        public string SelectedUnit { get; private set; }

        public UnitSelectForm(User u)
        {
            user = u;
            this.Text = "Select Unit";
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.ClientSize = new Size(300, 120);

            var lbl = new Label { Text = "Unit:", Location = new Point(20, 25), AutoSize = true };
            cmb.Location = new Point(70, 22); cmb.Width = 200;
            cmb.Items.Add("Sugar"); cmb.Items.Add("Distillery");

            if (!string.Equals(u.Unit, "ALL", StringComparison.OrdinalIgnoreCase))
            { cmb.SelectedItem = u.Unit; cmb.Enabled = false; }
            else { cmb.SelectedIndex = 0; }

            btn.Text = "OK"; btn.Location = new Point(100, 60);
            btn.Click += (s, e) =>
            {
                if (cmb.SelectedItem == null) { MessageBox.Show("Select a unit."); return; }
                SelectedUnit = cmb.SelectedItem.ToString(); this.DialogResult = DialogResult.OK; this.Close();
            };

            this.Controls.AddRange(new Control[] { lbl, cmb, btn });
        }
    }

    public class MainForm : Form
    {
        private User currentUser;
        private string currentUnit;

        private Label lblAlerts = new Label();
        private TabControl tabs = new TabControl();

        // Assets editor
        private TextBox txtSN = new TextBox();
        private TextBox txtMake = new TextBox();
        private TextBox txtModel = new TextBox();
        private TextBox txtLoc = new TextBox();
        private TextBox txtAssetNo = new TextBox();
        private DateTimePicker dtInv = new DateTimePicker();
        private TextBox txtInvNo = new TextBox();
        private TextBox txtPO = new TextBox();
        private DateTimePicker dtWr = new DateTimePicker();
        private Label lblCategory = new Label();

        private Button btnSave = new Button();
        private Button btnDel = new Button();
        private Button btnRepo = new Button();

        // Grids
        private System.Collections.Generic.Dictionary<string, DataGridView> grids = new System.Collections.Generic.Dictionary<string, DataGridView>();

        // Admin: User Management
        private DataGridView gvUsers;
        private TextBox txtUName, txtUPass;
        private ComboBox cmbUUnit, cmbURole;
        private CheckBox chkUEdit, chkUDel, chkURepo;
        private Button btnUNew, btnUSave, btnUDelete, btnUReset, btnURefresh;
        private long editingUserId = 0;

        // Admin: Audit
        private DataGridView gvLogs;
        private DateTimePicker dtFrom, dtTo;
        private TextBox txtContains;
        private Button btnLogFilter, btnLogExport;

        public MainForm(User u, string unit)
        {
            currentUser = u; currentUnit = unit;

            this.Text = "Asset Manager Pro (SQLite) — " + unit;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(1120, 720);

            lblAlerts.Text = "Checking expiries...";
            lblAlerts.Font = new Font("Arial", 10, FontStyle.Bold);
            lblAlerts.Location = new Point(20, 15); lblAlerts.AutoSize = true;

            tabs.Bounds = new Rectangle(20, 50, 1080, 500);

            BuildTabs();

            // Editor panel
            int y = 560;
            var lblCat = new Label { Text = "Category:", Location = new Point(20, y + 2), AutoSize = true };
            lblCategory.Text = "-"; lblCategory.Location = new Point(90, y + 2); lblCategory.AutoSize = true; lblCategory.Font = new Font("Segoe UI", 9, FontStyle.Bold);

            var lbl1 = new Label { Text = "Serial No:", Location = new Point(180, y + 2), AutoSize = true };
            txtSN.Location = new Point(250, y); txtSN.Width = 140;

            var lblMk = new Label { Text = "Make:", Location = new Point(400, y + 2), AutoSize = true };
            txtMake.Location = new Point(450, y); txtMake.Width = 120;

            var lblMd = new Label { Text = "Model:", Location = new Point(580, y + 2), AutoSize = true };
            txtModel.Location = new Point(630, y); txtModel.Width = 160;

            var lblLc = new Label { Text = "Location:", Location = new Point(800, y + 2), AutoSize = true };
            txtLoc.Location = new Point(865, y); txtLoc.Width = 200;

            var lblAs = new Label { Text = "Asset No:", Location = new Point(20, y + 35), AutoSize = true };
            txtAssetNo.Location = new Point(90, y + 32); txtAssetNo.Width = 140;

            var lblInv = new Label { Text = "Invoice Date:", Location = new Point(240, y + 35), AutoSize = true };
            dtInv.Location = new Point(320, y + 32); dtInv.Width = 120; dtInv.Format = DateTimePickerFormat.Custom; dtInv.CustomFormat = "dd-MM-yyyy";

            var lblIno = new Label { Text = "Invoice No:", Location = new Point(450, y + 35), AutoSize = true };
            txtInvNo.Location = new Point(520, y + 32); txtInvNo.Width = 120;

            var lblPO = new Label { Text = "P.O. No:", Location = new Point(650, y + 35), AutoSize = true };
            txtPO.Location = new Point(700, y + 32); txtPO.Width = 120;

            var lblWr = new Label { Text = "Warranty Upto:", Location = new Point(830, y + 35), AutoSize = true };
            dtWr.Location = new Point(920, y + 32); dtWr.Width = 120; dtWr.Format = DateTimePickerFormat.Custom; dtWr.CustomFormat = "dd-MM-yyyy";

            btnSave.Text = "Save/Insert"; btnSave.Location = new Point(250, 660); btnSave.Width = 120; btnSave.BackColor = Color.LightGreen;
            btnDel.Text = "Delete Selected"; btnDel.Location = new Point(380, 660); btnDel.Width = 140; btnDel.BackColor = Color.Salmon;
            btnRepo.Text = "Export Report"; btnRepo.Location = new Point(530, 660); btnRepo.Width = 160; btnRepo.BackColor = Color.LightBlue;

            // RBAC enablement
            btnSave.Enabled = currentUser.CanEdit || string.Equals(currentUser.Role, "Admin", StringComparison.OrdinalIgnoreCase);
            btnDel.Enabled = currentUser.CanDel || string.Equals(currentUser.Role, "Admin", StringComparison.OrdinalIgnoreCase);
            btnRepo.Enabled = currentUser.CanRepo || string.Equals(currentUser.Role, "Admin", StringComparison.OrdinalIgnoreCase);

            btnSave.Click += (s, e) => SaveRow();
            btnDel.Click += (s, e) => DeleteSelected();
            btnRepo.Click += BtnRepo_Click;

            this.Shown += (s, e) => ShowExpiryAlert();

            this.Controls.AddRange(new Control[] {
                lblAlerts, tabs,
                lblCat, lblCategory,
                lbl1, txtSN, lblMk, txtMake, lblMd, txtModel, lblLc, txtLoc,
                lblAs, txtAssetNo, lblInv, dtInv, lblIno, txtInvNo, lblPO, txtPO, lblWr, dtWr,
                btnSave, btnDel, btnRepo
            });
        }

        private void ShowExpiryAlert()
        {
            try
            {
                var dt = Db.GetExpiringWithinDays(currentUnit, 30);
                if (dt.Rows.Count > 0)
                {
                    lblAlerts.Text = "⚠️ " + dt.Rows.Count + " expiry within 30 days";
                    lblAlerts.ForeColor = Color.Red;
                }
                else
                {
                    lblAlerts.Text = "No upcoming expiries within 30 days.";
                    lblAlerts.ForeColor = Color.DarkGreen;
                }
            }
            catch
            {
                lblAlerts.Text = "Expiry check failed.";
                lblAlerts.ForeColor = Color.DarkRed;
            }
        }

        private void BuildTabs()
        {
            tabs.TabPages.Clear();
            grids.Clear();

            // Asset tabs
            var cats = Db.GetCategoriesForUnit(currentUnit);
            if (cats.Length == 0)
            {
                var tp = new TabPage("No Data");
                var lbl = new Label { Text = "No data imported yet.", AutoSize = true, Location = new Point(20, 20) };
                tp.Controls.Add(lbl);
                tabs.TabPages.Add(tp);
            }
            else
            {
                foreach (var c in cats)
                {
                    var tp = new TabPage(c);
                    var grid = new DataGridView();
                    grid.AllowUserToAddRows = false;
                    grid.ReadOnly = true;
                    grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    grid.MultiSelect = false;
                    grid.Dock = DockStyle.Fill;

                    grid.DataSource = Db.GetAssetsByUnitAndCategory(currentUnit, c);
                    grid.CellClick += (s, e) =>
                    {
                        lblCategory.Text = c;
                        var gv = (DataGridView)s;
                        if (e.RowIndex >= 0 && e.RowIndex < gv.Rows.Count)
                        {
                            var row = gv.Rows[e.RowIndex];
                            txtSN.Text = Safe(row, "SerialNo");
                            txtMake.Text = Safe(row, "Make");
                            txtModel.Text = Safe(row, "Model");
                            txtLoc.Text = Safe(row, "Location");
                            txtAssetNo.Text = Safe(row, "AssetNo");
                            DateTime d1, d2;
                            if (Db.TryParseDateAny(Safe(row, "InvoiceDate"), out d1)) dtInv.Value = d1;
                            if (Db.TryParseDateAny(Safe(row, "WarrantyUpTo"), out d2)) dtWr.Value = d2;
                            txtInvNo.Text = Safe(row, "InvoiceNumber");
                            txtPO.Text = Safe(row, "PONumber");
                        }
                    };

                    tp.Controls.Add(grid);
                    tabs.TabPages.Add(tp);
                    grids[c] = grid;
                }
            }

            // Admin tab (Import + User Management)
            if (string.Equals(currentUser.Role, "Admin", StringComparison.OrdinalIgnoreCase))
            {
                var adminTab = new TabPage("👥 Admin / Import & Users");

                // Import area
                var lblA = new Label { Text = "Bulk Import (Zero-Install XLSX):", AutoSize = true, Location = new Point(20, 20) };
                var btnImpSugar = new Button { Text = "Import Sugar XLSX", Location = new Point(20, 50), Width = 180 };
                btnImpSugar.Click += (s, e) => ImportXlsxForUnit("Sugar", Path.Combine(Db.AppDir, "Sugar_System Details_24-25.xlsx"));

                var btnImpAlco = new Button { Text = "Import Distillery XLSX", Location = new Point(210, 50), Width = 200 };
                btnImpAlco.Click += (s, e) => ImportXlsxForUnit("Distillery", Path.Combine(Db.AppDir, "Alco_System Details_24-25 .xlsx"));

                var sep = new Label { Text = "– OR – CSV import (folder with CSV per sheet)", AutoSize = true, Location = new Point(20, 90) };
                var btnCsvSugar = new Button { Text = "Import Sugar CSV Folder", Location = new Point(20, 120), Width = 200 };
                btnCsvSugar.Click += (s, e) => ImportCsvFolder("Sugar");
                var btnCsvAlco = new Button { Text = "Import Distillery CSV Folder", Location = new Point(230, 120), Width = 220 };
                btnCsvAlco.Click += (s, e) => ImportCsvFolder("Distillery");

                var btnExportUsers = new Button { Text = "Export Users CSV", Location = new Point(20, 170), Width = 160 };
                btnExportUsers.Click += (s, e) =>
                {
                    try
                    {
                        using (var con = new SQLiteConnection(Db.ConnStr))
                        using (var cmd = con.CreateCommand())
                        {
                            con.Open();
                            cmd.CommandText = "SELECT Username,Unit,Role,CanEdit,CanDel,CanRepo FROM Users";
                            using (var r = cmd.ExecuteReader())
                            {
                                var dt = new DataTable(); dt.Load(r);
                                var path = Path.Combine(Db.AppDir, "Users_Master.csv");
                                Db.DataTableToCsv(dt, path);
                                MessageBox.Show("Exported: " + path);
                                Db.Log("USERS EXPORT CSV: " + path, currentUser.Username);
                            }
                        }
                    }
                    catch (Exception ex2) { MessageBox.Show("Export failed: " + ex2.Message); }
                };

                // User management
                var lblUM = new Label { Text = "User Management", AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(20, 210) };

                gvUsers = new DataGridView
                {
                    Location = new Point(20, 240),
                    Size = new Size(600, 220),
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    MultiSelect = false
                };
                gvUsers.CellClick += (s, e) =>
                {
                    if (e.RowIndex >= 0 && e.RowIndex < gvUsers.Rows.Count)
                    {
                        var row = gvUsers.Rows[e.RowIndex];
                        BindUserFormFromRow(row);
                    }
                };

                int fx = 640, fy = 240, sp = 28, w1 = 260;

                var lU = new Label { Text = "Username:", Location = new Point(fx, fy + 2), AutoSize = true };
                txtUName = new TextBox { Location = new Point(fx + 100, fy), Width = w1 };

                var lP = new Label { Text = "Password:", Location = new Point(fx, fy + sp + 2), AutoSize = true };
                txtUPass = new TextBox { Location = new Point(fx + 100, fy + sp), Width = w1, PasswordChar = '*' };

                var lUn = new Label { Text = "Unit:", Location = new Point(fx, fy + sp * 2 + 2), AutoSize = true };
                cmbUUnit = new ComboBox { Location = new Point(fx + 100, fy + sp * 2), Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
                cmbUUnit.Items.AddRange(new object[] { "ALL", "Sugar", "Distillery" });
                cmbUUnit.SelectedIndex = 0;

                var lR = new Label { Text = "Role:", Location = new Point(fx, fy + sp * 3 + 2), AutoSize = true };
                cmbURole = new ComboBox { Location = new Point(fx + 100, fy + sp * 3), Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
                cmbURole.Items.AddRange(new object[] { "Admin", "User" });
                cmbURole.SelectedIndex = 1;

                var lPerm = new Label { Text = "Permissions:", Location = new Point(fx, fy + sp * 4 + 2), AutoSize = true };
                chkUEdit = new CheckBox { Text = "CanEdit", Location = new Point(fx + 100, fy + sp * 4 - 2), AutoSize = true, Checked = true };
                chkUDel = new CheckBox { Text = "CanDel", Location = new Point(fx + 180, fy + sp * 4 - 2), AutoSize = true, Checked = false };
                chkURepo = new CheckBox { Text = "CanRepo", Location = new Point(fx + 260, fy + sp * 4 - 2), AutoSize = true, Checked = true };

                btnUNew = new Button { Text = "New", Location = new Point(fx, fy + sp * 6), Width = 80 };
                btnUSave = new Button { Text = "Save", Location = new Point(fx + 90, fy + sp * 6), Width = 80 };
                btnUDelete = new Button { Text = "Delete", Location = new Point(fx + 180, fy + sp * 6), Width = 80, BackColor = Color.Salmon };
                btnUReset = new Button { Text = "Reset/Set Password", Location = new Point(fx + 270, fy + sp * 6), Width = 150 };
                btnURefresh = new Button { Text = "Refresh", Location = new Point(fx + 430, fy + sp * 6), Width = 80 };

                btnUNew.Click += (s, e) => ClearUserForm();
                btnURefresh.Click += (s, e) => { LoadUsersGrid(); ClearUserForm(); };

                btnUSave.Click += (s, e) =>
                {
                    try
                    {
                        string msg;
                        if (!ValidateUserForm(out msg)) { MessageBox.Show(msg); return; }

                        var uName = txtUName.Text.Trim();
                        var pw = txtUPass.Text; // optional on update
                        var unit = Convert.ToString(cmbUUnit.SelectedItem);
                        var role = Convert.ToString(cmbURole.SelectedItem);
                        bool e1 = chkUEdit.Checked, d1 = chkUDel.Checked, r1 = chkURepo.Checked;

                        if (editingUserId == 0)
                        {
                            // Password policy for new user
                            string pm; if (!Db.PasswordPolicy.Validate(uName, pw, out pm)) { MessageBox.Show(pm); return; }

                            var uo = new User
                            {
                                Username = uName, Password = pw ?? "", Unit = unit, Role = role,
                                CanEdit = e1, CanDel = d1, CanRepo = r1
                            };
                            var id = Db.InsertUser(uo);
                            MessageBox.Show("User created (Id=" + id + ").");
                            Db.Log("USER CREATE: " + uName + " (Role=" + role + ", Unit=" + unit + ")", currentUser.Username);
                        }
                        else
                        {
                            // Last Admin demotion guard
                            if (WillBreakLastAdminOnUpdate(role, editingUserId))
                            {
                                MessageBox.Show("This is the last Admin. Create another Admin before demoting.");
                                return;
                            }

                            var uo = new User
                            {
                                Id = editingUserId, Username = uName, Unit = unit, Role = role,
                                CanEdit = e1, CanDel = d1, CanRepo = r1
                            };
                            Db.UpdateUser(uo);

                            if (!string.IsNullOrEmpty(pw))
                            {
                                string pm; if (!Db.PasswordPolicy.Validate(uName, pw, out pm)) { MessageBox.Show(pm); return; }
                                Db.UpdatePassword(editingUserId, pw);
                                Db.Log("USER PASSWORD RESET: Id=" + editingUserId, currentUser.Username);
                                txtUPass.Text = "";
                            }

                            MessageBox.Show("User updated.");
                            Db.Log("USER UPDATE: " + uName + " (Role=" + role + ", Unit=" + unit + ")", currentUser.Username);
                        }

                        LoadUsersGrid();
                        ClearUserForm();
                    }
                    catch (Exception exx)
                    {
                        MessageBox.Show("Save failed: " + exx.Message);
                    }
                };

                btnUDelete.Click += (s, e) =>
                {
                    if (editingUserId == 0) { MessageBox.Show("Select a user."); return; }
                    if (IsSelf(editingUserId)) { MessageBox.Show("You cannot delete yourself."); return; }

                    // Last admin guard
                    var dtAll = Db.GetAllUsers();
                    var row = dtAll.AsEnumerable().FirstOrDefault(r => Convert.ToInt64(r["Id"]) == editingUserId);
                    if (row != null && string.Equals(Convert.ToString(row["Role"]), "Admin", StringComparison.OrdinalIgnoreCase))
                    {
                        if (Db.AdminCount() <= 1)
                        {
                            MessageBox.Show("Cannot delete last Admin. Create another Admin first.");
                            return;
                        }
                    }

                    if (MessageBox.Show("Delete selected user?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        try
                        {
                            var ok = Db.DeleteUser(editingUserId);
                            MessageBox.Show(ok ? "Deleted." : "Not found.");
                            Db.Log("USER DELETE: Id=" + editingUserId, currentUser.Username);
                            LoadUsersGrid();
                            ClearUserForm();
                        }
                        catch (Exception exx)
                        {
                            MessageBox.Show("Delete failed: " + exx.Message);
                        }
                    }
                };

                btnUReset.Click += (s, e) =>
                {
                    if (editingUserId == 0) { MessageBox.Show("Select a user."); return; }
                    var newPw = txtUPass.Text;
                    if (string.IsNullOrEmpty(newPw))
                    {
                        MessageBox.Show("Enter new password in Password box, then click Reset/Set Password.");
                        return;
                    }
                    try
                    {
                        string uName = txtUName.Text.Trim();
                        string pm; if (!Db.PasswordPolicy.Validate(uName, newPw, out pm)) { MessageBox.Show(pm); return; }
                        Db.UpdatePassword(editingUserId, newPw);
                        Db.Log("USER PASSWORD RESET: Id=" + editingUserId, currentUser.Username);
                        MessageBox.Show("Password updated.");
                        txtUPass.Text = "";
                    }
                    catch (Exception exx)
                    {
                        MessageBox.Show("Password update failed: " + exx.Message);
                    }
                };

                adminTab.Controls.AddRange(new Control[] {
                    lblA, btnImpSugar, btnImpAlco, sep, btnCsvSugar, btnCsvAlco, btnExportUsers,
                    lblUM, gvUsers,
                    lU, txtUName, lP, txtUPass, lUn, cmbUUnit, lR, cmbURole, lPerm, chkUEdit, chkUDel, chkURepo,
                    btnUNew, btnUSave, btnUDelete, btnUReset, btnURefresh
                });

                tabs.TabPages.Add(adminTab);
                LoadUsersGrid(); ClearUserForm();

                // ---- Audit Logs tab
                var logTab = new TabPage("📝 Audit Logs");
                dtFrom = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd-MM-yyyy", Location = new Point(20, 20), Width = 120, Value = DateTime.Now.AddDays(-30) };
                dtTo   = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd-MM-yyyy", Location = new Point(150, 20), Width = 120, Value = DateTime.Now };
                txtContains = new TextBox { Location = new Point(280, 20), Width = 200 };
                btnLogFilter = new Button { Text = "Filter", Location = new Point(490, 18), Width = 80 };
                btnLogExport = new Button { Text = "Export CSV", Location = new Point(580, 18), Width = 100 };

                gvLogs = new DataGridView
                {
                    Location = new Point(20, 60),
                    Size = new Size(1030, 380),
                    ReadOnly = true,
                    AllowUserToAddRows = false,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    MultiSelect = false,
                    Dock = DockStyle.None
                };

                Action refreshLogs = () =>
                {
                    gvLogs.DataSource = Db.GetLogs(dtFrom.Value.Date, dtTo.Value.Date, txtContains.Text.Trim());
                };
                btnLogFilter.Click += (s, e) => refreshLogs();
                btnLogExport.Click += (s, e) =>
                {
                    try
                    {
                        var dt = Db.GetLogs(dtFrom.Value.Date, dtTo.Value.Date, txtContains.Text.Trim());
                        var path = Path.Combine(Db.AppDir, "Audit_Logs.csv");
                        Db.DataTableToCsv(dt, path);
                        Db.Log("AUDIT EXPORT CSV: " + path, currentUser.Username);
                        MessageBox.Show("Exported: " + path);
                    }
                    catch (Exception ex2) { MessageBox.Show("Export failed: " + ex2.Message); }
                };

                logTab.Controls.AddRange(new Control[] { dtFrom, dtTo, txtContains, btnLogFilter, btnLogExport, gvLogs });
                tabs.TabPages.Add(logTab);
                refreshLogs();
            }
        }

        private static string Safe(DataGridViewRow row, string col)
        {
            if (!row.DataGridView.Columns.Contains(col)) return "";
            var v = row.Cells[col].Value;
            return v == null ? "" : Convert.ToString(v);
        }

        private bool ValidateUserForm(out string msg)
        {
            msg = "";
            var un = txtUName.Text.Trim();
            if (string.IsNullOrEmpty(un)) { msg = "Username required."; return false; }
            if (cmbUUnit.SelectedItem == null) { msg = "Unit required."; return false; }
            if (cmbURole.SelectedItem == null) { msg = "Role required."; return false; }
            if (Db.UsernameExists(un, editingUserId)) { msg = "Username already exists."; return false; }
            return true;
        }

        private bool WillBreakLastAdminOnUpdate(string newRole, long targetUserId)
        {
            if (newRole == null) return false;
            newRole = newRole.Trim();
            if (!string.Equals(newRole, "User", StringComparison.OrdinalIgnoreCase)) return false;

            var dt = Db.GetAllUsers();
            var row = dt.AsEnumerable().FirstOrDefault(r => Convert.ToInt64(r["Id"]) == targetUserId);
            if (row == null) return false;
            var curRole = Convert.ToString(row["Role"]);
            if (!string.Equals(curRole, "Admin", StringComparison.OrdinalIgnoreCase)) return false;

            return Db.AdminCount() <= 1;
        }

        private bool IsSelf(long userId) => (currentUser != null && currentUser.Id == userId);

        // ===== Assets: Save/Delete/Report =====
        private void SaveRow()
        {
            if (!(currentUser.CanEdit || string.Equals(currentUser.Role, "Admin", StringComparison.OrdinalIgnoreCase)))
            { MessageBox.Show("No Access"); return; }

            if (string.IsNullOrEmpty(lblCategory.Text) || lblCategory.Text == "-")
            { MessageBox.Show("Select a category tab row first."); return; }

            try
            {
                var a = new AssetRow();
                a.Unit = currentUnit;
                a.Category = lblCategory.Text;
                a.SerialNo = txtSN.Text.Trim();
                a.Make = txtMake.Text.Trim();
                a.Model = txtModel.Text.Trim();
                a.Location = txtLoc.Text.Trim();
                a.AssetNo = txtAssetNo.Text.Trim();
                a.InvoiceDate = Db.ToDdMMyyyy(dtInv.Value.Date);
                a.InvoiceNumber = txtInvNo.Text.Trim();
                a.PONumber = txtPO.Text.Trim();
                a.WarrantyUpTo = Db.ToDdMMyyyy(dtWr.Value.Date);
                a.ExtraJson = "";
                a.CreatedBy = currentUser.Username;

                Db.UpsertAsset(a);
                MessageBox.Show("Saved.");
                RefreshCurrentTab();
                ShowExpiryAlert();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Save failed: " + ex.Message);
            }
        }

        private void DeleteSelected()
        {
            if (!(currentUser.CanDel || string.Equals(currentUser.Role, "Admin", StringComparison.OrdinalIgnoreCase)))
            { MessageBox.Show("No Access"); return; }

            var tp = tabs.SelectedTab;
            if (tp == null) return;
            DataGridView grid;
            if (!grids.TryGetValue(tp.Text, out grid)) return;
            if (grid.SelectedRows.Count == 0) { MessageBox.Show("Select a row."); return; }

            var idObj = grid.SelectedRows[0].Cells["Id"].Value;
            long id;
            if (idObj == null || !long.TryParse(Convert.ToString(idObj), out id))
            { MessageBox.Show("Invalid row."); return; }

            if (MessageBox.Show("Delete selected row?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    var ok = Db.DeleteAsset(id, currentUser.Username);
                    MessageBox.Show(ok ? "Deleted." : "Not found.");
                    RefreshCurrentTab();
                }
                catch (Exception ex) { MessageBox.Show("Delete failed: " + ex.Message); }
            }
        }

        private void RefreshCurrentTab()
        {
            var tp = tabs.SelectedTab;
            if (tp == null) return;
            DataGridView g;
            if (grids.TryGetValue(tp.Text, out g))
            {
                g.DataSource = Db.GetAssetsByUnitAndCategory(currentUnit, tp.Text);
            }
        }

        // ===== Importers (Admin-only with audit) =====
        private void ImportXlsxForUnit(string unit, string xlsxPath)
        {
            if (!string.Equals(currentUser.Role, "Admin", StringComparison.OrdinalIgnoreCase))
            { MessageBox.Show("Only Admin can import."); return; }

            if (!File.Exists(xlsxPath))
            { MessageBox.Show("File not found:\n" + xlsxPath); return; }

            try
            {
                Db.Log("IMPORT XLSX START: " + unit + " -> " + Path.GetFileName(xlsxPath), currentUser.Username);
                var book = XlsxReader.ReadWorkbook(xlsxPath);
                foreach (var sheet in book.Sheets)
                {
                    var dt = sheet.Table;
                    if (dt == null || dt.Columns.Count == 0) continue;
                    ImportGenericTable(unit, sheet.Name, dt);
                }
                MessageBox.Show("XLSX imported for " + unit + ". Rebuilding tabs…");
                BuildTabs();
                Db.Log("IMPORT XLSX FINISH: " + unit, currentUser.Username);
            }
            catch (Exception ex)
            {
                Db.Log("IMPORT XLSX FAIL: " + unit + " :: " + ex.Message, currentUser.Username);
                MessageBox.Show("Import failed: " + ex.Message);
            }
        }

        private void ImportCsvFolder(string unit)
        {
            if (!string.Equals(currentUser.Role, "Admin", StringComparison.OrdinalIgnoreCase))
            { MessageBox.Show("Only Admin can import."); return; }

            using (var fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Select folder containing CSV files (one per sheet/category)";
                if (fbd.ShowDialog() != DialogResult.OK) return;

                var files = Directory.GetFiles(fbd.SelectedPath, "*.csv");
                if (files.Length == 0) { MessageBox.Show("No CSV files found."); return; }

                Db.Log("IMPORT CSV START: " + unit + " -> " + fbd.SelectedPath, currentUser.Username);
                foreach (var f in files)
                {
                    try
                    {
                        var dt = CsvToDataTable(f);
                        var category = Path.GetFileNameWithoutExtension(f);
                        ImportGenericTable(unit, category, dt);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("CSV import failed for: " + f + "\n" + ex.Message);
                    }
                }

                MessageBox.Show("CSV imported for " + unit + ". Rebuilding tabs…");
                BuildTabs();
                Db.Log("IMPORT CSV FINISH: " + unit, currentUser.Username);
            }
        }

        private static DataTable CsvToDataTable(string path)
        {
            var dt = new DataTable();
            using (var sr = new StreamReader(path))
            {
                string header = sr.ReadLine();
                if (header == null) return dt;
                var cols = SplitCsv(header);
                foreach (var c in cols) dt.Columns.Add(c);

                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    var cells = SplitCsv(line);
                    var row = dt.NewRow();
                    for (int i = 0; i < dt.Columns.Count && i < cells.Length; i++)
                        row[i] = cells[i];
                    dt.Rows.Add(row);
                }
            }
            return dt;
        }

        private static string[] SplitCsv(string line)
        {
            var list = new System.Collections.Generic.List<string>();
            var sb = new StringBuilder();
            bool inQ = false;
            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];
                if (inQ)
                {
                    if (ch == '"' && i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                    else if (ch == '"') inQ = false;
                    else sb.Append(ch);
                }
                else
                {
                    if (ch == '"') inQ = true;
                    else if (ch == ',') { list.Add(sb.ToString()); sb.Clear(); }
                    else sb.Append(ch);
                }
            }
            list.Add(sb.ToString());
            return list.ToArray();
        }

        private void ImportGenericTable(string unit, string category, DataTable dt)
        {
            if (dt == null || dt.Columns.Count == 0) return;

            var normCols = dt.Columns.Cast<DataColumn>().Select(c => new { Col = c, Key = Normalize(c.ColumnName) }).ToList();
            string sn = null, make = null, model = null, loc = null, asset = null, invDt = null, invNo = null, po = null, wr = null;

            Func<string[], string> pick = (cands) =>
            {
                foreach (var c in cands)
                {
                    var k = Normalize(c);
                    var hit = normCols.FirstOrDefault(z => z.Key == k);
                    if (hit != null) return hit.Col.ColumnName;
                }
                return null;
            };

            sn = pick(new[] { "Serial No","Printer S.N.","Sys. SN (Service Tag)","SN (Service Tag)","SN.","HDD S.N. (Service Tag)","Dongal S.N. (Service Tag)","Serial Number","TFT SN.","Sys. SN","SN","HHT Num" });
            make = pick(new[] { "Make","Mack" });
            model = pick(new[] { "Model","Sys. Model","Printer Model","Device Name" });
            loc = pick(new[] { "Location" });
            asset = pick(new[] { "Asset no.","Asset No.","Desktop Asset no.","Laptop Asset no.","HHT Asset no." });
            invDt = pick(new[] { "Invoice date","INV. DATE","Start Date" });
            invNo = pick(new[] { "Invoice Number" });
            po = pick(new[] { "P.O. Number","P.O. Num.","P.O. Num" });
            wr = pick(new[] { "Warranty up to","End Date","Battery Instt. Date" });

            foreach (DataRow row in dt.Rows)
            {
                bool allEmpty = true;
                foreach (var c in dt.Columns.Cast<DataColumn>())
                {
                    if (!string.IsNullOrWhiteSpace(Convert.ToString(row[c]))) { allEmpty = false; break; }
                }
                if (allEmpty) continue;

                var a = new AssetRow();
                a.Unit = unit;
                a.Category = category;
                a.SerialNo = sn == null ? "" : Convert.ToString(row[sn]).Trim();
                a.Make = make == null ? "" : Convert.ToString(row[make]).Trim();
                a.Model = model == null ? "" : Convert.ToString(row[model]).Trim();
                a.Location = loc == null ? "" : Convert.ToString(row[loc]).Trim();
                a.AssetNo = asset == null ? "" : Convert.ToString(row[asset]).Trim();
                a.InvoiceNumber = invNo == null ? "" : Convert.ToString(row[invNo]).Trim();
                a.PONumber = po == null ? "" : Convert.ToString(row[po]).Trim();

                a.InvoiceDate = "";
                if (invDt != null)
                {
                    var s = Convert.ToString(row[invDt]).Trim();
                    DateTime d; if (Db.TryParseDateAny(s, out d)) a.InvoiceDate = Db.ToDdMMyyyy(d);
                }

                a.WarrantyUpTo = "";
                if (wr != null)
                {
                    var s = Convert.ToString(row[wr]).Trim();
                    DateTime d; if (Db.TryParseDateAny(s, out d)) a.WarrantyUpTo = Db.ToDdMMyyyy(d);
                }

                var jb = new StringBuilder();
                jb.Append("{");
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    var key = dt.Columns[i].ColumnName.Replace("\"", "");
                    var val = row[i] == DBNull.Value ? "" : Convert.ToString(row[i]).Replace("\"", "");
                    if (i > 0) jb.Append(",");
                    jb.Append("\"").Append(key).Append("\":\"").Append(val).Append("\"");
                }
                jb.Append("}");
                a.ExtraJson = jb.ToString();

                a.CreatedBy = currentUser != null ? currentUser.Username : "import";
                Db.UpsertAsset(a);
            }
        }

        private static string Normalize(string s)
        {
            if (s == null) return "";
            var sb = new StringBuilder();
            foreach (var ch in s.ToLowerInvariant())
            {
                if (char.IsLetterOrDigit(ch)) sb.Append(ch);
                else if (char.IsWhiteSpace(ch)) { }
            }
            return sb.ToString();
        }

        // ===== Report Export (CSV/XLSX/PDF) =====
        private void BtnRepo_Click(object sender, EventArgs e)
        {
            if (!(currentUser.CanRepo || string.Equals(currentUser.Role, "Admin", StringComparison.OrdinalIgnoreCase)))
            { MessageBox.Show("No Access"); return; }

            try
            {
                string unitScope = (string.Equals(currentUser.Unit, "ALL", StringComparison.OrdinalIgnoreCase)) ? currentUnit : currentUser.Unit;
                var dt = GetMasterData(unitScope);

                using (var sfd = new SaveFileDialog())
                {
                    sfd.Title = "Export Master Report";
                    sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx|PDF Document (*.pdf)|*.pdf|CSV (Excel Openable) (*.csv)|*.csv";
                    sfd.FileName = "Assets_Master_" + unitScope;
                    if (sfd.ShowDialog() != DialogResult.OK) return;

                    string path = sfd.FileName;
                    if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportXlsx(dt, path, "Master");
                        Db.Log("REPORT EXPORT XLSX: " + Path.GetFileName(path), currentUser.Username);
                        MessageBox.Show("Excel exported: " + path);
                    }
                    else if (path.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportPdfPrettyTable(dt, path, "Assets Master - " + unitScope);
                        Db.Log("REPORT EXPORT PDF: " + Path.GetFileName(path), currentUser.Username);
                        MessageBox.Show("PDF exported: " + path);
                    }
                    else
                    {
                        Db.DataTableToCsv(dt, path);
                        Db.Log("REPORT EXPORT CSV: " + Path.GetFileName(path), currentUser.Username);
                        MessageBox.Show("CSV exported: " + path);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Export failed: " + ex.Message);
            }
        }

        private DataTable GetMasterData(string unitScope)
        {
            using (var con = new SQLiteConnection(Db.ConnStr))
            using (var cmd = con.CreateCommand())
            {
                con.Open();
                cmd.CommandText = "SELECT Unit,Category,SerialNo,Make,Model,Location,AssetNo,InvoiceDate,InvoiceNumber,PONumber,WarrantyUpTo FROM Assets WHERE Unit=@u ORDER BY Category,Id DESC";
                cmd.Parameters.AddWithValue("@u", unitScope);
                using (var r = cmd.ExecuteReader())
                {
                    var dt = new DataTable(); dt.Load(r); return dt;
                }
            }
        }

        // === Zero-install XLSX (inlineStr) ===
        private void ExportXlsx(DataTable dt, string path, string sheetName)
        {
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                var eCT = zip.CreateEntry("[Content_Types].xml");
                using (var s = new StreamWriter(eCT.Open(), new UTF8Encoding(false)))
                {
                    s.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
<Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
<Default Extension=""xml"" ContentType=""application/xml""/>
<Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
<Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
</Types>");
                }

                var erels = zip.CreateEntry("_rels/.rels");
                using (var s = new StreamWriter(erels.Open(), new UTF8Encoding(false)))
                {
                    s.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""xl/workbook.xml""/>
</Relationships>");
                }

                var eWb = zip.CreateEntry("xl/workbook.xml");
                using (var s = new StreamWriter(eWb.Open(), new UTF8Encoding(false)))
                {
                    s.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <sheets>
    <sheet name=""" + XmlEscape(sheetName) + @""" sheetId=""1"" r:id=""rId1""/>
  </sheets>
</workbook>");
                }

                var eWrels = zip.CreateEntry("xl/_rels/workbook.xml.rels");
                using (var s = new StreamWriter(eWrels.Open(), new UTF8Encoding(false)))
                {
                    s.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
<Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""worksheets/sheet1.xml""/>
</Relationships>");
                }

                var eSheet = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using (var s = new StreamWriter(eSheet.Open(), new UTF8Encoding(false)))
                {
                    s.Write(@"<?xml version=""1.0"" encoding=""UTF-8""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <sheetData>");
                    s.Write("\n    <row>");
                    foreach (DataColumn c in dt.Columns)
                        s.Write(@"<c t=""inlineStr""><is><t>{0}</t></is></c>", XmlEscape(c.ColumnName));
                    s.Write("</row>");
                    foreach (DataRow r in dt.Rows)
                    {
                        s.Write("\n    <row>");
                        foreach (DataColumn c in dt.Columns)
                        {
                            var v = r[c] == DBNull.Value ? "" : Convert.ToString(r[c]);
                            s.Write(@"<c t=""inlineStr""><is><t>{0}</t></is></c>", XmlEscape(v));
                        }
                        s.Write("</row>");
                    }
                    s.Write(@"
  </sheetData>
</worksheet>");
                }
            }
        }
        private static string XmlEscape(string v)
        {
            if (string.IsNullOrEmpty(v)) return "";
            return v.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }

        // === Pretty PDF table (grid lines + header band + pagination) ===
        private void ExportPdfPrettyTable(DataTable dt, string path, string title)
        {
            // Simple PDF generator with Helvetica & Helvetica-Bold (standard 14 fonts)
            // Page metrics
            int pageW = 595, pageH = 842; // A4
            int marginL = 36, marginR = 36, marginT = 40, marginB = 40;
            int usableW = pageW - marginL - marginR;
            int yTop = pageH - marginT;

            // Fonts
            const string F_REG = "/F1", F_BOLD = "/F2";
            float fontSize = 9f, headerSize = 11f;

            // Compute column widths in pixels approx (character-based heuristic)
            var cols = dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray();
            int n = cols.Length;
            if (n == 0) { File.WriteAllBytes(path, new byte[] { }); return; }

            // measure widths based on content up to a limit
            int[] charW = new int[n];
            for (int i = 0; i < n; i++)
            {
                int w = Math.Max(cols[i].Length, 6);
                int sample = Math.Min(dt.Rows.Count, 200);
                for (int r = 0; r < sample; r++)
                {
                    var v = dt.Rows[r][i] == DBNull.Value ? "" : Convert.ToString(dt.Rows[r][i]);
                    w = Math.Max(w, Math.Min(v.Length, 40)); // cap to control width
                }
                charW[i] = w;
            }
            int totalChars = charW.Sum() + (n - 1) * 2; // 2-char gap
            double pxPerChar = usableW / (double)totalChars;

            int[] colPx = charW.Select(c => (int)Math.Max(40, Math.Round(c * pxPerChar))).ToArray();

            // Build lines with per-cell fitting
            string Fit(string s, int px, int approxPerChar)
            {
                if (string.IsNullOrEmpty(s)) return "";
                int maxChars = Math.Max(1, (int)Math.Floor(px / (double)approxPerChar));
                if (s.Length <= maxChars) return s;
                return s.Substring(0, Math.Max(1, maxChars - 1)) + "…";
            }
            int approxCharPix = 5; // heuristic for Helvetica 9pt

            // PDF assembly
            var xref = new System.Collections.Generic.List<long>();
            var buf = new MemoryStream();
            var sw = new StreamWriter(buf, new ASCIIEncoding()) { NewLine = "\n" };
            Action<string> W = (s) => sw.Write(s);

            W("%PDF-1.4\n");
            // 1: Catalog
            xref.Add(buf.Position); W("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
            // 2: Pages (we'll add page objects dynamically; for simplicity, single pages appended)
            xref.Add(buf.Position); W("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n");

            // 4: Fonts
            xref.Add(buf.Position); W("4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n");
            xref.Add(buf.Position); W("5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>\nendobj\n");

            // content stream (5 -> adjust indices)
            // We'll write one page per content stream for simplicity.
            // Create content
            var content = new MemoryStream();
            var cw = new StreamWriter(content, new ASCIIEncoding()) { NewLine = "\n" };
            Action<string> C = (s) => { cw.Write(s); cw.Write("\n"); };

            // helpers
            Action<float, float> MoveText = (x, y) => C(string.Format("{0} {1} Td", x.ToString("0.##", CultureInfo.InvariantCulture), y.ToString("0.##", CultureInfo.InvariantCulture)));
            Action<string> Tj = (text) => C("(" + EscapePdf(text) + ") Tj");
            Action<string, float> SetFont = (f, size) => C(f + " " + size.ToString("0.##", CultureInfo.InvariantCulture) + " Tf");
            Action<float> SetGrayFill = (g) => C(g.ToString("0.##", CultureInfo.InvariantCulture) + " g");
            Action<float> SetGrayStroke = (g) => C(g.ToString("0.##", CultureInfo.InvariantCulture) + " G");
            Action<float> SetLineWidth = (w) => C(w.ToString("0.##", CultureInfo.InvariantCulture) + " w");
            Action<float, float, float, float> Rect = (x, y, w, h) => C(string.Format("{0} {1} {2} {3} re", x.ToString("0.##", CultureInfo.InvariantCulture), y.ToString("0.##", CultureInfo.InvariantCulture), w.ToString("0.##", CultureInfo.InvariantCulture), h.ToString("0.##", CultureInfo.InvariantCulture)));
            Action Stroke = () => C("S");
            Action Fill = () => C("f");

            // layout
            int rowH = 16; // px
            int headerH = 22;
            int y = yTop;

            // New page function (we only produce one page here; for large tables we paginate)
            var pagesContent = new System.Collections.Generic.List<byte[]>();
            Action NewPage = () =>
            {
                cw.Flush();
                if (content.Length > 0) pagesContent.Add(content.ToArray());
                content.Dispose();
                content = new MemoryStream();
                cw = new StreamWriter(content, new ASCIIEncoding()) { NewLine = "\n" };
                y = yTop;
                // header band
                SetGrayFill(0.95f); Rect(marginL, y - headerH, usableW, headerH); Fill();
                SetGrayFill(0f);
                C("BT"); SetFont(F_BOLD, headerSize); MoveText(marginL, y - 14); Tj(title); C("ET");
                y -= (headerH + 6);

                // table header background
                SetGrayFill(0.90f); Rect(marginL, y - rowH, usableW, rowH); Fill();
                SetGrayStroke(0.2f); SetLineWidth(0.5f); Rect(marginL, y - rowH, usableW, rowH); Stroke();

                // header text
                C("BT"); SetFont(F_BOLD, 9f); MoveText(0, 0);
                int x = marginL + 4;
                MoveText(x, y - 12);
                for (int i = 0; i < n; i++)
                {
                    string h = cols[i];
                    Tj(Fit(h, colPx[i] - 8, approxCharPix));
                    // move to next col
                    MoveText(colPx[i], 0);
                }
                C("ET");
                y -= rowH;
            };

            // compute rows per page
            int rowsPerPage = (pageH - marginB - (pageH - yTop) - headerH - rowH - 40) / rowH;
            if (rowsPerPage < 10) rowsPerPage = 10;

            int rowCount = dt.Rows.Count;
            int written = 0;
            int pageNum = 0;
            while (written < rowCount || (written == 0 && rowCount == 0))
            {
                pageNum++;
                NewPage();

                // Draw column vertical lines (grid later per row)
                // iterate rows for this page
                int take = Math.Min(rowsPerPage, rowCount - written);
                if (rowCount == 0) take = 0;

                // rows
                for (int r = 0; r < take; r++)
                {
                    // row box
                    SetGrayStroke(0.8f); SetLineWidth(0.3f);
                    Rect(marginL, y - rowH, usableW, rowH); Stroke();

                    // text
                    C("BT"); SetFont(F_REG, fontSize); MoveText(marginL + 4, y - 12);
                    for (int i = 0; i < n; i++)
                    {
                        var v = dt.Rows[written + r][i] == DBNull.Value ? "" : Convert.ToString(dt.Rows[written + r][i]);
                        Tj(Fit(v, colPx[i] - 8, approxCharPix));
                        MoveText(colPx[i], 0);
                    }
                    C("ET");

                    y -= rowH;
                }

                // footer page number
                C("BT"); SetFont(F_REG, 8f); MoveText(marginL, marginB); Tj("Page " + pageNum); C("ET");
                written += take;
            }

            // add last content buffer
            cw.Flush();
            pagesContent.Add(content.ToArray());
            content.Dispose();

            // Create content object(s)
            // For simplicity: single page in Pages tree was set to 3 0 R. We need to fix if >1 pages.
            // If >1, we will rebuild Pages object with all pages.
            int pageCount = pagesContent.Count;
            int firstPageObjId = 3;

            // Write each page's content object and page object
            var pageObjIds = new System.Collections.Generic.List<int>();
            var contentObjIds = new System.Collections.Generic.List<int>();

            for (int i = 0; i < pageCount; i++)
            {
                // content obj
                content = new MemoryStream(pagesContent[i]);
                xref.Add(buf.Position);
                int contentId = xref.Count; // 1-based
                contentObjIds.Add(contentId);
                W(contentId + " 0 obj\n<< /Length " + content.Length + " >>\nstream\n");
                content.Position = 0; content.CopyTo(buf);
                W("\nendstream\nendobj\n");

                // page obj
                xref.Add(buf.Position);
                int pageObjId = xref.Count;
                pageObjIds.Add(pageObjId);
                W(pageObjId + " 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + pageW + " " + pageH + "] ");
                W("/Resources << /Font << /F1 4 0 R /F2 5 0 R >> >> /Contents " + contentId + " 0 R >>\nendobj\n");
            }

            // Rebuild Pages (obj 2) now with proper Kids/Count
            long pagesOverridePos = buf.Position;
            xref[1] = pagesOverridePos; // replace obj 2 position
            W("2 0 obj\n<< /Type /Pages /Kids [");
            foreach (var pid in pageObjIds) W(pid + " 0 R ");
            W("] /Count " + pageCount + " >>\nendobj\n");

            // xref
            long xrefPos = buf.Position;
            W("xref\n0 " + (xref.Count + 1) + "\n");
            W("0000000000 65535 f \n");
            foreach (var p in xref) W(p.ToString("0000000000") + " 00000 n \n");
            W("trailer\n<< /Size " + (xref.Count + 1) + " /Root 1 0 R >>\nstartxref\n" + xrefPos + "\n%%EOF");

            sw.Flush();
            File.WriteAllBytes(path, buf.ToArray());
        }

        private static string EscapePdf(string s)
        {
            if (s == null) return "";
            return s.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)").Replace("\r", "").Replace("\n", "\\n");
        }

        // ======= XLSX Reader (zero-install) =======
        public class XlsxReader
        {
            public class SheetData { public string Name; public DataTable Table; }
            public class Workbook { public System.Collections.Generic.List<SheetData> Sheets = new System.Collections.Generic.List<SheetData>(); }

            public static Workbook ReadWorkbook(string path)
            {
                using (var fs = File.OpenRead(path))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    var sst = ReadSharedStrings(zip);
                    var rels = ReadWorkbookRelationships(zip);
                    var sheets = ReadWorkbookSheets(zip);

                    var wb = new Workbook();
                    foreach (var sh in sheets)
                    {
                        string partPath;
                        if (!rels.TryGetValue(sh.RelId, out partPath)) continue;
                        var full = "xl/" + partPath.TrimStart('/');
                        var entry = zip.GetEntry(full);
                        if (entry == null) continue;

                        using (var s = entry.Open())
                        {
                            var dt = ReadWorksheetToTable(s, sst);
                            if (dt != null && dt.Columns.Count > 0)
                                wb.Sheets.Add(new SheetData { Name = sh.Name, Table = dt });
                        }
                    }
                    return wb;
                }
            }

            private class SheetInfo { public string Name; public string RelId; }

            private static System.Collections.Generic.List<SheetInfo> ReadWorkbookSheets(ZipArchive zip)
            {
                var list = new System.Collections.Generic.List<SheetInfo>();
                var entry = zip.GetEntry("xl/workbook.xml");
                if (entry == null) return list;

                var doc = new XmlDocument();
                using (var s = entry.Open()) doc.Load(s);

                var nsm = new XmlNamespaceManager(doc.NameTable);
                nsm.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                nsm.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                var nodes = doc.SelectNodes("//d:sheets/d:sheet", nsm);
                foreach (XmlNode n in nodes)
                {
                    var name = n.Attributes["name"] != null ? n.Attributes["name"].Value : "Sheet";
                    var rid = n.Attributes["r:id"] != null ? n.Attributes["r:id"].Value : null;
                    if (!string.IsNullOrEmpty(rid)) list.Add(new SheetInfo { Name = name, RelId = rid });
                }
                return list;
            }

            private static System.Collections.Generic.Dictionary<string, string> ReadWorkbookRelationships(ZipArchive zip)
            {
                var map = new System.Collections.Generic.Dictionary<string, string>();
                var entry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                if (entry == null) return map;

                var doc = new XmlDocument();
                using (var s = entry.Open()) doc.Load(s);

                var nsm = new XmlNamespaceManager(doc.NameTable);
                nsm.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");

                var rels = doc.SelectNodes("//r:Relationship", nsm);
                foreach (XmlNode r in rels)
                {
                    var type = r.Attributes["Type"] != null ? r.Attributes["Type"].Value : "";
                    if (!type.EndsWith("/worksheet")) continue;
                    var id = r.Attributes["Id"] != null ? r.Attributes["Id"].Value : null;
                    var target = r.Attributes["Target"] != null ? r.Attributes["Target"].Value : null;
                    if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                        map[id] = target;
                }
                return map;
            }

            private static string[] ReadSharedStrings(ZipArchive zip)
            {
                var entry = zip.GetEntry("xl/sharedStrings.xml");
                if (entry == null) return null;

                var list = new System.Collections.Generic.List<string>();
                var doc = new XmlDocument();
                using (var s = entry.Open()) doc.Load(s);

                var nsm = new XmlNamespaceManager(doc.NameTable);
                nsm.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                var siNodes = doc.SelectNodes("//d:si", nsm);
                foreach (XmlNode si in siNodes)
                {
                    var tNodes = si.SelectNodes(".//d:t", nsm);
                    if (tNodes != null && tNodes.Count > 0)
                    {
                        var sb = new StringBuilder();
                        foreach (XmlNode t in tNodes) sb.Append(t.InnerText);
                        list.Add(sb.ToString());
                    }
                    else list.Add(si.InnerText);
                }
                return list.ToArray();
            }

            private static DataTable ReadWorksheetToTable(Stream sheetXmlStream, string[] sharedStrings)
            {
                var doc = new XmlDocument();
                doc.Load(sheetXmlStream);

                var nsm = new XmlNamespaceManager(doc.NameTable);
                nsm.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                var rows = doc.SelectNodes("//d:worksheet/d:sheetData/d:row", nsm);
                if (rows == null || rows.Count == 0) return new DataTable();

                string[] header = null;
                var dataRows = new System.Collections.Generic.List<string[]>();
                int maxCols = 0;

                foreach (XmlNode r in rows)
                {
                    var cells = RowToValues(r, sharedStrings);
                    bool allEmpty = true; foreach (var v in cells) if (!string.IsNullOrWhiteSpace(v)) { allEmpty = false; break; }
                    if (allEmpty) continue;

                    if (header == null) header = MakeHeaderUnique(cells);
                    else dataRows.Add(cells);

                    maxCols = Math.Max(maxCols, cells.Length);
                }

                if (header == null) return new DataTable();
                var dt = new DataTable();
                foreach (var h in header) dt.Columns.Add(string.IsNullOrWhiteSpace(h) ? "Column" + (dt.Columns.Count + 1) : h);
                foreach (var row in dataRows)
                {
                    var dr = dt.NewRow();
                    for (int i = 0; i < dt.Columns.Count && i < row.Length; i++) dr[i] = row[i];
                    dt.Rows.Add(dr);
                }
                return dt;
            }

            private static string[] MakeHeaderUnique(string[] cells)
            {
                var arr = new string[cells.Length];
                var used = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < cells.Length; i++)
                {
                    var h = string.IsNullOrWhiteSpace(cells[i]) ? "Column" + (i + 1) : cells[i].Trim();
                    var baseH = h; int n = 1;
                    while (used.Contains(h)) { n++; h = baseH + " (" + n + ")"; }
                    used.Add(h); arr[i] = h;
                }
                return arr;
            }

            private static string[] RowToValues(XmlNode rowNode, string[] sharedStrings)
            {
                var cells = rowNode.SelectNodes("./d:c", new XmlNamespaceManager(rowNode.OwnerDocument.NameTable) { { "d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main" } });
                if (cells == null || cells.Count == 0) return new string[0];

                int maxIdx = 0;
                var tmp = new System.Collections.Generic.List<(int idx, string val)>();
                foreach (XmlNode c in cells)
                {
                    string r = Attr(c, "r");
                    int colIdx = ColumnIndexFromA1(r);
                    string val = CellValue(c, sharedStrings);
                    tmp.Add((colIdx, val));
                    if (colIdx > maxIdx) maxIdx = colIdx;
                }
                var arr = new string[maxIdx];
                foreach (var t in tmp)
                {
                    int i = t.idx - 1;
                    if (i >= 0 && i < arr.Length) arr[i] = t.val;
                }
                for (int i = 0; i < arr.Length; i++) if (arr[i] == null) arr[i] = "";
                return arr;
            }

            private static string Attr(XmlNode n, string name)
            {
                var a = n.Attributes != null ? n.Attributes[name] : null;
                return a == null ? null : a.Value;
            }

            private static int ColumnIndexFromA1(string a1)
            {
                if (string.IsNullOrWhiteSpace(a1)) return 1;
                int i = 0; int col = 0;
                while (i < a1.Length && char.IsLetter(a1[i]))
                {
                    col = col * 26 + (char.ToUpperInvariant(a1[i]) - 'A' + 1);
                    i++;
                }
                return col == 0 ? 1 : col;
            }

            private static string CellValue(XmlNode c, string[] sharedStrings)
            {
                var t = Attr(c, "t");
                if (t == "s")
                {
                    var v = c.SelectSingleNode("./d:v", ns(c));
                    if (v == null) return "";
                    int idx;
                    if (int.TryParse(v.InnerText, out idx) && sharedStrings != null && idx >= 0 && idx < sharedStrings.Length)
                        return sharedStrings[idx];
                    return v.InnerText;
                }
                else if (t == "inlineStr")
                {
                    var tnode = c.SelectSingleNode("./d:is/d:t", ns(c));
                    return tnode == null ? "" : tnode.InnerText;
                }
                else
                {
                    var v = c.SelectSingleNode("./d:v", ns(c));
                    return v == null ? "" : v.InnerText;
                }
            }

            private static XmlNamespaceManager ns(XmlNode n)
            {
                var nsm = new XmlNamespaceManager(n.OwnerDocument.NameTable);
                nsm.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                return nsm;
            }
        }
    }
}
