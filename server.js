const express = require("express");
const odbc = require("odbc");
const path = require("path");

const app = express();
const PORT = 3000;

// Update DB_PATH to point at your .mdb or .accdb file
const DB_PATH = process.env.DB_PATH || "C:\\path\\to\\your\\database.accdb";

const connectionString =
  `Driver={Microsoft Access Driver (*.mdb, *.accdb)};` +
  `Dbq=${DB_PATH};`;

app.use(express.static(path.join(__dirname)));

app.get("/test-connection", async (req, res) => {
  let connection;
  try {
    connection = await odbc.connect(connectionString);
    const result = await connection.query(
      "SELECT COUNT(*) AS table_count FROM MSysObjects WHERE Type=1 AND Flags=0"
    );
    const count = result[0]?.table_count ?? 0;
    res.json({ success: true, tableCount: count, dbPath: DB_PATH });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message, dbPath: DB_PATH });
  } finally {
    if (connection) await connection.close();
  }
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
  console.log(`Testing DB: ${DB_PATH}`);
});
