const odbc = require("odbc");

// Update these values to match your environment
const DB_PATH = process.env.DB_PATH || "C:\\path\\to\\your\\database.accdb";

const connectionString =
  `Driver={Microsoft Access Driver (*.mdb, *.accdb)};` +
  `Dbq=${DB_PATH};`;

async function testConnection() {
  console.log("Testing connection to Microsoft Access database...");
  console.log(`DB path: ${DB_PATH}\n`);

  let connection;
  try {
    connection = await odbc.connect(connectionString);
    console.log("SUCCESS: Connected to the database.");

    // Run a simple query to confirm the connection works
    const result = await connection.query(
      "SELECT COUNT(*) AS table_count FROM MSysObjects WHERE Type=1 AND Flags=0"
    );
    const count = result[0]?.table_count ?? "(unknown)";
    console.log(`User tables found: ${count}`);
  } catch (err) {
    console.error("FAILED: Could not connect to the database.");
    console.error(`Error: ${err.message}`);
    process.exit(1);
  } finally {
    if (connection) await connection.close();
  }
}

testConnection();
