require("dotenv").config({ path: ".env.local" });

const mysql = require("mysql2/promise");
const bcrypt = require("bcryptjs");
const { randomUUID } = require("crypto");

async function main() {
  const email = "admin@gmail.com";
  const password = "admin123456";
  const name = "Admin";

  const pool = mysql.createPool({
    host: process.env.MYSQL_HOST,
    port: Number(process.env.MYSQL_PORT || 3306),
    user: process.env.MYSQL_USER,
    password: process.env.MYSQL_PASSWORD,
    database: process.env.MYSQL_DATABASE,
  });

  const passwordHash = await bcrypt.hash(password, 10);

  await pool.execute(
    `
    INSERT INTO users (id, name, email, password_hash, role)
    VALUES (?, ?, ?, ?, ?)
    ON DUPLICATE KEY UPDATE
      name = VALUES(name),
      password_hash = VALUES(password_hash),
      role = VALUES(role)
    `,
    [randomUUID(), name, email, passwordHash, "admin"]
  );

  await pool.end();

  console.log("Admin password reset successfully");
  console.log("Email:", email);
  console.log("Password:", password);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
