import sql, { ConnectionPool, config as SqlConfig, ISqlTypeFactory, ISqlType } from "mssql";
import config from "./config";

class DatabaseHelper {
  private static pool: ConnectionPool;

  private static sqlConfig: SqlConfig = {
    user: config.dbUser,
    password: config.dbPassword,
    server: config.dbServer,
    database: config.dbName,
    options: {
      encrypt: true, // Use true for Azure SQL
      trustServerCertificate: true, // Change to true for local dev if needed
    },
    pool: {
      max: 10,
      min: 0,
      idleTimeoutMillis: 30000,
    },
  };

  private static async getPool(): Promise<ConnectionPool> {
    if (!DatabaseHelper.pool) {
      DatabaseHelper.pool = await sql.connect(DatabaseHelper.sqlConfig);
    }
    return DatabaseHelper.pool;
  }

  public static async executeQuery(
    query: string,
    params?: { [key: string]: any }
  ): Promise<any> {
    try {
      const pool = await DatabaseHelper.getPool();
      const request = pool.request();

      if (params) {
        for (const [key, value] of Object.entries(params)) {
          request.input(key, value);
        }
      }

      const result = await request.query(query);
      return result.recordset;
    } catch (err) {
      throw new Error(`Database query failed: ${(err as Error).message}`);
    }
  }

  public static async bulkInsert(
    tableName: string,
    columns: { name: string; type: ISqlType | (() => ISqlType); nullable?: boolean }[],
    rows: any[]
  ): Promise<void> {
    try {
      const pool = await DatabaseHelper.getPool();

      const table = new sql.Table(tableName);
      table.create = false;

      // Add columns with correct type
      for (const col of columns) {
        table.columns.add(col.name, col.type, { nullable: col.nullable ?? true });
      }

      // Add rows
      for (const row of rows) {
        if (typeof row === "object" && !Array.isArray(row)) {
          const values = columns.map((col) => row[col.name]);
          table.rows.add(...values);
        } else if (Array.isArray(row)) {
          table.rows.add(...row);
        } else {
          throw new Error("Row must be an object or array");
        }
      }

      await pool.request().bulk(table);
    } catch (err) {
      throw new Error(`Bulk insert failed: ${(err as Error).message}`);
    }
  }
}

export default DatabaseHelper;
