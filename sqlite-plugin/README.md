# BrightSign SQLite Plugin (`sql.brs`)

This plugin runs a local SQLite-backed API server on a BrightSign player and also accepts UDP commands for common table/row operations.

- Plugin file: `sql.brs`
- Default DB file: `sd:/userData.db`
- Default table: `UserInfo`
- Default HTTP port: `8009`
- Add this plugin to BrightAuthor, naming it sql, and open the webpage at default port: `8009`

## What It Does

- Creates/opens a SQLite database on player storage.
- Ensures a default table exists at startup.
- Exposes REST endpoints for CRUD + schema/table utilities.
- Supports UDP commands for quick insert/update/delete/table operations.
- Serves a built-in UI/help page at `/` and `/help`.

## Configuration (in `sql.brs`)

Inside `sql_Initialize`, you can adjust:

- `h.dbFile` (database location)
- `h.defaultTable` (default table name)
- `h.webPort` (REST API port)
- `h.columnArray` (default schema for table creation)

Notes:
- `rowID` is always auto-managed as `INTEGER PRIMARY KEY AUTOINCREMENT`.
- `rowID` is reserved and should not be changed.

## Access URLs

From your network:

- `http://<PLAYER_IP>:8009/`
- `http://<PLAYER_IP>:8009/help`
- Or mDNS style: `http://brightsign-<SERIAL>.local:8009`

## REST API

### GET Endpoints

- `GET /` : Home UI page
- `GET /help` : Interactive API help page
- `GET /readAll?tableName=UserInfo`
- `GET /readRecord?tableName=UserInfo&rowID=1`
- `GET /readColumn?tableName=UserInfo&columnName=column1`
- `GET /search?tableName=UserInfo&columnName=column1&value=Alice`
- `GET /tableCount`
- `GET /schema`
- `GET /rowCount?tableName=UserInfo`

### POST Endpoints (`application/x-www-form-urlencoded`)

- `POST /createTable`
  - Required: `tableName`
  - Optional: `columnsJson` (JSON array, e.g. `[{"name":"orderNo"},{"name":"amount"}]`)
- `POST /insert`
  - Required: `tableName`
  - Then send column/value pairs for that table
- `POST /bulkInsert`
  - Required: `tableName`, `rowsJson` (JSON array of row objects)
- `POST /update`
  - Required: `tableName`, `rowID`
  - Then send updated column/value pairs
- `POST /delete`
  - Required: `tableName`, `rowID` (single-row delete)
- `POST /deleteWhere`
  - Required: `tableName`, `columnName`, `value`
- `POST /emptyTable`
  - Required: `tableName` (removes all rows)

## curl Examples

```bash
# Read all rows
curl "http://<PLAYER_IP>:8009/readAll?tableName=UserInfo"

# Create table with default columnArray
curl -X POST "http://<PLAYER_IP>:8009/createTable" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  --data "tableName=Orders"

# Create table with custom schema
curl -X POST "http://<PLAYER_IP>:8009/createTable" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  --data-urlencode "tableName=Orders" \
  --data-urlencode 'columnsJson=[{"name":"orderNo"},{"name":"amount"}]'

# Insert one row
curl -X POST "http://<PLAYER_IP>:8009/insert" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  --data "tableName=UserInfo&column1=Alice&column2=Engineer&column3=Mississauga"

# Update row 1
curl -X POST "http://<PLAYER_IP>:8009/update" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  --data "tableName=UserInfo&rowID=1&column1=AliceUpdated"

# Delete row 1
curl -X POST "http://<PLAYER_IP>:8009/delete" \
  -H "Content-Type: application/x-www-form-urlencoded" \
  --data "tableName=UserInfo&rowID=1"
```

## UDP Commands

The plugin parses underscore-delimited commands.

- `Insert_<tableName>_<newValue1>_<newValue2>_<newValue3>`
- `UpdateRow_<tableName>_<rowID>_<newValue1>_<newValue2>_<newValue3>`
- `UpdateColumn_<tableName>_<rowID>_<columnName>_<newValue>`
- `Delete_<tableName>_<rowID>`
- `CreateTable_<tableName>`
- `EmptyTable_<tableName>`

Example:

```text
CreateTable_UserTest
Insert_UserInfo_Alice_Engineer_Mississauga
UpdateColumn_UserInfo_1_column2_Manager
Delete_UserInfo_1
```

## Included Files

- `sql.brs` : Main SQLite plugin and API server

## Troubleshooting

- If API is unreachable, verify player IP and port `8009`.
- If table/column actions fail, validate table/column names and request body fields.
- If DB operations fail, confirm writable storage path (`sd:/`).
- Open `/help` on the player for a quick interactive endpoint tester.
