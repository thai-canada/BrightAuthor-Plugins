' plugin name: sql
' Open http://brightsign_IP:8009 to get started
' Or http://brightsign-<SERIAL>.local:8009
' created by: Thai Nguyen 2026-06-08

Function sql_Initialize(msgPort As Object, userVariables As Object, bsp as Object)

	h = {}
	h.version = "1.00.00"
	h.msgPort = msgPort
	h.userVariables = userVariables
	h.bsp = bsp
	h.SystemLog = CreateObject("roSystemLog")
	h.ProcessEvent = sql_ProcessEvent
	h.userData = {}
	h.tableNames = []
	h.isReady = false  ' Set to true when initialization is complete

	'-----------------------------------------------------------------------------------------------------------
	h.dbFile = "sd:/userData.db"
	h.defaultTable = "UserInfo"
	h.defaultRowID = "rowID" ' DO NOT CHANGE THIS
	h.webPort = 8009 ' port of REST API.
	' the column structure. ------------------------------------------------------------------------------------
	h.columnArray = [
		{ name: "column1" },
		{ name: "column2" },
		{ name: "column3" }
	]
	' This creates the UserInfo table during startup using the default column array above.
	' You can edit the column names and number of columns in this columnArray.
	' It will be used to create the new table when columnsJson is not provided in the
	' POST /createTable request.

	' rowID is added automatically as INTEGER PRIMARY KEY AUTOINCREMENT.
	' All other column values are TEXT.
	' -----------------------------------------------------------------------------------------------------------

	h.SqlHomeAA = { HandleEvent: sql_GetHomePage, mVar: h }
	h.SqlHelpAA = { HandleEvent: sql_GetHelpPage, mVar: h }
	h.SqlReadSelectedAA = { HandleEvent: sql_PostReadRecordsRest, mVar: h }
	h.SqlReadRowAA = { HandleEvent: sql_GetReadRowRest, mVar: h }
	h.SqlReadColumnAA = { HandleEvent: sql_GetReadColumnRest, mVar: h }
	h.SqlSearchAA = { HandleEvent: sql_GetSearchRowsRest, mVar: h }
	h.SqlTablesAA = { HandleEvent: sql_GetTablesRest, mVar: h }
	h.SqlSchemaAA = { HandleEvent: sql_GetSchemaRest, mVar: h }
	h.SqlCreateAA = { HandleEvent: sql_CreateTableRest, mVar: h }
	h.SqlInsertAA = { HandleEvent: sql_InsertRecordRest, mVar: h }
	h.SqlUpdateAA = { HandleEvent: sql_UpdateRecordRest, mVar: h }
	h.SqlDeleteAA = { HandleEvent: sql_DeleteRecordRest, mVar: h }
	h.SqlDeleteTableAA = { HandleEvent: sql_DeleteTableRest, mVar: h }
	h.SqlCountAA = { HandleEvent: sql_GetCountRest, mVar: h }
	h.SqlDeleteWhereAA = { HandleEvent: sql_DeleteWhereRest, mVar: h }
	h.SqlBulkInsertAA = { HandleEvent: sql_BulkInsertRest, mVar: h }

	h.localServer = CreateObject("roHttpServer", { port: h.webPort })
	if type(h.localServer) = "roHttpServer" then
		h.localServer.SetPort(msgPort)
		h.localServer.AddGetFromEvent({ url_path: "/", user_data: h.SqlHomeAA })
		h.localServer.AddGetFromEvent({ url_path: "/help", user_data: h.SqlHelpAA })
		h.localServer.AddGetFromEvent({ url_path: "/readAll", user_data: h.SqlReadSelectedAA })
		h.localServer.AddGetFromEvent({ url_path: "/readRecord", user_data: h.SqlReadRowAA })
		h.localServer.AddGetFromEvent({ url_path: "/readColumn", user_data: h.SqlReadColumnAA })
		h.localServer.AddGetFromEvent({ url_path: "/search", user_data: h.SqlSearchAA })
		h.localServer.AddGetFromEvent({ url_path: "/tableCount", user_data: h.SqlTablesAA })
		h.localServer.AddGetFromEvent({ url_path: "/schema", user_data: h.SqlSchemaAA })
		h.localServer.AddGetFromEvent({ url_path: "/rowCount", user_data: h.SqlCountAA })
		h.localServer.AddPostToFormData({ url_path: "/createTable", user_data: h.SqlCreateAA })
		h.localServer.AddPostToFormData({ url_path: "/insert", user_data: h.SqlInsertAA })
		h.localServer.AddPostToFormData({ url_path: "/update", user_data: h.SqlUpdateAA })
		h.localServer.AddPostToFormData({ url_path: "/delete", user_data: h.SqlDeleteAA })
		h.localServer.AddPostToFormData({ url_path: "/emptyTable", user_data: h.SqlDeleteTableAA })
		h.localServer.AddPostToFormData({ url_path: "/deleteWhere", user_data: h.SqlDeleteWhereAA })
		h.localServer.AddPostToFormData({ url_path: "/bulkInsert", user_data: h.SqlBulkInsertAA })
		h.SystemLog.SendLine("---------------------- SQLite web server running at port:" + h.webPort.ToStr())
	else
		h.SystemLog.SendLine("Failed to create SQLite web server")
	end if


	' normalize the default columnArray to ensure it is in the correct format and has valid column names. ------
	normalizedDefaultColumns = sql_NormalizeCreateColumnArray(h.columnArray, h.defaultRowID)
	if type(normalizedDefaultColumns) = "roArray" and normalizedDefaultColumns.Count() > 0 then
		h.columnArray = normalizedDefaultColumns
	end if

	' Make sure the database file exists and is initialized before accepting any commands or API calls. ---------
	sql_EnsureDatabaseFile(h)
	' -----------------------------------------------------------------------------------------------------------


	h.tableNames = sql_GetTableNames(h.dbFile)
	sql_SyncDefaultTable(h)

	return h

End Function

Function sql_ProcessEvent(event As Object) as boolean
	retval = false

	if m.isReady = false then
		if m.bsp <> invalid and m.bsp.udpReceivePort <> invalid then
			m.isReady = true
			EnableUDPNotifications(m)
		end if
	end if

    if type(event) = "roAssociativeArray" then ' Receive a message from BA
    else if type(event) = "roHtmlWidgetEvent" then
		' HTML widget messages are intentionally ignored in this plugin build.

		eventData = event.GetData()
		if type(eventData) = "roAssociativeArray" and type(eventData.reason) = "roString" then

			' Example of sending message from HTML5 zone to BrightScript:
			' var bsMessage = new BSMessagePort();
			' bsMessage.PostBSMessage({complete: true, result: "Insert_<tableName>_<newValue1>_<newValue2>_<newValue3>"});

			if eventData.reason = "message" then
				if eventData.message.result = valid then

					msg$ = eventData.message.result
					
					regex = CreateObject("roRegEx", "_", "i")
					parts = sql_CommandPartsToArray(regex.Split(msg$))
					'm.SystemLog.SendLine("------------------ Received roHtmlWidgetEvent message: " + msg$)
					retval = sql_HandleCommand(m, msg$, parts)
				end if
			end if
		end if

	else if type(event) = "roDatagramEvent" then

		msg$ = event.GetString()
		
		regex = CreateObject("roRegEx", "_", "i")
		parts = sql_CommandPartsToArray(regex.Split(msg$))
		'm.SystemLog.SendLine("------------------ Received UDP message: " + msg$)
		retval = sql_HandleCommand(m, msg$, parts)

    end if
    return retval
End Function

Function sql_HandleCommand(m as Object, msg$ as String, parts as Object) as Boolean
	if left(msg$, 7) = "Insert_" then
		' Example: Insert_<tableName>_<newValue1>_<newValue2>_<newValue3>
		sql_HandleInsert(m, parts)
		return true
	else if left(msg$, 10) = "UpdateRow_" then
		' Example: UpdateRow_<tableName>_<rowID>_<newValue1>_<newValue2>_<newValue3>
		sql_HandleUpdateRow(m, parts)
		return true
	else if left(msg$, 13) = "UpdateColumn_" then
		' Example: UpdateColumn_<tableName>_<rowID>_<columnName>_<newValue>
		sql_HandleUpdateColumn(m, parts)
		return true
	else if left(msg$, 7) = "Delete_" then
		' Example: Delete_<tableName>_<rowID>
		sql_HandleDelete(m, parts)
		return true
	else if left(msg$, 12) = "CreateTable_" then
		' Example: CreateTable_<tableName>
		sql_HandleCreateTable(m, parts)
		return true
	else if left(msg$, 11) = "EmptyTable_" then
		' Example: EmptyTable_<tableName>
		sql_HandleEmptyTable(m, parts)
		return true
	end if

	return false
End Function

Function sql_CommandPartsToArray(parts as Object) as Object
	if type(parts) = "roArray" then
		return parts
	end if

	arr = []
	if parts = invalid then
		return arr
	end if

	for each item in parts
		arr.Push(item)
	next

	return arr
End Function

Sub sql_HandleInsert(m as Object, parts as Object)
	if type(parts) <> "roArray" or parts.Count() < 2 then
		return
	end if

	tableName$ = sql_TrimString(parts[1])
	if not sql_IsValidTableName(tableName$) then
		return
	end if

	columns = sql_GetInsertableColumnsForTable(m, m.dbFile, tableName$)
	values = sql_BuildValuesFromParts(parts, 2, columns.Count())
	sql_InsertUserData(m, m.dbFile, tableName$, columns, values)
End Sub

Sub sql_HandleUpdateRow(m as Object, parts as Object)
	if type(parts) <> "roArray" or parts.Count() < 3 then
		return
	end if

	tableName$ = sql_TrimString(parts[1])
	rowID$ = sql_TrimString(parts[2])
	rowIdNum = sql_ParsePositiveIntStrict(rowID$)
	if not sql_IsValidTableName(tableName$) or rowIdNum <= 0 then
		return
	end if

	columns = sql_GetInsertableColumnsForTable(m, m.dbFile, tableName$)
	values = sql_BuildValuesFromParts(parts, 3, columns.Count())
	sql_UpdateRecord(m, m.dbFile, tableName$, rowIdNum, columns, values)
End Sub

Sub sql_HandleUpdateColumn(m as Object, parts as Object)
	if type(parts) <> "roArray" or parts.Count() < 5 then
		return
	end if

	tableName$ = sql_TrimString(parts[1])
	rowID$ = sql_TrimString(parts[2])
	columnName$ = sql_TrimString(parts[3])
	rowIdNum = sql_ParsePositiveIntStrict(rowID$)
	if not sql_IsValidTableName(tableName$) or rowIdNum <= 0 or not sql_IsValidTableName(columnName$) then
		return
	end if

	if not sql_ColumnExists(m.dbFile, tableName$, columnName$) then
		return
	end if

	newValue$ = sql_JoinPartsWithUnderscore(parts, 4)
	columns = [columnName$]
	values = [newValue$]
	sql_UpdateRecord(m, m.dbFile, tableName$, rowIdNum, columns, values)
End Sub

Sub sql_HandleDelete(m as Object, parts as Object)
	if type(parts) <> "roArray" or parts.Count() <> 3 then
		return
	end if

	tableName$ = sql_TrimString(parts[1])
	rowID$ = sql_TrimString(parts[2])
	rowIdNum = sql_ParsePositiveIntStrict(rowID$)
	if not sql_IsValidTableName(tableName$) or rowIdNum <= 0 then
		return
	end if

	sql_DeleteRecord(m, m.dbFile, tableName$, rowIdNum)
End Sub

Sub sql_HandleCreateTable(m as Object, parts as Object)
	if type(parts) <> "roArray" or parts.Count() < 2 then
		return
	end if

	tableName$ = sql_TrimString(parts[1])
	if not sql_IsValidTableName(tableName$) then
		return
	end if

	createTableSQL$ = sql_BuildCreateTableSQL(tableName$, m.columnArray)
	if createTableSQL$ = "" then
		return
	end if

	if sql_CreateAndInitializeCustomDB(m, m.dbFile, createTableSQL$) then
		m.tableNames = sql_GetTableNames(m.dbFile)
	end if
End Sub

Sub sql_HandleEmptyTable(m as Object, parts as Object)
	if type(parts) <> "roArray" or parts.Count() < 2 then
		return
	end if

	tableName$ = sql_TrimString(parts[1])
	if not sql_IsValidTableName(tableName$) then
		return
	end if

	sql_DeleteTableRows(m, m.dbFile, tableName$)
End Sub




Function sql_CreateAndInitializeCustomDB(m as Object, dbName$ as String, createTableSQL$ as String) as Boolean
    myDB = CreateObject("roSqliteDatabase")
    success = false
    resolvedDbName$ = sql_ResolveDbPath(dbName$)

    if myDB.Create(resolvedDbName$) = invalid then
		if myDB.Open(resolvedDbName$) = invalid then
			m.SystemLog.SendLine("---------------------- Failed to create/open custom database file: " + resolvedDbName$)
			return false
		end if
	end if

    createStatement = myDB.CreateStatement(createTableSQL$)
    if type(createStatement) = "roSqliteStatement" then
        SQLITE_COMPLETE = 100
		result = createStatement.Run()
		if result = SQLITE_COMPLETE or result = 101 then
			m.SystemLog.SendLine("---------------------- Custom table created successfully.")
			success = true
		else
			m.SystemLog.SendLine("---------------------- Error creating table. Error code: " + result.ToStr())
		end if
		createStatement.Finalise()
    else
		' Some firmware builds return invalid statement on first handle. Re-open and retry once.
		myDB.Close()
		myDB = CreateObject("roSqliteDatabase")
		if myDB.Open(resolvedDbName$) = invalid then
			m.SystemLog.SendLine("---------------------- Failed to re-open database for statement retry: " + resolvedDbName$)
			return false
		end if

		createStatement = myDB.CreateStatement(createTableSQL$)
		if type(createStatement) <> "roSqliteStatement" then
			m.SystemLog.SendLine("---------------------- Failed to create SQL statement for table creation (retry).")
			myDB.Close()
			return false
		end if

		SQLITE_COMPLETE = 100
		result = createStatement.Run()
		if result = SQLITE_COMPLETE or result = 101 then
			m.SystemLog.SendLine("---------------------- Custom table created successfully (retry).")
			success = true
		else
			m.SystemLog.SendLine("---------------------- Error creating table on retry. Error code: " + result.ToStr())
		end if
		createStatement.Finalise()
    end if

	myDB.Close()
	return success
End Function


' ===== HTML Pages =====
Sub sql_GetHelpPage(userData as Object, e as Object)
	mVar = userData.mVar
	portText$ = mVar.webPort.ToStr()
	columnsJsonPlaceholder$ = "[{" + Chr(34) + "name" + Chr(34) + ":" + Chr(34) + "orderNo" + Chr(34) + "},{" + Chr(34) + "name" + Chr(34) + ":" + Chr(34) + "amount" + Chr(34) + "}]"
	html$ = "<html>"
	html$ = html$ + "<head>"
	html$ = html$ + "<title>SQLite API Help</title>"
	html$ = html$ + "<style>"
	html$ = html$ + "body{font-family:Arial,sans-serif;background:#f6f6f6;color:#222;margin:0;padding:18px;}"
	html$ = html$ + "h1{margin:0 0 12px 0;}"
	html$ = html$ + "h2{margin:18px 0 6px 0;font-size:18px;}"
	html$ = html$ + "p{margin:6px 0;}"
	html$ = html$ + ".card{background:#fff;border:1px solid #ddd;border-radius:10px;padding:12px;margin:10px 0;}"
	html$ = html$ + "pre{margin:8px 0 0 0;padding:10px;background:#111;color:#f3f3f3;border-radius:8px;overflow:auto;white-space:pre-wrap;}"
	html$ = html$ + "code{display:block;margin:8px 0 0 0;padding:10px;background:#eef3f8;color:#1d2a38;border:1px solid #d4dde7;border-radius:8px;white-space:pre-wrap;}"
	html$ = html$ + ".api-head{display:flex;align-items:center;justify-content:space-between;gap:10px;}"
	html$ = html$ + ".try-btn{background:#111;color:#fff;border:1px solid #111;border-radius:8px;padding:7px 10px;cursor:pointer;font-weight:600;}"
	html$ = html$ + ".try-btn:hover{background:#2d2d2d;}"
	html$ = html$ + ".sql-try-modal{position:fixed;inset:0;background:rgba(0,0,0,0.48);display:none;align-items:center;justify-content:center;z-index:9999;}"
	html$ = html$ + ".sql-try-modal.is-open{display:flex;}"
	html$ = html$ + ".sql-try-card{width:min(760px,94vw);max-height:86vh;overflow:auto;background:#fff;border:1px solid #d4d4d4;border-radius:12px;box-shadow:0 14px 38px rgba(0,0,0,0.28);padding:14px;}"
	html$ = html$ + ".sql-try-grid{display:grid;grid-template-columns:1fr;gap:10px;margin-top:10px;}"
	html$ = html$ + ".sql-try-grid label{font-size:13px;font-weight:600;color:#1d2a38;}"
	html$ = html$ + ".sql-try-grid input,.sql-try-grid select{padding:8px 10px;border:1px solid #888;border-radius:8px;font-size:14px;}"
	html$ = html$ + "#sqlPostField_tableName{width:100%;box-sizing:border-box;min-width:560px;}"
	html$ = html$ + "#sqlPostField_columnsJson,#sqlPostField_rawBody{width:100%;box-sizing:border-box;min-width:560px;}"
	html$ = html$ + "@media (max-width: 720px){#sqlPostField_tableName{min-width:0;}}"
	html$ = html$ + "@media (max-width: 720px){#sqlPostField_columnsJson,#sqlPostField_rawBody{min-width:0;}}"
	html$ = html$ + ".sql-try-actions{display:flex;gap:8px;align-items:center;margin-top:12px;}"
	html$ = html$ + ".sql-try-result{margin-top:10px;max-height:220px;overflow:auto;}"
	html$ = html$ + "</style>"
	html$ = html$ + "</head>"
	html$ = html$ + "<body>"
	html$ = html$ + "<h1>SQLite API Help</h1>"
	html$ = html$ + "<p>Server IP: <strong id='serverIp'></strong> at port: <b>" + portText$ + "</b></p>"
	'html$ = html$ + "<p>Default table use in this page is: <strong>UserInfo</strong></p>"
	html$ = html$ + "<div class='card'><h2>GET /</h2><pre>curl http://window.location.hostname:8009</pre><code>Returns the UI with all records in the table.</code></div>"
	html$ = html$ + "<div class='card'><h2>GET /readAll</h2><pre>curl http://window.location.hostname:8009/readAll (no parameter uses default table)</pre><pre>curl http://window.location.hostname:8009/readAll?tableName=UserInfo</pre><code>{&quot;rows&quot;:[{&quot;rowID&quot;:1,&quot;column1&quot;:&quot;Alice&quot;,&quot;column2&quot;:&quot;Engineer&quot;,&quot;column3&quot;:&quot;Mississauga&quot;}]}</code></div>"
	html$ = html$ + "<div class='card'><h2>GET /readRecord</h2><pre>curl http://window.location.hostname:8009/readRecord?tableName=UserInfo&rowID=1</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;table&quot;:&quot;UserInfo&quot;,&quot;row&quot;:{&quot;rowID&quot;:1,&quot;column1&quot;:&quot;Alice&quot;,&quot;column2&quot;:&quot;Engineer&quot;,&quot;column3&quot;:&quot;Mississauga&quot;}}</code></div>"
	html$ = html$ + "<div class='card'><h2>GET /readColumn</h2><pre>curl http://window.location.hostname:8009/readColumn?tableName=UserInfo&columnName=column1</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;table&quot;:&quot;UserInfo&quot;,&quot;column&quot;:&quot;column1&quot;,&quot;values&quot;:[&quot;Alice&quot;,&quot;Bob&quot;]}</code></div>"
	html$ = html$ + "<div class='card'><h2>GET /search</h2><pre>curl http://window.location.hostname:8009/search?tableName=UserInfo&columnName=column1&value=Alice</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;dbFile&quot;:&quot;sd:/userData.db&quot;,&quot;table&quot;:&quot;UserInfo&quot;,&quot;column&quot;:&quot;column1&quot;,&quot;value&quot;:&quot;Alice&quot;,&quot;count&quot;:1,&quot;rows&quot;:[{&quot;rowID&quot;:1,&quot;column1&quot;:&quot;Alice&quot;,&quot;column2&quot;:&quot;Engineer&quot;,&quot;column3&quot;:&quot;Mississauga&quot;}]}</code></div>"
	html$ = html$ + "<div class='card'><h2>GET /tableCount</h2><pre>curl http://window.location.hostname:8009/tableCount (returns an array of existing tables)</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;count&quot;:2,&quot;tables&quot;:[&quot;UserInfo&quot;,&quot;UserTest&quot;]}</code></div>"
	html$ = html$ + "<div class='card'><h2>GET /rowCount</h2><pre>curl http://window.location.hostname:8009/rowCount (no parameter uses default table)</pre><pre>curl http://window.location.hostname:8009/rowCount?tableName=UserInfo</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;table&quot;:&quot;UserInfo&quot;,&quot;count&quot;:5}</code></div>"
	html$ = html$ + "<div class='card'><h2>GET /schema</h2><pre>curl http://window.location.hostname:8009/schema</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;countTable&quot;:2,&quot;schema&quot;:[{&quot;table&quot;:&quot;UserInfo&quot;,&quot;columnCount&quot;:3,&quot;columns&quot;:[&quot;column1&quot;,&quot;column2&quot;,&quot;column3&quot;]},{&quot;table&quot;:&quot;UserTest&quot;,&quot;columnCount&quot;:2,&quot;columns&quot;:[&quot;name&quot;,&quot;email&quot;]}]}</code></div>"
	html$ = html$ + "<div class='card'><div class='api-head'><h2>POST /insert</h2><input type='button' value='TRY' class='try-btn' data-endpoint='/insert' onclick='sqlOpenPostPopupFromButton(this)' /></div><pre>curl.exe -X POST http://window.location.hostname:8009/insert -H Content-Type:application/x-www-form-urlencoded" + Chr(10) + "--data tableName=UserInfo&column1=Alice&column2=Engineer&column3=Mississauga</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;table&quot;:&quot;UserInfo&quot;,&quot;message&quot;:&quot;Record inserted&quot;}</code></div>"
	html$ = html$ + "<div class='card'><div class='api-head'><h2>POST /bulkInsert</h2><input type='button' value='TRY' class='try-btn' data-endpoint='/bulkInsert' onclick='sqlOpenPostPopupFromButton(this)' /></div><pre>curl.exe -X POST http://window.location.hostname:8009/bulkInsert -H Content-Type:application/x-www-form-urlencoded" + Chr(10) + " --data-urlencode tableName=UserInfo" + Chr(10) + " --data-urlencode rowsJson=[{&quot;column1&quot;:&quot;Alice&quot;,&quot;column2&quot;:&quot;Engineer&quot;,&quot;column3&quot;:&quot;Mississauga&quot;},{&quot;column1&quot;:&quot;Bob&quot;,&quot;column2&quot;:&quot;Manager&quot;,&quot;column3&quot;:&quot;LA&quot;}]</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;table&quot;:&quot;UserInfo&quot;,&quot;requestedCount&quot;:2,&quot;insertedCount&quot;:2,&quot;failedCount&quot;:0}<BR>rowsJson must be in JSON format</code></div>"
	html$ = html$ + "<div class='card'><div class='api-head'><h2>POST /update</h2><input type='button' value='TRY' class='try-btn' data-endpoint='/update' onclick='sqlOpenPostPopupFromButton(this)' /></div><pre>curl.exe -X POST http://window.location.hostname:8009/update -H Content-Type:application/x-www-form-urlencoded" + Chr(10) + " --data tableName=UserInfo&rowID=1&column1=AliceUpdated&column2=Lead&column3=LA</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;table&quot;:&quot;UserInfo&quot;,&quot;rowID&quot;:&quot;1&quot;,&quot;message&quot;:&quot;Record updated&quot;}</code></div>"
	html$ = html$ + "<div class='card'><div class='api-head'><h2>POST /delete</h2><input type='button' value='TRY' class='try-btn' data-endpoint='/delete' onclick='sqlOpenPostPopupFromButton(this)' /></div><pre>curl.exe -X POST http://window.location.hostname:8009/delete -H Content-Type:application/x-www-form-urlencoded" + Chr(10) + " --data tableName=UserInfo&rowID=1</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;table&quot;:&quot;UserInfo&quot;,&quot;rowID&quot;:&quot;1&quot;,&quot;message&quot;:&quot;Record deleted&quot;}</code></div>"
	html$ = html$ + "<div class='card'><div class='api-head'><h2>POST /deleteWhere</h2><input type='button' value='TRY' class='try-btn' data-endpoint='/deleteWhere' onclick='sqlOpenPostPopupFromButton(this)' /></div><pre>curl.exe -X POST http://window.location.hostname:8009/deleteWhere -H Content-Type:application/x-www-form-urlencoded" + Chr(10) + " --data tableName=UserInfo&columnName=column1&value=Alice</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;table&quot;:&quot;UserInfo&quot;,&quot;columnName&quot;:&quot;column1&quot;,&quot;deletedCount&quot;:1,&quot;message&quot;:&quot;Rows deleted&quot;}</code></div>"
	html$ = html$ + "<div class='card'><div class='api-head'><h2>POST /createTable</h2><input type='button' value='TRY' class='try-btn' data-endpoint='/createTable' onclick='sqlOpenPostPopupFromButton(this)' /></div><pre>curl.exe -X POST http://window.location.hostname:8009/createTable -H Content-Type:application/x-www-form-urlencoded" + Chr(10) + " --data tableName=UserInfo</pre><code>This request uses the default columnArray defined in plugin file.<BR>{&quot;status&quot;:&quot;success&quot;,&quot;table&quot;:&quot;Orders&quot;,&quot;message&quot;:&quot;Table is ready&quot;}</code><pre>curl.exe -X POST http://window.location.hostname:8009/createTable -H Content-Type:application/x-www-form-urlencoded" + Chr(10) + " --data-urlencode tableName=Orders" + Chr(10) + " --data-urlencode columnsJson=[{&quot;name&quot;:&quot;orderNo&quot;},{&quot;name&quot;:&quot;amount&quot;}]</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;table&quot;:&quot;Orders&quot;,&quot;message&quot;:&quot;Table is ready&quot;}<BR>columnsJson must be a JSON array of {name} with unique valid names.</code></div>"
	html$ = html$ + "<div class='card'><div class='api-head'><h2>POST /emptyTable</h2><input type='button' value='TRY' class='try-btn' data-endpoint='/emptyTable' onclick='sqlOpenPostPopupFromButton(this)' /></div><pre>curl.exe -X POST http://window.location.hostname:8009/emptyTable -H Content-Type:application/x-www-form-urlencoded" + Chr(10) + "--data tableName=UserInfo</pre><code>{&quot;status&quot;:&quot;success&quot;,&quot;table&quot;:&quot;UserInfo&quot;,&quot;message&quot;:&quot;All rows deleted&quot;}</code></div>"
	html$ = html$ + "<div id='sqlPostTryModal' class='sql-try-modal' onclick='sqlClosePostPopupOnBackdrop(event)'><div class='sql-try-card'><div class='api-head'><h2 style='margin:0;'>Try POST Endpoint</h2><input type='button' value='Close' class='try-btn' onclick='sqlClosePostPopup()' /></div><div class='sql-try-grid'><label for='sqlPostEndpointSelect'>POST endpoint</label><select id='sqlPostEndpointSelect'><option value='/insert'>/insert</option><option value='/bulkInsert'>/bulkInsert</option><option value='/update'>/update</option><option value='/delete'>/delete</option><option value='/deleteWhere'>/deleteWhere</option><option value='/createTable'>/createTable</option><option value='/emptyTable'>/emptyTable</option></select><p id='sqlPostFieldCount'>Input fields: 0</p><div id='sqlPostFields'></div></div><div class='sql-try-actions'><input type='button' value='SEND REQUEST' class='try-btn' onclick='sqlSendPostTryRequest()' /></div><code id='sqlPostTryResult' class='sql-try-result'>Select endpoint and fill body fields.</code></div></div>"
	html$ = html$ + "<h1>UDP Command Help</h1>"
	html$ = html$ + "<div class='card'><pre>Insert_&lt;tableName&gt;_&lt;newValue1&gt;_&lt;newValue2&gt;_&lt;newValue3&gt;</pre><pre>UpdateRow_&lt;tableName&gt;_&lt;rowID&gt;_&lt;newValue1&gt;_&lt;newValue2&gt;_&lt;newValue3&gt;</pre><pre>UpdateColumn_&lt;tableName&gt;_&lt;rowID&gt;_&lt;columnName&gt;_&lt;newValue&gt;</pre><pre>Delete_&lt;tableName&gt;_&lt;rowID&gt;</pre><pre>CreateTable_&lt;tableName&gt; (UDP create uses current default schema, no columnsJson in UDP format)</pre><pre>EmptyTable_&lt;tableName&gt;</pre><h2>Send UDP to <span id='udpTargetHost'></span></h2><p><input id='udpCommandInput' type='text' placeholder='UDP command string. For example: CreateTable_UserTest' style='min-width:420px;padding:8px;border:1px solid #999;border-radius:6px;' /> <input type='button' value='Send UDP' onclick='sqlSendUdpCommand()' style='padding:8px;border:1px solid #999;border-radius:6px;cursor:pointer;' /></p><code id='udpSendResult'>Waiting for command...</code></div>"
	html$ = html$ + "<script>"
	html$ = html$ + "var serverIp = window.location.hostname;"
	html$ = html$ + "var serverBase = 'http://' + serverIp + ':8009';"
	html$ = html$ + "var postTryModal = document.getElementById('sqlPostTryModal');"
	html$ = html$ + "var postTryEndpointSelect = document.getElementById('sqlPostEndpointSelect');"
	html$ = html$ + "var postTryFields = document.getElementById('sqlPostFields');"
	html$ = html$ + "var postTryFieldCount = document.getElementById('sqlPostFieldCount');"
	html$ = html$ + "var postTryResult = document.getElementById('sqlPostTryResult');"
	html$ = html$ + "var endpointFieldMap = {};"
	html$ = html$ + "endpointFieldMap['/createTable'] = ['tableName','columnsJson'];"
	html$ = html$ + "endpointFieldMap['/insert'] = ['tableName','rawBody'];"
	html$ = html$ + "endpointFieldMap['/update'] = ['tableName','rawBody'];"
	html$ = html$ + "endpointFieldMap['/delete'] = ['tableName','rowID'];"
	html$ = html$ + "endpointFieldMap['/deleteWhere'] = ['tableName','columnName','value'];"
	html$ = html$ + "endpointFieldMap['/emptyTable'] = ['tableName'];"
	html$ = html$ + "endpointFieldMap['/bulkInsert'] = ['tableName','rawBody'];"
	html$ = html$ + "var endpointFieldDefaults = {};"
	html$ = html$ + "endpointFieldDefaults.tableName = 'UserInfo';"
	html$ = html$ + "endpointFieldDefaults.rowID = '1';"
	html$ = html$ + "endpointFieldDefaults.rawBody = '';"
	html$ = html$ + "endpointFieldDefaults.columnName = 'column1';"
	html$ = html$ + "endpointFieldDefaults.value = 'Alice';"
	html$ = html$ + "endpointFieldDefaults.columnsJson = '';"
	html$ = html$ + "endpointFieldDefaults.rowsJson = '';"
	html$ = html$ + "var endpointFieldPlaceholders = {};"
	html$ = html$ + "endpointFieldPlaceholders.columnsJson = '" + columnsJsonPlaceholder$ + "';"
	html$ = html$ + "endpointFieldPlaceholders.rawBodyInsert = 'column1=Alice&column2=Engineer&column3=Mississauga';"
	html$ = html$ + "endpointFieldPlaceholders.rawBodyUpdate = 'rowID=1&column1=AliceUpdated&column2=Lead&column3=LA';"
	html$ = html$ + "endpointFieldPlaceholders.rawBodyBulkInsert = 'rowsJson=[{" + Chr(34) + "column1" + Chr(34) + ":" + Chr(34) + "Alice" + Chr(34) + "," + Chr(34) + "column2" + Chr(34) + ":" + Chr(34) + "Engineer" + Chr(34) + "},{" + Chr(34) + "column1" + Chr(34) + ":" + Chr(34) + "Bob" + Chr(34) + "," + Chr(34) + "column2" + Chr(34) + ":" + Chr(34) + "Manager" + Chr(34) + "}]';"
	html$ = html$ + "var tableOptionsFromHelp = ['UserInfo'];"
	html$ = html$ + "document.getElementById('serverIp').textContent = serverIp;"
	html$ = html$ + "var udpTargetHostNode = document.getElementById('udpTargetHost');"
	html$ = html$ + "if (udpTargetHostNode) { udpTargetHostNode.textContent = serverIp; }"
	html$ = html$ + "var examples = document.getElementsByTagName('pre');"
	html$ = html$ + "for (var i = 0; i < examples.length; i++) { examples[i].textContent = examples[i].textContent.split('window.location.hostname').join(serverIp); }"
	html$ = html$ + "window.sqlRenderPostTryFields = function(){"
	html$ = html$ + "  if(!postTryEndpointSelect || !postTryFields){ return; }"
	html$ = html$ + "  var endpoint = postTryEndpointSelect.value || '/createTable';"
	html$ = html$ + "  var fields = endpointFieldMap[endpoint] || [];"
	html$ = html$ + "  postTryFields.innerHTML = '';"
	html$ = html$ + "  if(postTryFieldCount){"
	html$ = html$ + "    if(endpoint === '/createTable'){"
	html$ = html$ + "      postTryFieldCount.innerHTML = 'Input fields: 2  (If <b>columnsJson</b> is empty, it will use the default columnArray defined in the plugin file).';"
	html$ = html$ + "    } else {"
	html$ = html$ + "      postTryFieldCount.innerHTML = 'Input fields: ' + fields.length;"
	html$ = html$ + "    }"
	html$ = html$ + "  }"
	html$ = html$ + "  for(var idx = 0; idx < fields.length; idx++) {"
	html$ = html$ + "    var fieldName = fields[idx];"
	html$ = html$ + "    var wrapper = document.createElement('div');"
	html$ = html$ + "    var label = document.createElement('label');"
	html$ = html$ + "    label.textContent = fieldName + ' (string)';"
	html$ = html$ + "    label.setAttribute('for', 'sqlPostField_' + fieldName);"
	html$ = html$ + "    wrapper.appendChild(label);"
	html$ = html$ + "    if(fieldName === 'tableName') {"
	html$ = html$ + "      if(endpoint === '/createTable') {"
	html$ = html$ + "        var tableInput = document.createElement('input');"
	html$ = html$ + "        tableInput.type = 'text';"
	html$ = html$ + "        tableInput.id = 'sqlPostField_' + fieldName;"
	html$ = html$ + "        tableInput.setAttribute('data-field-name', fieldName);"
	html$ = html$ + "        tableInput.value = '';"
	html$ = html$ + "        tableInput.placeholder = 'NewTableName';"
	html$ = html$ + "        wrapper.appendChild(tableInput);"
	html$ = html$ + "      } else {"
	html$ = html$ + "        var tableSelect = document.createElement('select');"
	html$ = html$ + "        tableSelect.id = 'sqlPostField_' + fieldName;"
	html$ = html$ + "        tableSelect.setAttribute('data-field-name', fieldName);"
	html$ = html$ + "        for(var t = 0; t < tableOptionsFromHelp.length; t++) {"
	html$ = html$ + "          var op = document.createElement('option');"
	html$ = html$ + "          op.value = tableOptionsFromHelp[t];"
	html$ = html$ + "          op.textContent = tableOptionsFromHelp[t];"
	html$ = html$ + "          if(op.value === (endpointFieldDefaults.tableName || 'UserInfo')) { op.selected = true; }"
	html$ = html$ + "          tableSelect.appendChild(op);"
	html$ = html$ + "        }"
	html$ = html$ + "        tableSelect.addEventListener('change', function(){ if(postTryResult){ postTryResult.textContent = 'Fill all body fields as string and click SEND REQUEST.'; } });"
	html$ = html$ + "        wrapper.appendChild(tableSelect);"
	html$ = html$ + "      }"
	html$ = html$ + "    } else {"
	html$ = html$ + "      var input = document.createElement('input');"
	html$ = html$ + "      input.type = 'text';"
	html$ = html$ + "      input.id = 'sqlPostField_' + fieldName;"
	html$ = html$ + "      input.setAttribute('data-field-name', fieldName);"
	html$ = html$ + "      input.value = endpointFieldDefaults[fieldName] || '';"
	html$ = html$ + "      if(fieldName === 'columnsJson') { input.placeholder = endpointFieldPlaceholders.columnsJson; }"
	html$ = html$ + "      if(fieldName === 'rawBody' && endpoint === '/insert') { input.placeholder = endpointFieldPlaceholders.rawBodyInsert; }"
	html$ = html$ + "      if(fieldName === 'rawBody' && endpoint === '/update') { input.placeholder = endpointFieldPlaceholders.rawBodyUpdate; }"
	html$ = html$ + "      if(fieldName === 'rawBody' && endpoint === '/bulkInsert') { input.placeholder = endpointFieldPlaceholders.rawBodyBulkInsert; }"
	html$ = html$ + "      wrapper.appendChild(input);"
	html$ = html$ + "    }"
	html$ = html$ + "    postTryFields.appendChild(wrapper);"
	html$ = html$ + "  }"
	html$ = html$ + "};"
	html$ = html$ + "window.sqlResetPostTryResult = function(){ if(postTryResult){ postTryResult.textContent = 'Fill all body fields as string and click SEND REQUEST.'; } };"
	html$ = html$ + "window.sqlOpenPostPopupFromButton = function(btn){"
	html$ = html$ + "  var endpoint = '/createTable';"
	html$ = html$ + "  if(btn && btn.getAttribute){ endpoint = btn.getAttribute('data-endpoint') || endpoint; }"
	html$ = html$ + "  if(postTryEndpointSelect){ postTryEndpointSelect.value = endpoint; }"
	html$ = html$ + "  window.sqlRenderPostTryFields();"
	html$ = html$ + "  if(postTryResult){ postTryResult.textContent = 'Fill all body fields as string and click SEND REQUEST.'; }"
	html$ = html$ + "  if(postTryModal){ postTryModal.classList.add('is-open'); }"
	html$ = html$ + "};"
	html$ = html$ + "window.sqlClosePostPopup = function(){ if(postTryModal){ postTryModal.classList.remove('is-open'); } };"
	html$ = html$ + "window.sqlClosePostPopupOnBackdrop = function(ev){ if(ev.target === postTryModal){ window.sqlClosePostPopup(); } };"
	html$ = html$ + "window.sqlSendPostTryRequest = function(){"
	html$ = html$ + "  if(!postTryEndpointSelect){ return; }"
	html$ = html$ + "  var endpoint = postTryEndpointSelect.value || '/createTable';"
	html$ = html$ + "  var selectedTableName = '';"
	html$ = html$ + "  var params = new URLSearchParams();"
	html$ = html$ + "  var inputs = postTryFields ? postTryFields.querySelectorAll('input[data-field-name]') : [];"
	html$ = html$ + "  var selects = postTryFields ? postTryFields.querySelectorAll('select[data-field-name]') : [];"
	html$ = html$ + "  for(var s = 0; s < selects.length; s++) {"
	html$ = html$ + "    var skey = selects[s].getAttribute('data-field-name') || '';"
	html$ = html$ + "    var svalue = selects[s].value || '';"
	html$ = html$ + "    if(skey === 'tableName'){ selectedTableName = svalue; }"
	html$ = html$ + "    params.append(skey, svalue);"
	html$ = html$ + "  }"
	html$ = html$ + "  for(var i = 0; i < inputs.length; i++) {"
	html$ = html$ + "    var key = inputs[i].getAttribute('data-field-name') || '';"
	html$ = html$ + "    var value = inputs[i].value || '';"
	html$ = html$ + "    if(endpoint === '/deleteWhere' && key === 'value' && value.trim() === '') {"
	html$ = html$ + "      if(postTryResult){ postTryResult.textContent = 'Error: value is empty string.'; }"
	html$ = html$ + "      return;"
	html$ = html$ + "    }"
	html$ = html$ + "    if(key === 'rawBody') {"
	html$ = html$ + "      if(value.trim() === ''){"
	html$ = html$ + "        if(postTryResult){ postTryResult.textContent = 'Error: rawBody is empty string.'; }"
	html$ = html$ + "        return;"
	html$ = html$ + "      }"
	html$ = html$ + "      var rawParams = new URLSearchParams(value);"
	html$ = html$ + "      var hasRawPair = false;"
	html$ = html$ + "      rawParams.forEach(function(rawValue, rawKey){ hasRawPair = true; params.set(rawKey, rawValue); });"
	html$ = html$ + "      if(!hasRawPair){"
	html$ = html$ + "        if(postTryResult){ postTryResult.textContent = 'Error: rawBody is empty string.'; }"
	html$ = html$ + "        return;"
	html$ = html$ + "      }"
	html$ = html$ + "    } else {"
	html$ = html$ + "      params.append(key, value);"
	html$ = html$ + "    }"
	html$ = html$ + "  }"
	html$ = html$ + "  if((endpoint === '/insert' || endpoint === '/update' || endpoint === '/bulkInsert') && selectedTableName !== ''){"
	html$ = html$ + "    params.set('tableName', selectedTableName);"
	html$ = html$ + "  }"
	html$ = html$ + "  if(postTryResult){ postTryResult.textContent = 'Sending...'; }"
	html$ = html$ + "  fetch(serverBase + endpoint, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' }, body: params.toString() })"
	html$ = html$ + "    .then(function(resp){ return resp.text(); })"
	html$ = html$ + "    .then(function(text){"
	html$ = html$ + "      if(!postTryResult){ return; }"
	html$ = html$ + "      try {"
	html$ = html$ + "        var respObj = JSON.parse(text);"
	html$ = html$ + "        postTryResult.textContent = JSON.stringify(respObj, null, 2);"
	html$ = html$ + "        if(endpoint === '/createTable' && respObj && respObj.status === 'Command sent'){"
	html$ = html$ + "          if(respObj.table){ endpointFieldDefaults.tableName = String(respObj.table); }"
	html$ = html$ + "          window.sqlRefreshTableOptions();"
	html$ = html$ + "        }"
	html$ = html$ + "      }"
	html$ = html$ + "      catch(err) { postTryResult.textContent = text; }"
	html$ = html$ + "    })"
	html$ = html$ + "    .catch(function(err){ if(postTryResult){ postTryResult.textContent = 'Request failed: ' + err; } });"
	html$ = html$ + "};"
	html$ = html$ + "if(postTryEndpointSelect){ postTryEndpointSelect.addEventListener('change', function(){ window.sqlRenderPostTryFields(); window.sqlResetPostTryResult(); }); }"
	html$ = html$ + "window.sqlRefreshTableOptions = function(){"
	html$ = html$ + "  fetch(serverBase + '/tableCount')"
	html$ = html$ + "    .then(function(resp){ return resp.text(); })"
	html$ = html$ + "    .then(function(text){"
	html$ = html$ + "      try {"
	html$ = html$ + "        var obj = JSON.parse(text);"
	html$ = html$ + "        if(obj && obj.tables && obj.tables.length){"
	html$ = html$ + "          tableOptionsFromHelp = obj.tables;"
	html$ = html$ + "        }"
	html$ = html$ + "      } catch(err) {}"
	html$ = html$ + "      window.sqlRenderPostTryFields();"
	html$ = html$ + "    })"
	html$ = html$ + "    .catch(function(){ window.sqlRenderPostTryFields(); });"
	html$ = html$ + "};"
	html$ = html$ + "window.sqlRefreshTableOptions();"
	html$ = html$ + "window.sqlRenderPostTryFields();"
	html$ = html$ + "window.sqlSendUdpCommand = function(){"
	html$ = html$ + "  var input = document.getElementById('udpCommandInput');"
	html$ = html$ + "  var resultNode = document.getElementById('udpSendResult');"
	html$ = html$ + "  var udpCommand = input ? input.value.trim() : '';"
	html$ = html$ + "  if(udpCommand === ''){ if(resultNode){ resultNode.textContent = 'Please enter UDP command string.'; } return; }"
	html$ = html$ + "  if(resultNode){ resultNode.textContent = 'Sending...'; }"
	html$ = html$ + "  fetch('http://' + serverIp + ':8080/SendUDP', { method: 'POST', mode: 'no-cors', body: new URLSearchParams({ value: udpCommand }) })"
	html$ = html$ + "    .then(function(){ if(resultNode){ resultNode.textContent = 'Success'; } })"
	html$ = html$ + "    .catch(function(err){ if(resultNode){ resultNode.textContent = 'Send failed: ' + err; } });"
	html$ = html$ + "};"
	html$ = html$ + "</script>"
	html$ = html$ + "</body></html>"
	e.AddResponseHeader("Content-type", "text/html; charset=utf-8")
	e.SetResponseBodyString(html$)
	e.SendResponse(200)
End Sub

Sub sql_GetHomePage(userData as Object, e as Object)

	if type(userData) <> "roAssociativeArray" or type(userData.mVar) <> "roAssociativeArray" then
		e.AddResponseHeader("Content-type", "text/plain; charset=utf-8")
		e.SetResponseBodyString("SQL handler context is invalid")
		e.SendResponse(500)
		return
	end if

	mVar = userData.mVar
	mVar.tableNames = sql_GetTableNames(mVar.dbFile)
	requestedTable$ = sql_GetRequestedTableName(e)
	showTableWarning = false
	tableWarningTitle$ = ""
	tableWarningMessage$ = ""
	if requestedTable$ <> "" and sql_IsValidTableName(requestedTable$) and sql_ArrayContainsStringExact(mVar.tableNames, requestedTable$) then
		mVar.defaultTable = requestedTable$
	else if requestedTable$ <> "" then
		showTableWarning = true
		tableWarningTitle$ = "Table not found"
		tableWarningMessage$ = "The table: " + requestedTable$ + " does not exist. Showing the default table instead."
	else if sql_ArrayContainsStringExact(mVar.tableNames, "UserInfo") then
		mVar.defaultTable = "UserInfo"
	else
		sql_SyncDefaultTable(mVar)
	end if
	tableSelectOptions$ = ""
	if type(mVar.tableNames) = "roArray" then
		for each tableItem in mVar.tableNames
			tableItemName$ = sql_TrimString(sql_FormValueToString(tableItem))
			if tableItemName$ <> "" then
				selectedAttr$ = ""
				if tableItemName$ = mVar.defaultTable then
					selectedAttr$ = " selected"
				end if
				tableSelectOptions$ = tableSelectOptions$ + "<option value='" + sql_HtmlEscape(tableItemName$) + "'" + selectedAttr$ + ">" + sql_HtmlEscape(tableItemName$) + "</option>"
			end if
		next
	end if

	html$ = "<html>"
	html$ = html$ + "<head>"
	html$ = html$ + "<title>SQLite</title>"
	html$ = html$ + "<style>"
	html$ = html$ + "body { background-color: white; color: #111; font-family: Arial, sans-serif; padding-left: 10px; }"
	html$ = html$ + "input[type='text'], input[type='number'] { background-color: transparent; color: black; border: 2px solid #111; border-radius: 10px; padding: 10px; outline: none;}"
	html$ = html$ + "input[type='text']:focus, input[type='number']:focus { outline: none; }"
	html$ = html$ + "input[type='submit']  { background-color: #111; color: white; border: 2px solid #111; border-radius: 10px; padding: 10px; }"
	html$ = html$ + "button { background-color: #111; color: white; border: 2px solid #111; border-radius: 10px; padding: 10px; cursor: pointer; }"
	html$ = html$ + "select { background-color: transparent; color: black; border: 2px solid #111; border-radius: 10px; padding: 10px; }"
	html$ = html$ + "#sqlTableSelectWrap { width: 100%; }"
	html$ = html$ + "#sqlTableSelect { display: block; width: 100%; box-sizing: border-box; }"
	html$ = html$ + ".sql-modal { position: fixed; inset: 0; background: rgba(0,0,0,0.45); display: none; align-items: center; justify-content: center; z-index: 9999; }"
	html$ = html$ + ".sql-modal.is-open { display: flex; }"
	html$ = html$ + ".sql-modal-card { width: min(900px, 92vw); max-height: 82vh; background: #ffffff; border-radius: 14px; border: 1px solid #d7d7d7; box-shadow: 0 12px 32px rgba(0,0,0,0.25); padding: 14px; }"
	html$ = html$ + ".sql-modal-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; }"
	html$ = html$ + ".sql-modal-title { margin: 0; color: #3f3f3f; }"
	html$ = html$ + ".sql-json-box { margin: 0; max-height: 64vh; overflow: auto; border: 1px solid #d0d0d0; border-radius: 10px; padding: 12px; background: #f7f7f7; color: #1f1f1f; font-size: 13px; line-height: 1.45; white-space: pre-wrap; }"
	html$ = html$ + ".sql-message-box { margin: 0; padding: 12px; border: 1px solid #d0d0d0; border-radius: 10px; background: #f7f7f7; color: #1f1f1f; font-size: 14px; line-height: 1.5; white-space: pre-wrap; }"
	html$ = html$ + "</style>"
	html$ = html$ + "</head>"
	html$ = html$ + "<body>"
	html$ = html$ + "<h1>SQLite</h1>"
	html$ = html$ + "<span>Database file: <b>" + mVar.dbFile + "</b> - Active table: <b>" + mVar.defaultTable + "</b></span>"
	html$ = html$ + "<p><a href='/help'>Open Help Page</a></p>"

	if tableSelectOptions$ <> "" then
		html$ = html$ + "<p id='sqlTableSelectWrap'><label for='sqlTableSelect'><b>Tables:</b></label><BR><select id='sqlTableSelect' name='tableName'>" + tableSelectOptions$ + "</select></p>"
	else
		html$ = html$ + "<p><b>No tables found.</b></p>"
	end if
	html$ = html$ + "<div id='sqlMessageModal' class='sql-modal'><div class='sql-modal-card'><div class='sql-modal-head'><h3 class='sql-modal-title'>" + sql_HtmlEscape(tableWarningTitle$) + "</h3><button type='button' onclick='sqlHideMessagePopup()'>Close</button></div><pre id='sqlMessageBody' class='sql-message-box'></pre></div></div>"

	columns = sql_GetInsertableColumnsForTable(mVar, mVar.dbFile, mVar.defaultTable)
	rows = sql_ReadAllRows(mVar.dbFile, mVar.defaultTable)
	isSqlEmpty = sql_IsTableEffectivelyEmpty(rows, columns)
	if isSqlEmpty then
		html$ = html$ + "<p><b>" + mVar.defaultTable + "</b> is empty. Fill up all rows and click <b>UPDATE ALL</b>.</p>"
	end if
	if type(columns) = "roArray" and columns.Count() > 0 and type(rows) = "roArray" and rows.Count() > 0 then
		html$ = html$ + "<form id='sqlUpdateAllForm' method='POST' action='/update'>"
		html$ = html$ + "<input type='hidden' name='tableName' value='" + sql_HtmlEscape(mVar.defaultTable) + "' />"
		html$ = html$ + "<input type='hidden' name='updateAll' value='1' />"
		rowIdsCsv$ = ""
		for each row in rows
			if type(row) = "roAssociativeArray" then
				rowIdRaw$ = sql_TrimString(sql_GetRowIdValue(row, mVar.defaultRowID))
				rowIdDisplay$ = sql_HtmlEscape(rowIdRaw$)
				if rowIdRaw$ <> "" then
					if rowIdsCsv$ <> "" then
						rowIdsCsv$ = rowIdsCsv$ + ","
					end if
					rowIdsCsv$ = rowIdsCsv$ + rowIdRaw$
				end if
				html$ = html$ + "<div style='margin-bottom: 12px; padding: 8px; border: 1px solid #d0d0d0; border-radius: 8px;'>"
				if rowIdDisplay$ <> "" then
					html$ = html$ + "<p><b>" + mVar.defaultRowID + ":</b> " + rowIdDisplay$ + "</p>"
				end if
				for each columnName in columns
					fieldName$ = "bulk_" + rowIdRaw$ + "_" + columnName
					cellValue$ = sql_HtmlEscape(sql_GetRowFieldValue(row, columnName))
					html$ = html$ + "<label for='" + fieldName$ + "'>&nbsp;&nbsp;" + columnName + ":</label>"
					html$ = html$ + "<input type='text' id='" + fieldName$ + "' name='" + fieldName$ + "' value='" + cellValue$ + "' />"
				next
				html$ = html$ + "</div>"
			end if
		next
		html$ = html$ + "<input type='hidden' name='rowIds' value='" + sql_HtmlEscape(rowIdsCsv$) + "' />"
		html$ = html$ + "<p><input type='submit' value='UPDATE ALL' /></p>"
		html$ = html$ + "</form>"
	else if type(columns) = "roArray" and columns.Count() > 0 then
		html$ = html$ + "<p>No records in this table <b>" + mVar.defaultTable + "</b>. Insert new records to update.</p>"
	else
		html$ = html$ + "<p>No columns available for table <b>" + mVar.defaultTable + "</b>.</p>"
	end if

	html$ = html$ + "<div id='sqlResultModal' class='sql-modal' onclick='sqlCloseResultPopup(event)'>"
	html$ = html$ + "<div class='sql-modal-card'>"
	html$ = html$ + "<div class='sql-modal-head'>"
	html$ = html$ + "<h3 class='sql-modal-title'>UPDATE ALL</h3>"
	html$ = html$ + "<button type='button' onclick='sqlHideResultPopup()'>Close</button>"
	html$ = html$ + "</div>"
	html$ = html$ + "<pre id='sqlResultBody' class='sql-json-box'></pre>"
	html$ = html$ + "</div>"
	html$ = html$ + "</div>"
	html$ = html$ + "<script>"
	html$ = html$ + "(function(){"
	html$ = html$ + "var messageModal = document.getElementById('sqlMessageModal');"
	html$ = html$ + "var messageBody = document.getElementById('sqlMessageBody');"
	html$ = html$ + "var tableSelect = document.getElementById('sqlTableSelect');"
	html$ = html$ + "var form = document.getElementById('sqlUpdateAllForm');"
	html$ = html$ + "var modal = document.getElementById('sqlResultModal');"
	html$ = html$ + "var body = document.getElementById('sqlResultBody');"
	html$ = html$ + "window.sqlHideMessagePopup = function(){ if(messageModal){ messageModal.classList.remove('is-open'); } };"
	html$ = html$ + "window.sqlShowMessagePopup = function(message){ if(messageBody){ messageBody.textContent = message; } if(messageModal){ messageModal.classList.add('is-open'); } };"
	html$ = html$ + "window.sqlHideResultPopup = function(){ modal.classList.remove('is-open'); window.location.reload(); };"
	html$ = html$ + "window.sqlCloseResultPopup = function(ev){ if(ev.target === modal){ window.sqlHideResultPopup(); } };"

	html$ = html$ + "if(tableSelect){ tableSelect.addEventListener('change', function(){ var selectedTable = (tableSelect.value || '').trim(); if(selectedTable === ''){ return; } window.location.href = '/?tableName=' + encodeURIComponent(selectedTable); }); }"
	if showTableWarning then
		html$ = html$ + "window.sqlShowMessagePopup(" + Chr(34) + sql_HtmlEscape(tableWarningMessage$) + Chr(34) + ");"
	end if
	html$ = html$ + "if(!form){ return; }"
	html$ = html$ + "form.addEventListener('submit', function(ev){"
	html$ = html$ + "ev.preventDefault();"
	html$ = html$ + "body.textContent = 'Updating...';"
	html$ = html$ + "modal.classList.add('is-open');"
	html$ = html$ + "fetch(form.action, { method: 'POST', body: new FormData(form) })"
	html$ = html$ + ".then(function(resp){ return resp.text(); })"
	html$ = html$ + ".then(function(text){"
	html$ = html$ + "try { body.textContent = JSON.stringify(JSON.parse(text), null, 2); }"
	html$ = html$ + "catch (err) { body.textContent = text; }"
	html$ = html$ + "})"
	html$ = html$ + ".catch(function(err){ body.textContent = 'Update failed: ' + err; });"
	html$ = html$ + "});"
	html$ = html$ + "})();"
	html$ = html$ + "</script>"

	html$ = html$ + "</body>"
	html$ = html$ + "</html>"
	e.AddResponseHeader("Content-type", "text/html; charset=utf-8")
	e.SetResponseBodyString(html$)
	e.SendResponse(200)
End Sub


' ===== Database Read/Write Helpers =====
Function sql_ReadAllRows(dbFile$ as String, dbTable$ as String) as Object
	myDB = CreateObject("roSqliteDatabase")
	results = []
	if myDB.Open(dbFile$) <> invalid then
		selectStatement = myDB.CreateStatement("SELECT * FROM " + dbTable$ + ";")
		if type(selectStatement) = "roSqliteStatement" then
			sqlResult = selectStatement.Run()
			SQLITE_ROWS = 102
			while sqlResult = SQLITE_ROWS
				results.Push(selectStatement.GetData())
				sqlResult = selectStatement.Run()
			end while
			selectStatement.Finalise()
		end if
		myDB.Close()
	end if
	return results
End Function

Function sql_InsertUserData(m as Object, dbFile$ as String, dbTable$ as String, columns as Object, values as Object) as Boolean
	myDB = CreateObject("roSqliteDatabase")
	if type(columns) <> "roArray" or type(values) <> "roArray" then
		return false
	end if

	if columns.Count() = 0 or columns.Count() <> values.Count() then
		return false
	end if

	columnsSql$ = ""
	placeholders$ = ""
	for i = 0 to columns.Count() - 1
		if i > 0 then
			columnsSql$ = columnsSql$ + ", "
			placeholders$ = placeholders$ + ", "
		end if
		columnsSql$ = columnsSql$ + columns[i]
		placeholders$ = placeholders$ + "?"
	next

	if myDB.Open(dbFile$) <> invalid then
		sql$ = "INSERT INTO " + dbTable$ + " (" + columnsSql$ + ") VALUES (" + placeholders$ + ");"
		insertStatement = myDB.CreateStatement(sql$)
		success = sql_RunWriteStatement(insertStatement, values, m, "------------------ SQL bind error (insert)", "------------------ SQL insert error: ")
		myDB.Close()
		return success
	end if

	return false
End Function

Sub sql_UpdateRecordRest(userData as Object, e as Object)
	mVar = userData.mVar
	args = e.GetFormData()
	tableName$ = sql_GetFormValue(args, "tableName", mVar.defaultTable)
	rowID$ = sql_GetFormValue(args, mVar.defaultRowID, "")
	updateAll$ = LCase(sql_TrimString(sql_GetFormValue(args, "updateAll", "")))
	isUpdateAll = updateAll$ = "1" or updateAll$ = "true" or updateAll$ = "yes"
	allColumns = sql_GetInsertableColumnsForTable(mVar, mVar.dbFile, tableName$)
	columns = []
	values = []

	if not sql_IsValidTableName(tableName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid tableName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if

	if not isUpdateAll then
		if rowID$ = "" then
			sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "rowID is required" })
			return
		end if

		rowIdNum = sql_ParsePositiveIntStrict(rowID$)
		if rowIdNum <= 0 then
			sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "rowID must be a positive integer" })
			return
		end if
	end if

	if type(allColumns) <> "roArray" or allColumns.Count() = 0 then
		sql_SendJsonResponse(e, 400, { status: "error", message: "No updatable columns in table schema" })
		return
	end if

	ok = false
	if isUpdateAll then
		ok = sql_UpdateAllRowsFromForm(mVar, mVar.dbFile, tableName$, allColumns, args)
	else
		for each columnName in allColumns
			if type(args) = "roAssociativeArray" and args.DoesExist(columnName) and args[columnName] <> invalid then
				columns.Push(columnName)
				values.Push(sql_GetFormValue(args, columnName, ""))
			end if
		next

		if columns.Count() = 0 then
			sql_SendJsonResponse(e, 400, { status: "error", message: "Provide at least one field to update" })
			return
		end if

		ok = sql_UpdateRecord(mVar, mVar.dbFile, tableName$, rowIdNum, columns, values)
	end if

	respBody = {
		status: "error"
		dbFile: mVar.dbFile
		table: tableName$
		rowID: rowID$
	}
	if ok then
		mVar.defaultTable = tableName$
		respBody.status = "success"
		if isUpdateAll then
			respBody.message = "All records updated"
		else
			respBody.message = "Record updated"
		end if
		sql_SendJsonResponse(e, 200, respBody)
	else
		if isUpdateAll then
			respBody.message = "Update all failed"
			sql_SendJsonResponse(e, 500, respBody)
		else
			respBody.message = "Update failed"
			sql_SendJsonResponse(e, 404, respBody)
		end if
	end if
End Sub

Function sql_UpdateAllRowsFromForm(m as Object, dbFile$ as String, dbTable$ as String, allColumns as Object, args as Object) as Boolean
	if type(allColumns) <> "roArray" or allColumns.Count() = 0 then
		return false
	end if

	rowIdsCsv$ = sql_TrimString(sql_GetFormValue(args, "rowIds", ""))
	if rowIdsCsv$ = "" then
		return false
	end if

	rowIds = sql_SplitByDelimiter(rowIdsCsv$, ",")
	updatedCount = 0

	for each rowIdEntry in rowIds
		rowIdText$ = sql_TrimString(rowIdEntry)
		if rowIdText$ <> "" then
			rowIdNum = sql_ParsePositiveIntStrict(rowIdText$)
			if rowIdNum > 0 then
				existingRow = sql_ReadRowById(dbFile$, dbTable$, rowIdNum, sql_GetDefaultRowIdColumnName(m))
				if type(existingRow) <> "roAssociativeArray" then
					existingRow = {}
				end if

				rowValues = []
				for each columnName in allColumns
					fieldName$ = "bulk_" + rowIdText$ + "_" + columnName
					if type(args) = "roAssociativeArray" and args.DoesExist(fieldName$) and args[fieldName$] <> invalid then
						rowValues.Push(sql_GetFormValue(args, fieldName$, ""))
					else
						rowValues.Push(sql_GetRowFieldValue(existingRow, columnName))
					end if
				next

				if sql_UpdateRecord(m, dbFile$, dbTable$, rowIdNum, allColumns, rowValues) then
					updatedCount = updatedCount + 1
				end if
			end if
		end if
	next

	return updatedCount > 0
End Function

Function sql_GetRequestedTableName(e as Object) as String
	if e = invalid then
		return ""
	end if
	args = e.GetFormData()
	requestedTable$ = sql_GetRequestValue(e, args, "tableName", "")
	if requestedTable$ = "" then
		requestedTable$ = sql_GetRequestValue(e, args, "tablaName", "")
	end if
	return requestedTable$
End Function

Function sql_ArrayContainsStringExact(values as Object, target$ as String) as Boolean
	if type(values) <> "roArray" then
		return false
	end if

	if target$ = invalid or target$ = "" then
		return false
	end if

	for each item in values
		if type(item) = "roString" and item = target$ then
			return true
		end if
	next

	return false
End Function

Function sql_SplitByDelimiter(text$ as String, delimiter$ as String) as Object
	parts = []
	if text$ = invalid then
		return parts
	end if

	if delimiter$ = invalid or delimiter$ = "" then
		parts.Push(text$)
		return parts
	end if

	current$ = ""
	delLen = Len(delimiter$)
	i = 1
	while i <= Len(text$)
		if Mid(text$, i, delLen) = delimiter$ then
			parts.Push(current$)
			current$ = ""
			i = i + delLen
		else
			current$ = current$ + Mid(text$, i, 1)
			i = i + 1
		end if
	end while

	parts.Push(current$)
	return parts
End Function

Function sql_BuildValuesFromParts(parts as Object, startIndex as Integer, expectedCount as Integer) as Object
	values = []
	if type(parts) <> "roArray" or expectedCount <= 0 then
		return values
	end if

	for i = 0 to expectedCount - 1
		partIndex = startIndex + i
		if partIndex >= 0 and partIndex < parts.Count() and parts[partIndex] <> invalid then
			values.Push(sql_FormValueToString(parts[partIndex]))
		else
			values.Push("")
		end if
	next

	return values
End Function

Function sql_JoinPartsWithUnderscore(parts as Object, startIndex as Integer) as String
	if type(parts) <> "roArray" or startIndex < 0 or startIndex >= parts.Count() then
		return ""
	end if

	joined$ = ""
	for i = startIndex to parts.Count() - 1
		if i > startIndex then
			joined$ = joined$ + "_"
		end if
		joined$ = joined$ + sql_FormValueToString(parts[i])
	next

	return joined$
End Function

' ===== REST Read Handlers =====
Sub sql_PostReadRecordsRest(userData as Object, e as Object)
	mVar = userData.mVar
	args = e.GetFormData()
	tableName$ = sql_GetRequestValue(e, args, "tableName", mVar.defaultTable)
	if not sql_IsValidTableName(tableName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid tableName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if
	mVar.defaultTable = tableName$
	sql_SendReadResponse(mVar, tableName$, e)
End Sub

Sub sql_GetReadRowRest(userData as Object, e as Object)
	mVar = userData.mVar
	args = e.GetFormData()
	tableName$ = sql_GetRequestValue(e, args, "tableName", mVar.defaultTable)
	rowID$ = sql_GetRequestValue(e, args, mVar.defaultRowID, "")

	if not sql_IsValidTableName(tableName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid tableName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if

	if rowID$ = "" then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: mVar.defaultRowID + " is required" })
		return
	end if

	rowIdNum = sql_ParsePositiveIntStrict(rowID$)
	if rowIdNum <= 0 then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: mVar.defaultRowID + " must be a positive integer" })
		return
	end if

	row = sql_ReadRowById(mVar.dbFile, tableName$, rowIdNum, mVar.defaultRowID)
	if type(row) <> "roAssociativeArray" then
		sql_SendJsonResponse(e, 404, {
			status: "error"
			dbFile: mVar.dbFile
			table: tableName$
			rowID: rowID$
			message: "Row not found"
		})
		return
	end if

	mVar.defaultTable = tableName$
	sql_SendJsonResponse(e, 200, {
		status: "success"
		dbFile: mVar.dbFile
		table: tableName$
		row: row
	})
End Sub

Sub sql_GetReadColumnRest(userData as Object, e as Object)
	mVar = userData.mVar
	args = e.GetFormData()
	tableName$ = sql_GetRequestValue(e, args, "tableName", mVar.defaultTable)
	columnName$ = sql_GetRequestValue(e, args, "columnName", "")
	countOnlyRaw$ = LCase(sql_TrimString(sql_GetRequestValue(e, args, "countOnly", "")))
	isCountOnly = countOnlyRaw$ = "1" or countOnlyRaw$ = "true" or countOnlyRaw$ = "yes"
	rowID$ = sql_GetRequestValue(e, args, mVar.defaultRowID, "")

	if not sql_IsValidTableName(tableName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid tableName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if

	if not sql_IsValidTableName(columnName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid columnName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if

	if not sql_ColumnExists(mVar.dbFile, tableName$, columnName$) then
		sql_SendJsonResponse(e, 404, {
			status: "error"
			dbFile: mVar.dbFile
			table: tableName$
			column: columnName$
			message: "Column not found"
		})
		return
	end if

	if isCountOnly then
		count = sql_CountNonEmptyColumnValues(mVar.dbFile, tableName$, columnName$)
		if count < 0 then
			sql_SendJsonResponse(e, 500, {
				status: "error"
				dbFile: mVar.dbFile
				table: tableName$
				column: columnName$
				message: "Count failed"
			})
			return
		end if

		mVar.defaultTable = tableName$
		sql_SendJsonResponse(e, 200, {
			status: "success"
			dbFile: mVar.dbFile
			table: tableName$
			column: columnName$
			count: count
		})
		return
	end if

	if rowID$ <> "" then
		rowIdNum = sql_ParsePositiveIntStrict(rowID$)
		if rowIdNum <= 0 then
			sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: mVar.defaultRowID + " must be a positive integer" })
			return
		end if

		value = sql_ReadColumnValueByRowId(mVar.dbFile, tableName$, columnName$, rowIdNum, mVar.defaultRowID)
		if value = invalid then
			sql_SendJsonResponse(e, 404, {
				status: "error"
				dbFile: mVar.dbFile
				table: tableName$
				column: columnName$
				rowID: rowID$
				message: "Row not found"
			})
			return
		end if

		mVar.defaultTable = tableName$
		sql_SendJsonResponse(e, 200, {
			status: "success"
			dbFile: mVar.dbFile
			table: tableName$
			column: columnName$
			rowID: rowID$
			value: sql_FormValueToString(value)
		})
		return
	end if

	values = sql_ReadColumnValues(mVar.dbFile, tableName$, columnName$)
	mVar.defaultTable = tableName$
	sql_SendJsonResponse(e, 200, {
		status: "success"
		dbFile: mVar.dbFile
		table: tableName$
		column: columnName$
		values: values
	})
End Sub

Sub sql_GetSearchRowsRest(userData as Object, e as Object)
	mVar = userData.mVar
	args = e.GetFormData()
	tableName$ = sql_GetRequestValue(e, args, "tableName", mVar.defaultTable)
	columnName$ = sql_GetRequestValue(e, args, "columnName", "")
	value$ = sql_GetRequestValue(e, args, "value", "")

	if not sql_IsValidTableName(tableName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid tableName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if

	if not sql_IsValidTableName(columnName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid columnName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if

	if not sql_ColumnExists(mVar.dbFile, tableName$, columnName$) then
		sql_SendJsonResponse(e, 404, {
			status: "error"
			dbFile: mVar.dbFile
			table: tableName$
			column: columnName$
			message: "Column not found"
		})
		return
	end if

	rows = sql_ReadRowsByColumnExact(mVar.dbFile, tableName$, columnName$, value$)
	mVar.defaultTable = tableName$
	sql_SendJsonResponse(e, 200, {
		status: "success"
		dbFile: mVar.dbFile
		table: tableName$
		column: columnName$
		value: value$
		count: rows.Count()
		rows: rows
	})
End Sub

Sub sql_GetTablesRest(userData as Object, e as Object)
	mVar = userData.mVar
	tableNames = sql_GetTableNames(mVar.dbFile)
	if type(tableNames) <> "roArray" then
		tableNames = []
	end if

	mVar.tableNames = tableNames
	sql_SendJsonResponse(e, 200, {
		status: "success"
		dbFile: mVar.dbFile
		count: tableNames.Count()
		tables: tableNames
	})
End Sub

Sub sql_GetSchemaRest(userData as Object, e as Object)
	mVar = userData.mVar
	tableNames = sql_GetTableNames(mVar.dbFile)
	if type(tableNames) <> "roArray" then
		tableNames = []
	end if

	rowIdColumn$ = LCase(sql_GetDefaultRowIdColumnName(mVar))

	schema = []
	for each tableName in tableNames
		tableColumnsMeta = sql_GetTableColumnArray(mVar.dbFile, tableName)
		columnNames = []
		if type(tableColumnsMeta) = "roArray" then
			for each colMeta in tableColumnsMeta
				if type(colMeta) = "roAssociativeArray" and colMeta.DoesExist("name") and colMeta.name <> invalid then
					columnName$ = sql_TrimString(sql_FormValueToString(colMeta.name))
					if LCase(columnName$) <> rowIdColumn$ then
						columnNames.Push(columnName$)
					end if
				end if
			next
		end if

		columnCount = columnNames.Count()
		schema.Push({ table: tableName, columnCount: columnCount, columns: columnNames })
	next

	mVar.tableNames = tableNames
	schemaJson$ = "["
	for i = 0 to schema.Count() - 1
		if i > 0 then
			schemaJson$ = schemaJson$ + ","
		end if

		tableObj = schema[i]
		tableNameOut$ = ""
		columnCountOut = 0
		columnsOut = []
		if type(tableObj) = "roAssociativeArray" then
			if tableObj.DoesExist("table") and tableObj.table <> invalid then
				tableNameOut$ = sql_FormValueToString(tableObj.table)
			end if
			if tableObj.DoesExist("columnCount") and tableObj.columnCount <> invalid then
				columnCountOut = int(val(sql_FormValueToString(tableObj.columnCount)))
			end if
			if tableObj.DoesExist("columns") and type(tableObj.columns) = "roArray" then
				columnsOut = tableObj.columns
			end if
		end if

		columnsJson$ = "["
		for j = 0 to columnsOut.Count() - 1
			if j > 0 then
				columnsJson$ = columnsJson$ + ","
			end if
			columnsJson$ = columnsJson$ + Chr(34) + sql_FormValueToString(columnsOut[j]) + Chr(34)
		next
		columnsJson$ = columnsJson$ + "]"

		schemaJson$ = schemaJson$ + "{" + Chr(34) + "table" + Chr(34) + ":" + Chr(34) + tableNameOut$ + Chr(34) + "," + Chr(34) + "columnCount" + Chr(34) + ":" + columnCountOut.ToStr() + "," + Chr(34) + "columns" + Chr(34) + ":" + columnsJson$ + "}"
	next
	schemaJson$ = schemaJson$ + "]"

	body$ = "{" + Chr(34) + "status" + Chr(34) + ":" + Chr(34) + "success" + Chr(34) + "," + Chr(34) + "dbFile" + Chr(34) + ":" + Chr(34) + mVar.dbFile + Chr(34) + "," + Chr(34) + "countTable" + Chr(34) + ":" + tableNames.Count().ToStr() + "," + Chr(34) + "schema" + Chr(34) + ":" + schemaJson$ + "}"
	e.AddResponseHeader("Content-type", "application/json")
	e.SetResponseBodyString(body$)
	e.SendResponse(200)
End Sub

Function sql_ReadRowById(dbFile$ as String, dbTable$ as String, rowID as Integer, rowIdColumn$ as String) as Object
	myDB = CreateObject("roSqliteDatabase")
	if rowID <= 0 then
		return invalid
	end if

	if not sql_IsValidTableName(rowIdColumn$) then
		rowIdColumn$ = "rowID"
	end if

	if myDB.Open(dbFile$) = invalid then
		return invalid
	end if

	selectStatement = myDB.CreateStatement("SELECT * FROM " + dbTable$ + " WHERE " + rowIdColumn$ + " = ? LIMIT 1;")
	if type(selectStatement) <> "roSqliteStatement" then
		myDB.Close()
		return invalid
	end if

	params = [rowID]
	if not selectStatement.BindByOffset(params) then
		selectStatement.Finalise()
		myDB.Close()
		return invalid
	end if

	result = selectStatement.Run()
	if result = 102 then
		row = selectStatement.GetData()
	else
		row = invalid
	end if

	selectStatement.Finalise()
	myDB.Close()
	return row
End Function

Function sql_ReadColumnValues(dbFile$ as String, dbTable$ as String, columnName$ as String) as Object
	myDB = CreateObject("roSqliteDatabase")
	values = []

	if myDB.Open(dbFile$) = invalid then
		return values
	end if

	selectStatement = myDB.CreateStatement("SELECT " + columnName$ + " FROM " + dbTable$ + ";")
	if type(selectStatement) = "roSqliteStatement" then
		sqlResult = selectStatement.Run()
		while sqlResult = 102
			row = selectStatement.GetData()
			if type(row) = "roAssociativeArray" and row.DoesExist(columnName$) then
				values.Push(sql_FormValueToString(row[columnName$]))
			else if type(row) = "roAssociativeArray" then
				values.Push(sql_GetRowFieldValue(row, columnName$))
			else
				values.Push("")
			end if
			sqlResult = selectStatement.Run()
		end while
		selectStatement.Finalise()
	end if

	myDB.Close()
	return values
End Function

Function sql_ReadRowsByColumnExact(dbFile$ as String, dbTable$ as String, columnName$ as String, value$ as String) as Object
	myDB = CreateObject("roSqliteDatabase")
	rows = []

	if myDB.Open(dbFile$) = invalid then
		return rows
	end if

	selectStatement = myDB.CreateStatement("SELECT * FROM " + dbTable$ + " WHERE " + columnName$ + " = ?;")
	if type(selectStatement) <> "roSqliteStatement" then
		myDB.Close()
		return rows
	end if

	params = [value$]
	if not selectStatement.BindByOffset(params) then
		selectStatement.Finalise()
		myDB.Close()
		return rows
	end if

	sqlResult = selectStatement.Run()
	while sqlResult = 102
		rows.Push(selectStatement.GetData())
		sqlResult = selectStatement.Run()
	end while

	selectStatement.Finalise()
	myDB.Close()
	return rows
End Function

Function sql_ReadColumnValueByRowId(dbFile$ as String, dbTable$ as String, columnName$ as String, rowID as Integer, rowIdColumn$ as String) as Object
	myDB = CreateObject("roSqliteDatabase")

	if rowID <= 0 then
		return invalid
	end if

	if not sql_IsValidTableName(rowIdColumn$) then
		rowIdColumn$ = "rowID"
	end if

	if myDB.Open(dbFile$) = invalid then
		return invalid
	end if

	selectStatement = myDB.CreateStatement("SELECT " + columnName$ + " FROM " + dbTable$ + " WHERE " + rowIdColumn$ + " = ? LIMIT 1;")
	if type(selectStatement) <> "roSqliteStatement" then
		myDB.Close()
		return invalid
	end if

	params = [rowID]
	if not selectStatement.BindByOffset(params) then
		selectStatement.Finalise()
		myDB.Close()
		return invalid
	end if

	result = selectStatement.Run()
	if result = 102 then
		row = selectStatement.GetData()
		if type(row) = "roAssociativeArray" and row.DoesExist(columnName$) then
			value = row[columnName$]
		else if type(row) = "roAssociativeArray" then
			value = sql_GetRowFieldValue(row, columnName$)
		else
			value = invalid
		end if
	else
		value = invalid
	end if

	selectStatement.Finalise()
	myDB.Close()
	return value
End Function

Function sql_ColumnExists(dbFile$ as String, dbTable$ as String, columnName$ as String) as Boolean
	myDB = CreateObject("roSqliteDatabase")
	if myDB.Open(dbFile$) = invalid then
		return false
	end if

	selectStatement = myDB.CreateStatement("PRAGMA table_info(" + dbTable$ + ");")
	if type(selectStatement) <> "roSqliteStatement" then
		myDB.Close()
		return false
	end if

	found = false
	sqlResult = selectStatement.Run()
	while sqlResult = 102
		row = selectStatement.GetData()
		if type(row) = "roAssociativeArray" then
			name$ = sql_TrimString(sql_GetRowFieldValue(row, "name"))
			if LCase(name$) = LCase(columnName$) then
				found = true
				exit while
			end if
		end if
		sqlResult = selectStatement.Run()
	end while

	selectStatement.Finalise()
	myDB.Close()
	return found
End Function

Function sql_GetRequestValue(e as Object, args as Object, fieldName$ as String, defaultValue$ as String) as String
	if e <> invalid then
		queryValue$ = sql_TrimString(e.GetRequestParam(fieldName$))
		if queryValue$ <> "" then
			return queryValue$
		end if
	end if

	return sql_GetFormValue(args, fieldName$, defaultValue$)
End Function

Function sql_IsValidCreateColumnArray(columnArray as Object, requiredRowId$ as String) as Boolean
	validation = sql_ValidateCreateColumnArray(columnArray, requiredRowId$)
	if type(validation) = "roAssociativeArray" and validation.DoesExist("ok") and validation.ok = true then
		return true
	end if

	return false
End Function

Function sql_ValidateCreateColumnArray(columnArray as Object, requiredRowId$ as String) as Object
	result = {
		ok: false
		reason: ""
	}

	if type(columnArray) <> "roArray" or columnArray.Count() = 0 then
		result.reason = "columnArray must be a non-empty array"
		return result
	end if

	if requiredRowId$ = invalid or requiredRowId$ = "" then
		requiredRowId$ = "rowID"
	end if

	seenNames = {}
	for each col in columnArray
		if type(col) <> "roAssociativeArray" then
			result.reason = "each column must be an associative array"
			return result
		end if

		nameValue = sql_GetAssocValueIgnoreCase(col, "name")

		if nameValue = invalid then
			result.reason = "each column must include name"
			return result
		end if

		colName$ = sql_TrimString(sql_FormValueToString(nameValue))
		if colName$ = "" then
			result.reason = "column name cannot be empty"
			return result
		end if

		if not sql_IsValidTableName(colName$) then
			result.reason = "invalid column name: " + colName$
			return result
		end if

		nameKey$ = LCase(colName$)
		if seenNames.DoesExist(nameKey$) then
			result.reason = "duplicate column name: " + colName$
			return result
		end if
		seenNames[nameKey$] = true

	next

	result.ok = true
	result.reason = "ok"
	return result
End Function

Function sql_NormalizeCreateColumnArray(columnArray as Object, requiredRowId$ as String) as Object
	if requiredRowId$ = invalid or requiredRowId$ = "" then
		requiredRowId$ = "rowID"
	end if

	if not sql_IsValidCreateColumnArray(columnArray, requiredRowId$) then
		return invalid
	end if

	normalized = [{ name: requiredRowId$, type: "INTEGER PRIMARY KEY AUTOINCREMENT" }]
	for each col in columnArray
		nameValue = sql_GetAssocValueIgnoreCase(col, "name")
		colName$ = sql_TrimString(sql_FormValueToString(nameValue))
		if LCase(colName$) <> LCase(requiredRowId$) then
			normalized.Push({ name: colName$, type: "TEXT" })
		end if
	next

	return normalized
End Function

Function sql_GetAssocValueIgnoreCase(item as Object, keyName$ as String) as Object
	if type(item) <> "roAssociativeArray" then
		return invalid
	end if

	if item.DoesExist(keyName$) then
		return item[keyName$]
	end if

	target$ = LCase(sql_TrimString(keyName$))
	for each k in item
		if type(k) = "roString" and LCase(k) = target$ then
			return item[k]
		end if
	next

	return invalid
End Function

Function sql_GetCreateTableColumns(args as Object, defaultColumns as Object, requiredRowId$ as String) as Object
	normalizedDefault = sql_NormalizeCreateColumnArray(defaultColumns, requiredRowId$)
	if type(normalizedDefault) <> "roArray" then
		normalizedDefault = defaultColumns
	end if

	rawColumns = invalid
	if type(args) = "roAssociativeArray" and args.DoesExist("columnsJson") then
		rawColumns = args["columnsJson"]
	end if

	if rawColumns = invalid then
		return normalizedDefault
	end if

	columnsJson$ = ""
	if type(rawColumns) = "roArray" then
		if rawColumns.Count() = 0 then
			return normalizedDefault
		end if

		if type(rawColumns[0]) = "roAssociativeArray" then
			if sql_IsValidCreateColumnArray(rawColumns, requiredRowId$) then
				return sql_NormalizeCreateColumnArray(rawColumns, requiredRowId$)
			else
				return invalid
			end if
		end if

		columnsJson$ = sql_TrimString(sql_FormValueToString(rawColumns[0]))
	else
		columnsJson$ = sql_TrimString(sql_FormValueToString(rawColumns))
	end if

	if columnsJson$ = "" then
		return normalizedDefault
	end if

	parsedColumns = sql_ParseCreateColumnsJson(columnsJson$)
	if not sql_IsValidCreateColumnArray(parsedColumns, requiredRowId$) then
		return invalid
	end if

	return sql_NormalizeCreateColumnArray(parsedColumns, requiredRowId$)
End Function

Function sql_ParseCreateColumnsJson(columnsJson$ as String) as Object
	text$ = sql_TrimString(columnsJson$)
	if text$ = "" then
		return invalid
	end if

	parsed = ParseJson(text$)
	if type(parsed) = "roArray" then
		return parsed
	end if

	text$ = sql_ReplaceAll(text$, "&quot;", Chr(34))
	text$ = sql_ReplaceAll(text$, "&#34;", Chr(34))
	parsed = ParseJson(text$)
	if type(parsed) = "roArray" then
		return parsed
	end if

	text$ = sql_ReplaceAll(text$, Chr(92) + Chr(34), Chr(34))
	parsed = ParseJson(text$)
	if type(parsed) = "roArray" then
		return parsed
	end if

	if Len(text$) >= 2 and Left(text$, 1) = "'" and Right(text$, 1) = "'" then
		text$ = Mid(text$, 2, Len(text$) - 2)
		parsed = ParseJson(text$)
		if type(parsed) = "roArray" then
			return parsed
		end if
	end if

	return invalid
End Function

Function sql_ReplaceAll(text$ as String, find$ as String, replacement$ as String) as String
	if text$ = invalid then
		return ""
	end if

	if find$ = invalid or find$ = "" then
		return text$
	end if

	parts = sql_SplitByDelimiter(text$, find$)
	if type(parts) <> "roArray" or parts.Count() = 0 then
		return text$
	end if

	result$ = ""
	for i = 0 to parts.Count() - 1
		if i > 0 then
			result$ = result$ + replacement$
		end if
		result$ = result$ + parts[i]
	next

	return result$
End Function

' ===== REST Write Handlers =====
Sub sql_CreateTableRest(userData as Object, e as Object)

	mVar = userData.mVar
	args = e.GetFormData()

	tableName$ = sql_GetFormValue(args, "tableName", "")
	if tableName$ = "" then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "tableName is required" })
		return
	end if
	if not sql_IsValidTableName(tableName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid tableName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if
	columnsForCreate = sql_GetCreateTableColumns(args, mVar.columnArray, mVar.defaultRowID)
	if columnsForCreate = invalid then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid columnsJson. It must be a JSON array of {name} with unique valid names." })
		return
	end if

	createTableSQL$ = sql_BuildCreateTableSQL(tableName$, columnsForCreate)
	if createTableSQL$ = "" then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid columnArray schema. Check column names." })
		return
	end if
	ok = sql_CreateAndInitializeCustomDB(mVar, mVar.dbFile, createTableSQL$)
	respBody = {
		status: "error"
		dbFile: mVar.dbFile
		table: tableName$
	}
	if ok then
		mVar.tableNames = sql_GetTableNames(mVar.dbFile)
		mVar.columnArray = columnsForCreate
		respBody.status = "success"
		respBody.message = "Table is ready"
		sql_SendJsonResponse(e, 200, respBody)
	else
		respBody.message = "Unable to create table"
		sql_SendJsonResponse(e, 500, respBody)
	end if
End Sub

Sub sql_InsertRecordRest(userData as Object, e as Object)
	mVar = userData.mVar
	args = e.GetFormData()
	tableName$ = sql_GetFormValue(args, "tableName", mVar.defaultTable)
	columns = sql_GetInsertableColumnsForTable(mVar, mVar.dbFile, tableName$)
	values = []

	if not sql_IsValidTableName(tableName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid tableName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if

	if type(columns) <> "roArray" or columns.Count() = 0 then
		sql_SendJsonResponse(e, 400, { status: "error", message: "No insertable columns in table schema" })
		return
	end if

	for each columnName in columns
		value$ = sql_GetFormValue(args, columnName, "")
		values.Push(value$)
	next

	ok = sql_InsertUserData(mVar, mVar.dbFile, tableName$, columns, values)
	respBody = {
		status: "error"
		dbFile: mVar.dbFile
		table: tableName$
	}
	if ok then
		mVar.defaultTable = tableName$
		respBody.status = "success"
		respBody.message = "Record inserted"
		sql_SendJsonResponse(e, 200, respBody)
	else
		respBody.message = "Insert failed"
		sql_SendJsonResponse(e, 500, respBody)
	end if
End Sub

Sub sql_DeleteRecordRest(userData as Object, e as Object)

	mVar = userData.mVar
	args = e.GetFormData()

	tableName$ = sql_GetFormValue(args, "tableName", mVar.defaultTable)
	rowID$ = sql_GetFormValue(args, mVar.defaultRowID, "")

	if not sql_IsValidTableName(tableName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid tableName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if

	if rowID$ = "" then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "rowID is required" })
		return
	end if

	rowIdNum = sql_ParsePositiveIntStrict(rowID$)
	if rowIdNum <= 0 then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "rowID must be a positive integer" })
		return
	end if

	ok = sql_DeleteRecord(mVar, mVar.dbFile, tableName$, rowIdNum)
	respBody = {
		status: "error"
		dbFile: mVar.dbFile
		table: tableName$
		rowID: rowID$
	}
	if ok then
		mVar.defaultTable = tableName$
		respBody.status = "success"
		respBody.message = "Record deleted"
		sql_SendJsonResponse(e, 200, respBody)
	else
		respBody.message = "Delete failed"
		sql_SendJsonResponse(e, 404, respBody)
	end if
End Sub

Sub sql_DeleteTableRest(userData as Object, e as Object)
	mVar = userData.mVar
	args = e.GetFormData()

	tableName$ = sql_GetFormValue(args, "tableName", mVar.defaultTable)
	if not sql_IsValidTableName(tableName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid tableName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if

	ok = sql_DeleteTableRows(mVar, mVar.dbFile, tableName$)
	respBody = {
		status: "error"
		dbFile: mVar.dbFile
		table: tableName$
	}

	if ok then
		mVar.defaultTable = tableName$
		respBody.status = "success"
		respBody.message = "All rows deleted"
		sql_SendJsonResponse(e, 200, respBody)
	else
		respBody.message = "Delete all rows failed"
		sql_SendJsonResponse(e, 500, respBody)
	end if
End Sub

Sub sql_DeleteWhereRest(userData as Object, e as Object)
	mVar = userData.mVar
	args = e.GetFormData()

	tableName$ = sql_GetFormValue(args, "tableName", mVar.defaultTable)
	columnName$ = sql_GetFormValue(args, "columnName", "")
	value$ = sql_GetFormValue(args, "value", "")

	if not sql_IsValidTableName(tableName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid tableName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if

	if not sql_IsValidTableName(columnName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid columnName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if

	if LCase(columnName$) = LCase(sql_GetDefaultRowIdColumnName(mVar)) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Deleting by rowID is not supported in /deleteWhere. Use /delete." })
		return
	end if

	if not sql_ColumnExists(mVar.dbFile, tableName$, columnName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "columnName does not exist in table" })
		return
	end if

	if sql_TrimString(value$) = "" then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "value is required and cannot be empty" })
		return
	end if

	deletedCount = sql_DeleteWhere(mVar, mVar.dbFile, tableName$, columnName$, value$)
	if deletedCount < 0 then
		sql_SendJsonResponse(e, 500, {
			status: "error"
			dbFile: mVar.dbFile
			table: tableName$
			columnName: columnName$
			message: "Delete where failed"
		})
		return
	end if

	mVar.defaultTable = tableName$
	sql_SendJsonResponse(e, 200, {
		status: "success"
		dbFile: mVar.dbFile
		table: tableName$
		columnName: columnName$
		value: value$
		deletedCount: deletedCount
		message: "Rows deleted"
	})
End Sub

Sub sql_GetCountRest(userData as Object, e as Object)
	mVar = userData.mVar
	args = e.GetFormData()
	requestedTable$ = sql_GetRequestValue(e, args, "tableName", mVar.defaultTable)
	if sql_IsValidTableName(requestedTable$) then
		tableName$ = requestedTable$
	else
		tableName$ = mVar.defaultTable
	end if

	if not sql_IsValidTableName(tableName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid default table name" })
		return
	end if

	count = sql_CountRows(mVar.dbFile, tableName$)
	if count < 0 then
		sql_SendJsonResponse(e, 500, {
			status: "error"
			dbFile: mVar.dbFile
			table: tableName$
			message: "Count failed"
		})
		return
	end if

	mVar.defaultTable = tableName$
	sql_SendJsonResponse(e, 200, {
		status: "success"
		dbFile: mVar.dbFile
		table: tableName$
		count: count
	})
End Sub

Sub sql_BulkInsertRest(userData as Object, e as Object)
	mVar = userData.mVar
	args = e.GetFormData()

	tableName$ = sql_GetFormValue(args, "tableName", mVar.defaultTable)
	rowsJson$ = sql_GetFormValue(args, "rowsJson", "")

	if not sql_IsValidTableName(tableName$) then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "Invalid tableName. Use letters, numbers or underscore, and do not start with a number." })
		return
	end if

	if rowsJson$ = "" then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "rowsJson is required" })
		return
	end if

	rows = ParseJson(rowsJson$)
	if type(rows) <> "roArray" then
		sql_SendJsonResponse(e, 400, { status: "error", dbFile: mVar.dbFile, message: "rowsJson must be a JSON array" })
		return
	end if

	columns = sql_GetInsertableColumnsForTable(mVar, mVar.dbFile, tableName$)
	if type(columns) <> "roArray" or columns.Count() = 0 then
		sql_SendJsonResponse(e, 400, { status: "error", message: "No insertable columns in table schema" })
		return
	end if

	insertedCount = 0
	failedCount = 0

	for each row in rows
		if type(row) <> "roAssociativeArray" then
			failedCount = failedCount + 1
		else
			values = []
			for each columnName in columns
				cellValue$ = ""
				if row.DoesExist(columnName) and row[columnName] <> invalid then
					cellValue$ = sql_FormValueToString(row[columnName])
				end if
				values.Push(cellValue$)
			next

			if sql_InsertUserData(mVar, mVar.dbFile, tableName$, columns, values) then
				insertedCount = insertedCount + 1
			else
				failedCount = failedCount + 1
			end if
		end if
	next

	mVar.defaultTable = tableName$
	sql_SendJsonResponse(e, 200, {
		status: "success"
		dbFile: mVar.dbFile
		table: tableName$
		requestedCount: rows.Count()
		insertedCount: insertedCount
		failedCount: failedCount
	})
End Sub


' ===== Database Mutation and Counting =====
Function sql_DeleteRecord(m as Object, dbFile$ as String, dbTable$ as String, rowID as Integer) as Boolean
	myDB = CreateObject("roSqliteDatabase")
	if myDB.Open(dbFile$) = invalid then
		return false
	end if

	rowIdColumn$ = sql_GetDefaultRowIdColumnName(m)

	if not sql_RowExistsById(myDB, dbTable$, rowID, rowIdColumn$) then
		m.SystemLog.SendLine("------------------ Delete failed: rowID " + rowID.ToStr() + " does not exist.")
		myDB.Close()
		return false
	end if

	params = [rowID]
	deleteStatement = myDB.CreateStatement("DELETE FROM " + dbTable$ + " WHERE " + rowIdColumn$ + " = ?;")
	success = sql_RunWriteStatement(deleteStatement, params, m, "------------------ SQL bind error (delete)", "------------------ SQL delete error: ")

	if success then
		success = sql_SyncAutoincrementSequence(m, myDB, dbTable$, rowIdColumn$)
	end if

	myDB.Close()
	return success
End Function

Function sql_DeleteWhere(m as Object, dbFile$ as String, dbTable$ as String, columnName$ as String, value$ as String) as Integer
	myDB = CreateObject("roSqliteDatabase")
	if myDB.Open(dbFile$) = invalid then
		return -1
	end if

	countBefore = sql_CountRowsInOpenDb(myDB, dbTable$)
	if countBefore < 0 then
		myDB.Close()
		return -1
	end if

	deleteStatement = myDB.CreateStatement("DELETE FROM " + dbTable$ + " WHERE " + columnName$ + " = ?;")
	success = sql_RunWriteStatement(deleteStatement, [value$], m, "------------------ SQL bind error (deleteWhere)", "------------------ SQL deleteWhere error: ")
	if not success then
		myDB.Close()
		return -1
	end if

	rowIdColumn$ = sql_GetDefaultRowIdColumnName(m)
	if not sql_SyncAutoincrementSequence(m, myDB, dbTable$, rowIdColumn$) then
		myDB.Close()
		return -1
	end if

	countAfter = sql_CountRowsInOpenDb(myDB, dbTable$)
	myDB.Close()
	if countAfter < 0 then
		return -1
	end if

	deletedCount = countBefore - countAfter
	if deletedCount < 0 then
		deletedCount = 0
	end if

	return deletedCount
End Function

Function sql_CountRows(dbFile$ as String, dbTable$ as String) as Integer
	myDB = CreateObject("roSqliteDatabase")
	if myDB.Open(dbFile$) = invalid then
		return -1
	end if

	count = sql_CountRowsInOpenDb(myDB, dbTable$)
	myDB.Close()
	return count
End Function

Function sql_CountRowsInOpenDb(myDB as Object, dbTable$ as String) as Integer
	if type(myDB) <> "roSqliteDatabase" then
		return -1
	end if

	countStatement = myDB.CreateStatement("SELECT COUNT(*) AS rowCount FROM " + dbTable$ + ";")
	if type(countStatement) <> "roSqliteStatement" then
		return -1
	end if

	countValue = -1
	result = countStatement.Run()
	if result = 102 then
		row = countStatement.GetData()
		if type(row) = "roAssociativeArray" then
			countValue = int(val(sql_FormValueToString(sql_GetRowFieldValue(row, "rowCount"))))
			if countValue < 0 then
				countValue = 0
			end if
		end if
	end if

	countStatement.Finalise()
	return countValue
End Function

Function sql_CountNonEmptyColumnValues(dbFile$ as String, dbTable$ as String, columnName$ as String) as Integer
	myDB = CreateObject("roSqliteDatabase")
	if myDB.Open(dbFile$) = invalid then
		return -1
	end if

	countStatement = myDB.CreateStatement("SELECT COUNT(*) AS valueCount FROM " + dbTable$ + " WHERE " + columnName$ + " IS NOT NULL AND TRIM(" + columnName$ + ") <> '';")
	if type(countStatement) <> "roSqliteStatement" then
		myDB.Close()
		return -1
	end if

	countValue = -1
	result = countStatement.Run()
	if result = 102 then
		row = countStatement.GetData()
		if type(row) = "roAssociativeArray" then
			countValue = int(val(sql_FormValueToString(sql_GetRowFieldValue(row, "valueCount"))))
			if countValue < 0 then
				countValue = 0
			end if
		end if
	end if

	countStatement.Finalise()
	myDB.Close()
	return countValue
End Function

Function sql_SyncAutoincrementSequence(m as Object, myDB as Object, dbTable$ as String, rowIdColumn$ as String) as Boolean
	if type(myDB) <> "roSqliteDatabase" then
		return false
	end if

	maxId = 0
	maxIdStatement = myDB.CreateStatement("SELECT MAX(" + rowIdColumn$ + ") AS maxId FROM " + dbTable$ + ";")
	if type(maxIdStatement) = "roSqliteStatement" then
		maxResult = maxIdStatement.Run()
		if maxResult = 102 then
			maxRow = maxIdStatement.GetData()
			if type(maxRow) = "roAssociativeArray" and maxRow.maxId <> invalid then
				maxId = int(val(sql_FormValueToString(maxRow.maxId)))
				if maxId < 0 then
					maxId = 0
				end if
			end if
		end if
		maxIdStatement.Finalise()
	else
		return false
	end if

	if maxId <= 0 then
		clearSeqStatement = myDB.CreateStatement("DELETE FROM sqlite_sequence WHERE name = ?;")
		if type(clearSeqStatement) = "roSqliteStatement" then
			return sql_RunWriteStatement(clearSeqStatement, [dbTable$], m, "------------------ SQL bind error (clear rowID sequence after delete)", "------------------ SQL clear rowID sequence after delete error: ")
		end if
		return false
	end if

	setSeqStatement = myDB.CreateStatement("UPDATE sqlite_sequence SET seq = ? WHERE name = ?;")
	if type(setSeqStatement) = "roSqliteStatement" then
		setOk = sql_RunWriteStatement(setSeqStatement, [maxId, dbTable$], m, "------------------ SQL bind error (set rowID sequence after delete)", "------------------ SQL set rowID sequence after delete error: ")
		if setOk then
			return true
		end if
	end if

	insertSeqStatement = myDB.CreateStatement("INSERT INTO sqlite_sequence(name, seq) VALUES(?, ?);")
	if type(insertSeqStatement) <> "roSqliteStatement" then
		return false
	end if

	return sql_RunWriteStatement(insertSeqStatement, [dbTable$, maxId], m, "------------------ SQL bind error (insert rowID sequence)", "------------------ SQL insert rowID sequence error: ")
End Function

Function sql_DeleteTableRows(m as Object, dbFile$ as String, dbTable$ as String) as Boolean
	myDB = CreateObject("roSqliteDatabase")
	if myDB.Open(dbFile$) = invalid then
		return false
	end if

	deleteStatement = myDB.CreateStatement("DELETE FROM " + dbTable$ + ";")
	success = sql_RunWriteStatement(deleteStatement, [], m, "------------------ SQL bind error (delete all rows)", "------------------ SQL delete all rows error: ")

	if success then
		' Reset AUTOINCREMENT so next insert starts again from rowID 1.
		resetSeqStatement = myDB.CreateStatement("DELETE FROM sqlite_sequence WHERE name = ?;")
		if type(resetSeqStatement) = "roSqliteStatement" then
			resetOk = sql_RunWriteStatement(resetSeqStatement, [dbTable$], m, "------------------ SQL bind error (reset rowID sequence)", "------------------ SQL reset rowID sequence error: ")
			if not resetOk then
				success = false
			end if
		else
			success = false
		end if
	end if

	myDB.Close()
	return success
End Function

Function sql_UpdateRecord(m as Object, dbFile$ as String, dbTable$ as String, rowID as Integer, columns as Object, values as Object) as Boolean
	myDB = CreateObject("roSqliteDatabase")

	if type(columns) <> "roArray" or type(values) <> "roArray" then
		return false
	end if

	if columns.Count() = 0 or columns.Count() <> values.Count() then
		return false
	end if

	if rowID <= 0 then
		return false
	end if

	if myDB.Open(dbFile$) = invalid then
		return false
	end if

	rowIdColumn$ = sql_GetDefaultRowIdColumnName(m)

	if not sql_RowExistsById(myDB, dbTable$, rowID, rowIdColumn$) then
		myDB.Close()
		return false
	end if

	setClause$ = ""
	for i = 0 to columns.Count() - 1
		if i > 0 then
			setClause$ = setClause$ + ", "
		end if
		setClause$ = setClause$ + columns[i] + " = ?"
	next

	sql$ = "UPDATE " + dbTable$ + " SET " + setClause$ + " WHERE " + rowIdColumn$ + " = ?;"
	updateStatement = myDB.CreateStatement(sql$)
	bindParams = []
	for each v in values
		bindParams.Push(v)
	next
	bindParams.Push(rowID)

	success = sql_RunWriteStatement(updateStatement, bindParams, m, "------------------ SQL bind error (update)", "------------------ SQL update error: ")
	myDB.Close()
	return success
End Function

Function sql_RowExistsById(myDB as Object, dbTable$ as String, rowID as Integer, rowIdColumn$ = "rowID" as String) as Boolean
	if type(myDB) <> "roSqliteDatabase" or rowID <= 0 then
		return false
	end if

	if not sql_IsValidTableName(rowIdColumn$) then
		rowIdColumn$ = "rowID"
	end if

	selectStatement = myDB.CreateStatement("SELECT 1 FROM " + dbTable$ + " WHERE " + rowIdColumn$ + " = ? LIMIT 1;")
	if type(selectStatement) <> "roSqliteStatement" then
		return false
	end if

	params = [rowID]
	if not selectStatement.BindByOffset(params) then
		selectStatement.Finalise()
		return false
	end if

	result = selectStatement.Run()
	selectStatement.Finalise()
	return result = 102
End Function

Function sql_GetDefaultRowIdColumnName(m as Object) as String
	rowIdColumn$ = "rowID"
	if type(m) = "roAssociativeArray" and m.defaultRowID <> invalid and type(m.defaultRowID) = "roString" then
		candidate$ = sql_TrimString(m.defaultRowID)
		if sql_IsValidTableName(candidate$) then
			rowIdColumn$ = candidate$
		end if
	end if
	return rowIdColumn$
End Function

Function sql_RunWriteStatement(statement as Object, params as Object, m as Object, bindErrorMessage$ as String, runErrorPrefix$ as String) as Boolean
	if type(statement) <> "roSqliteStatement" then
		return false
	end if

	if not statement.BindByOffset(params) then
		if type(m) = "roAssociativeArray" and type(m.SystemLog) = "roSystemLog" then
			m.SystemLog.SendLine(bindErrorMessage$)
		end if
		statement.Finalise()
		return false
	end if

	result = statement.Run()
	statement.Finalise()
	if result = 100 then
		return true
	end if

	if type(m) = "roAssociativeArray" and type(m.SystemLog) = "roSystemLog" then
		m.SystemLog.SendLine(runErrorPrefix$ + result.ToStr())
	end if
	return false
End Function

Sub sql_SendReadResponse(mVar as Object, tableName$ as String, e as Object)
	respBody = {
		rows: sql_ReadAllRows(mVar.dbFile, tableName$)
	}
	e.AddResponseHeader("Content-type", "application/json")
	e.SetResponseBodyString(sql_PrettyJson(FormatJson(respBody), 2))
	e.SendResponse(200)
End Sub


Function sql_ResolveDbPath(dbFile$ as String) as String
	if dbFile$ = invalid or dbFile$ = "" then
		return "sd:/userData.db"
	end if

	if InStr(1, LCase(dbFile$), ":/") = 0 then
		return "sd:/" + dbFile$
	end if

	return dbFile$
End Function


Sub sql_EnsureDatabaseFile(m as Object)

	resolvedDb$ = sql_ResolveDbPath(m.dbFile)
	m.dbFile = resolvedDb$
	myDB = CreateObject("roSqliteDatabase")

	if not sql_CheckIfFileExists(resolvedDb$) then
		if myDB.Create(resolvedDb$) = invalid then
			m.SystemLog.SendLine("---------------------- Failed to create database file: " + resolvedDb$)
			return
		end if
		m.SystemLog.SendLine("---------------------- Created database file: " + resolvedDb$)

		createTableSQL$ = sql_BuildCreateTableSQL(m.defaultTable, m.columnArray)
		if createTableSQL$ = "" then
			m.SystemLog.SendLine("---------------------- Failed to build default table SQL from defaultTable/columnArray")
		else
			myDB.Close()

			if sql_CreateAndInitializeCustomDB(m, m.dbFile, createTableSQL$) then
				m.SystemLog.SendLine("---------------------- Default table is ready: " + m.defaultTable)
			else
				m.SystemLog.SendLine("---------------------- Failed to create default table: " + m.defaultTable)
			end if
			
			myDB = CreateObject("roSqliteDatabase")
			if myDB.Open(resolvedDb$) = invalid then
				m.SystemLog.SendLine("---------------------- Failed to reopen database file after table creation: " + resolvedDb$)
				return
			end if

		end if
	else
		if myDB.Open(resolvedDb$) = invalid then
			m.SystemLog.SendLine("---------------------- Failed to open database file: " + resolvedDb$)
			return
		end if
		m.SystemLog.SendLine("---------------------- Opened database file: " + resolvedDb$)
	end if

	myDB.Close()

End Sub


' ===== Form/Type Utilities =====
Function sql_GetFormValue(args as Object, fieldName$ as String, defaultValue$ as String) as String
	if type(args) = "roAssociativeArray" and args.DoesExist(fieldName$) and args[fieldName$] <> invalid then
		return sql_FormValueToString(args[fieldName$])
	end if
	return defaultValue$
End Function

Function sql_FormValueToString(value as Object) as String
	if value = invalid then
		return ""
	end if

	vType$ = type(value)
	if vType$ = "roString" then
		return value
	end if

	if vType$ = "Integer" or vType$ = "LongInteger" or vType$ = "Float" or vType$ = "Double" or vType$ = "Boolean" then
		return value.ToStr()
	end if

	if vType$ = "roInt" or vType$ = "roInteger" or vType$ = "roLongInteger" or vType$ = "roFloat" or vType$ = "roDouble" or vType$ = "roBoolean" then
		return value.ToStr()
	end if

	if vType$ = "roArray" then
		if value.Count() = 0 or value[0] = invalid then
			return ""
		end if
		return sql_FormValueToString(value[0])
	end if

	if GetInterface(value, "ifString") <> invalid then
		return value.ToStr()
	end if

	if GetInterface(value, "ifInt") <> invalid then
		return value.ToStr()
	end if

	if GetInterface(value, "ifFloat") <> invalid then
		return value.ToStr()
	end if

	return ""
End Function

Function sql_HtmlEscape(value$ as String) as String
	if value$ = invalid or value$ = "" then
		return ""
	end if

	encoded$ = ""
	for i = 1 to Len(value$)
		ch$ = Mid(value$, i, 1)
		if ch$ = "&" then
			encoded$ = encoded$ + "&amp;"
		else if ch$ = "<" then
			encoded$ = encoded$ + "&lt;"
		else if ch$ = ">" then
			encoded$ = encoded$ + "&gt;"
		else if ch$ = Chr(34) then
			encoded$ = encoded$ + "&quot;"
		else if ch$ = "'" then
			encoded$ = encoded$ + "&#39;"
		else
			encoded$ = encoded$ + ch$
		end if
	next

	return encoded$
End Function

Function sql_GetRowIdValue(row as Object, rowKey$ as String) as String
	if type(row) <> "roAssociativeArray" then
		return ""
	end if

	if row.DoesExist(rowKey$) and row[rowKey$] <> invalid then
		return sql_FormValueToString(row[rowKey$])
	end if

	for each key in row
		if LCase(key) = LCase(rowKey$) and row[key] <> invalid then
			return sql_FormValueToString(row[key])
		end if
	next

	return ""
End Function

Function sql_GetRowFieldValue(row as Object, fieldName$ as String) as String
	if type(row) <> "roAssociativeArray" then
		return ""
	end if

	if row.DoesExist(fieldName$) and row[fieldName$] <> invalid then
		return sql_FormValueToString(row[fieldName$])
	end if

	for each key in row
		if LCase(key) = LCase(fieldName$) and row[key] <> invalid then
			return sql_FormValueToString(row[key])
		end if
	next

	return ""
End Function

Function sql_IsTableEffectivelyEmpty(rows as Object, columns as Object) as Boolean
	if type(rows) <> "roArray" or type(columns) <> "roArray" then
		return false
	end if

	if rows.Count() = 0 or columns.Count() = 0 then
		return false
	end if

	for each row in rows
		if type(row) = "roAssociativeArray" then
			for each columnName in columns
				if sql_TrimString(sql_GetRowFieldValue(row, columnName)) <> "" then
					return false
				end if
			next
		end if
	next

	return true
End Function

Function sql_TrimString(value$ as String) as String
	if value$ = invalid then
		return ""
	end if

	startIdx = 1
	endIdx = Len(value$)

	while startIdx <= endIdx
		ch$ = Mid(value$, startIdx, 1)
		if ch$ <> " " and ch$ <> Chr(9) and ch$ <> Chr(10) and ch$ <> Chr(13) then
			exit while
		end if
		startIdx = startIdx + 1
	end while

	while endIdx >= startIdx
		ch$ = Mid(value$, endIdx, 1)
		if ch$ <> " " and ch$ <> Chr(9) and ch$ <> Chr(10) and ch$ <> Chr(13) then
			exit while
		end if
		endIdx = endIdx - 1
	end while

	if startIdx > endIdx then
		return ""
	end if

	return Mid(value$, startIdx, endIdx - startIdx + 1)
End Function

Function sql_GetTableNames(dbFile$ as String) as Object
	myDB = CreateObject("roSqliteDatabase")
	tableNames = []
	if myDB.Open(dbFile$) = invalid then
		if myDB.Create(dbFile$) = invalid then
			return tableNames
		end if
		if myDB.Open(dbFile$) = invalid then
			return tableNames
		end if
	end if

	selectStatement = myDB.CreateStatement("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;")
	if type(selectStatement) = "roSqliteStatement" then
		sqlResult = selectStatement.Run()
		SQLITE_ROWS = 102
		while sqlResult = SQLITE_ROWS
			row = selectStatement.GetData()
			if type(row) = "roAssociativeArray" and row.name <> invalid and not sql_IsInternalTable(row.name) then
				tableNames.Push(row.name)
			end if
			sqlResult = selectStatement.Run()
		end while
		selectStatement.Finalise()
	end if
	myDB.Close()
	return tableNames
End Function

Sub sql_SyncDefaultTable(m as Object)
	if type(m.tableNames) <> "roArray" then
		m.tableNames = []
	end if

	if m.tableNames.Count() = 0 then
		m.defaultTable = "UserInfo"
		return
	end if

	for each tableName in m.tableNames
		if tableName = m.defaultTable then
			return
		end if
	next

	m.defaultTable = m.tableNames[0]
End Sub

Function sql_GetInsertableColumns(columnArray as Object) as Object
	columns = []
	if type(columnArray) <> "roArray" then
		return columns
	end if

	for each col in columnArray
		if type(col) = "roAssociativeArray" then
			nameValue = sql_GetAssocValueIgnoreCase(col, "name")
			typeValue = sql_GetAssocValueIgnoreCase(col, "type")
			colName$ = sql_TrimString(sql_FormValueToString(nameValue))
			colType$ = sql_TrimString(sql_FormValueToString(typeValue))

			if colName$ <> "" and colType$ <> "" then
				colTypeUpper$ = UCase(colType$)
				if Instr(1, colTypeUpper$, "PRIMARY KEY") = 0 and Instr(1, colTypeUpper$, "AUTOINCREMENT") = 0 then
					columns.Push(colName$)
				end if
			end if
		end if
	next

	return columns
End Function

Function sql_GetTableColumnArray(dbFile$ as String, tableName$ as String) as Object
	columns = []
	if not sql_IsValidTableName(tableName$) then
		return columns
	end if

	myDB = CreateObject("roSqliteDatabase")
	if myDB.Open(dbFile$) = invalid then
		return columns
	end if

	selectStatement = myDB.CreateStatement("PRAGMA table_info(" + tableName$ + ");")
	if type(selectStatement) <> "roSqliteStatement" then
		myDB.Close()
		return columns
	end if

	sqlResult = selectStatement.Run()
	while sqlResult = 102
		row = selectStatement.GetData()
		if type(row) = "roAssociativeArray" then
			colName$ = sql_TrimString(sql_GetRowFieldValue(row, "name"))
			colType$ = sql_TrimString(sql_GetRowFieldValue(row, "type"))
			pkValue = int(val(sql_GetRowFieldValue(row, "pk")))

			if colName$ <> "" and sql_IsValidTableName(colName$) then
				if colType$ = "" then
					colType$ = "TEXT"
				end if

				colTypeUpper$ = UCase(colType$)
				if pkValue > 0 and Instr(1, colTypeUpper$, "PRIMARY KEY") = 0 then
					colType$ = colType$ + " PRIMARY KEY"
					if LCase(colName$) = "rowid" and Instr(1, colTypeUpper$, "INT") > 0 and Instr(1, UCase(colType$), "AUTOINCREMENT") = 0 then
						colType$ = colType$ + " AUTOINCREMENT"
					end if
				end if

				columns.Push({ name: colName$, type: colType$ })
			end if
		end if
		sqlResult = selectStatement.Run()
	end while

	selectStatement.Finalise()
	myDB.Close()
	return columns
End Function

Function sql_GetInsertableColumnsForTable(m as Object, dbFile$ as String, tableName$ as String) as Object
	tableColumns = sql_GetTableColumnArray(dbFile$, tableName$)
	columns = sql_GetInsertableColumns(tableColumns)
	if type(columns) = "roArray" and columns.Count() > 0 then
		return columns
	end if

	if type(m) = "roAssociativeArray" and type(m.columnArray) = "roArray" then
		return sql_GetInsertableColumns(m.columnArray)
	end if

	return []
End Function

Function sql_IsInternalTable(tableName$ as String) as Boolean
	if tableName$ = invalid or tableName$ = "" then
		return true
	end if
	if Left(tableName$, 7) = "sqlite_" then
		return true
	end if
	return false
End Function


Function sql_IsValidTableName(tableName$ as String) as Boolean
	if tableName$ = invalid or tableName$ = "" then
		return false
	end if

	firstChar$ = Mid(tableName$, 1, 1)
	if not sql_IsLetterOrUnderscore(firstChar$) then
		return false
	end if

	for i = 1 to Len(tableName$)
		ch$ = Mid(tableName$, i, 1)
		if not sql_IsLetterOrDigitOrUnderscore(ch$) then
			return false
		end if
	next

	return true
End Function

Function sql_IsLetterOrUnderscore(ch$ as String) as Boolean
	if ch$ = "_" then
		return true
	end if
	code = Asc(ch$)
	if (code >= 65 and code <= 90) or (code >= 97 and code <= 122) then
		return true
	end if
	return false
End Function

Function sql_ParsePositiveIntStrict(value$ as String) as Integer
	text$ = sql_TrimString(value$)
	if text$ = "" then
		return 0
	end if

	num = int(val(text$))
	if num <= 0 or num.ToStr() <> text$ then
		return 0
	end if

	return num
End Function

Function sql_IsLetterOrDigitOrUnderscore(ch$ as String) as Boolean
	if ch$ = "_" then
		return true
	end if
	code = Asc(ch$)
	if (code >= 65 and code <= 90) or (code >= 97 and code <= 122) or (code >= 48 and code <= 57) then
		return true
	end if
	return false
End Function

Function sql_BuildCreateTableSQL(tableName$ as String, columnArray as Object) as String
	if not sql_IsValidTableName(tableName$) then
		return ""
	end if

	if type(columnArray) <> "roArray" or columnArray.Count() = 0 then
		return ""
	end if

	columnsSql$ = ""
	for each col in columnArray
		if type(col) <> "roAssociativeArray" then
			return ""
		end if

		nameValue = sql_GetAssocValueIgnoreCase(col, "name")
		typeValue = sql_GetAssocValueIgnoreCase(col, "type")
		if nameValue = invalid or typeValue = invalid then
			return ""
		end if

		colName$ = sql_TrimString(sql_FormValueToString(nameValue))
		colType$ = sql_TrimString(sql_FormValueToString(typeValue))
		if not sql_IsValidTableName(colName$) then
			return ""
		end if
		if colType$ = "" then
			return ""
		end if

		if columnsSql$ <> "" then
			columnsSql$ = columnsSql$ + ", "
		end if
		columnsSql$ = columnsSql$ + colName$ + " " + colType$
	next

	return "CREATE TABLE IF NOT EXISTS " + tableName$ + " (" + columnsSql$ + ");"
End Function


' ===== Error Mapping and JSON Response =====
Function sql_GetErrorCodeFromMessage(message$ as String, statusCode as Integer) as String
	msg$ = LCase(sql_TrimString(message$))

	if Instr(1, msg$, "invalid tablename") > 0 then
		return "INVALID_TABLE_NAME"
	else if Instr(1, msg$, "invalid columnname") > 0 then
		return "INVALID_COLUMN_NAME"
	else if Instr(1, msg$, "rowid is required") > 0 then
		return "MISSING_ROW_ID"
	else if Instr(1, msg$, "rowid must be a positive integer") > 0 then
		return "INVALID_ROW_ID"
	else if Instr(1, msg$, "row not found") > 0 then
		return "ROW_NOT_FOUND"
	else if Instr(1, msg$, "column not found") > 0 then
		return "COLUMN_NOT_FOUND"
	else if Instr(1, msg$, "rowsjson is required") > 0 then
		return "MISSING_ROWS_JSON"
	else if Instr(1, msg$, "rowsjson must be a json array") > 0 then
		return "INVALID_ROWS_JSON"
	else if Instr(1, msg$, "invalid columnsjson") > 0 then
		return "INVALID_COLUMNS_JSON"
	else if Instr(1, msg$, "invalid columnarray schema") > 0 then
		return "INVALID_COLUMN_ARRAY"
	else if Instr(1, msg$, "provide at least one field to update") > 0 then
		return "NO_UPDATE_FIELDS"
	else if Instr(1, msg$, "no insertable columns") > 0 then
		return "NO_INSERTABLE_COLUMNS"
	else if Instr(1, msg$, "no updatable columns") > 0 then
		return "NO_UPDATABLE_COLUMNS"
	else if Instr(1, msg$, "count failed") > 0 then
		return "COUNT_FAILED"
	else if Instr(1, msg$, "update failed") > 0 then
		return "UPDATE_FAILED"
	else if Instr(1, msg$, "update all failed") > 0 then
		return "UPDATE_ALL_FAILED"
	else if Instr(1, msg$, "insert failed") > 0 then
		return "INSERT_FAILED"
	else if Instr(1, msg$, "delete failed") > 0 then
		return "DELETE_FAILED"
	else if Instr(1, msg$, "delete where failed") > 0 then
		return "DELETE_WHERE_FAILED"
	else if Instr(1, msg$, "delete all rows failed") > 0 then
		return "DELETE_ALL_ROWS_FAILED"
	else if Instr(1, msg$, "unable to create table") > 0 then
		return "CREATE_TABLE_FAILED"
	end if

	if statusCode = 400 then
		return "BAD_REQUEST"
	else if statusCode = 401 then
		return "UNAUTHORIZED"
	else if statusCode = 403 then
		return "FORBIDDEN"
	else if statusCode = 404 then
		return "NOT_FOUND"
	else if statusCode >= 500 then
		return "INTERNAL_ERROR"
	end if

	return "UNSPECIFIED_ERROR"
End Function

Sub sql_EnsureErrorCode(body as Object, statusCode as Integer)
	if type(body) <> "roAssociativeArray" then
		return
	end if

	if not body.DoesExist("status") then
		return
	end if

	status$ = LCase(sql_TrimString(sql_FormValueToString(body.status)))
	if status$ <> "error" then
		return
	end if

	if body.DoesExist("code") and sql_TrimString(sql_FormValueToString(body.code)) <> "" then
		return
	end if

	message$ = ""
	if body.DoesExist("message") then
		message$ = sql_FormValueToString(body.message)
	end if
	body.code = sql_GetErrorCodeFromMessage(message$, statusCode)
End Sub

Sub sql_SendJsonResponse(e as Object, statusCode as Integer, body as Object)
	sql_EnsureErrorCode(body, statusCode)
	e.AddResponseHeader("Content-type", "application/json")
	e.SetResponseBodyString(FormatJson(body))
	e.SendResponse(statusCode)
End Sub

Function sql_PrettyJson(json$ as String, indentSize as Integer) as String
	if json$ = invalid then
		return ""
	end if

	if indentSize <= 0 then
		indentSize = 2
	end if

	out$ = ""
	indentLevel = 0
	inString = false
	escapeNext = false

	for i = 1 to Len(json$)
		ch$ = Mid(json$, i, 1)

		if inString then
			out$ = out$ + ch$
			if escapeNext then
				escapeNext = false
			else if ch$ = "\\" then
				escapeNext = true
			else if ch$ = Chr(34) then
				inString = false
			end if
		else
			if ch$ = Chr(34) then
				inString = true
				out$ = out$ + ch$
			else if ch$ = "{" or ch$ = "[" then
				indentLevel = indentLevel + 1
				out$ = out$ + ch$ + Chr(10) + sql_RepeatSpace(indentLevel * indentSize)
			else if ch$ = "}" or ch$ = "]" then
				indentLevel = indentLevel - 1
				if indentLevel < 0 then indentLevel = 0
				out$ = out$ + Chr(10) + sql_RepeatSpace(indentLevel * indentSize) + ch$
			else if ch$ = "," then
				out$ = out$ + ch$ + Chr(10) + sql_RepeatSpace(indentLevel * indentSize)
			else if ch$ = ":" then
				out$ = out$ + ": "
			else if ch$ <> " " and ch$ <> Chr(9) and ch$ <> Chr(10) and ch$ <> Chr(13) then
				out$ = out$ + ch$
			end if
		end if
	next

	return out$
End Function

Function sql_RepeatSpace(count as Integer) as String
	if count <= 0 then
		return ""
	end if

	spaces$ = ""
	for i = 1 to count
		spaces$ = spaces$ + " "
	next

	return spaces$
End Function


Function sql_CheckIfFileExists(filePath as String) as Boolean
	resolvedPath$ = sql_ResolveDbPath(filePath)
	retval = false

	lastSlashPos = 0
	for i = 1 to Len(resolvedPath$)
		if Mid(resolvedPath$, i, 1) = "/" then
			lastSlashPos = i
		end if
	next

	if lastSlashPos <= 0 then
		return false
	end if

	dirPath$ = Left(resolvedPath$, lastSlashPos)
	fileName$ = Mid(resolvedPath$, lastSlashPos + 1)
	targetFile$ = LCase(fileName$)

	if fileName$ = "" then
		return false
	end if


	dirList = ListDir(dirPath$)
	if dirList <> invalid then
		for each file in dirList
			if LCase(file) = targetFile$ then
				retval = true
				exit for
			end if
		next
	end if

	return retval

End Function

Sub EnableUDPNotifications(m as object)
	if m.bsp <> invalid and type(m.bsp) = "roAssociativeArray" then
		if type(m.bsp.CreateDatagramReceiver) = "roFunction" then
			if m.bsp.udpReceivePort <> invalid then
				m.bsp.CreateDatagramReceiver(m.bsp.udpReceivePort)
				m.bsp.diagnostics.PrintDebug("[EnableUDPNotifications] udpReceivePort already set: " + m.bsp.udpReceivePort.ToStr())
			else
				m.bsp.CreateDatagramReceiver(m.bsp.udpNotificationPort%)
				m.bsp.diagnostics.PrintDebug("[EnableUDPNotifications] Creating DatagramReceiver on port: " + m.bsp.udpNotificationPort%.ToStr())
			end if
		end if
	end if
End Sub
