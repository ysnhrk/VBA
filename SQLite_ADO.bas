'Connect SQLite3 with ADO
  'Install Required: ODBC driver for SQLite3
Sub Test_SQLite_ADO_BigData()

    Dim con: Set con = CreateObject("ADODB.Connection") 'ADOコネクション
    Dim rs: Set rs = CreateObject("ADODB.Recordset")    'ADOレコードセット
    Dim sql As String                                   ’SQL文
    Dim db As String                                    'SQLiteのDBファイル名
    db = "C:\Users\USERNAME\PATH\test.db"
    
    '接続
    con.ConnectionString = "DRIVER=SQLite3 ODBC Driver;Database=" & db
    con.Open
        
'    'テーブル作成
'    sql = "CREATE TABLE Music ('Title','Genre','Composer') ;"
'    Call con.Execute(sql)
'
'    'INSERT
    Dim i As Long, title As String, composer As String
'    For i = 1 To 10000
'        title = "symphony_" & i
'        Select Case (i Mod 3)
'            Case 1: composer = "Mozart"
'            Case 2: composer = "Beetoven"
'            Case Else: composer = "Ravel"
'        End Select
'        sql = "INSERT INTO Music ('Title','Genre','Composer') VALUES ('" & _
'              title & "','Classical','" & _
'              composer & "');"
'        Call con.Execute(sql)
'    Next i
    
    'SELECT
    sql = "SELECT * FROM Music;"    '10000件で0.01秒
    '    sql = "SELECT Composer, COUNT(Composer) FROM Music GROUP BY Composer;" '10000件で0.007秒
    Set rs = con.Execute(sql)
    '    Sheet1.[A1].CopyFromRecordset rs
    
    '終了処理
    rs.Close: Set rs = Nothing                    'レコードセットの破棄
    con.Close: Set con = Nothing                  'コネクションの破棄

End Sub
