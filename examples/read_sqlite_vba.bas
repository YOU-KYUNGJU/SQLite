Option Explicit

Sub ReadShrinkageSQLite()
    Dim conn As Object
    Dim rs As Object
    Dim dbPath As String
    Dim sql As String
    Dim rowIndex As Long

    dbPath = "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\DB\접수현황_수축률DB\접수현황_수축률DB_2026.db"
    sql = "SELECT part_name, receipt_no, received_date, test_item, status, client_name " & _
          "FROM receipt_status " & _
          "WHERE part_code = 'shrinkage' " & _
          "ORDER BY received_date DESC " & _
          "LIMIT 100;"

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Driver=SQLite3 ODBC Driver;Database=" & dbPath & ";"

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn

    rowIndex = 1
    Do Until rs.EOF
        Cells(rowIndex, 1).Value = rs.Fields("part_name").Value
        Cells(rowIndex, 2).Value = rs.Fields("receipt_no").Value
        Cells(rowIndex, 3).Value = rs.Fields("received_date").Value
        Cells(rowIndex, 4).Value = rs.Fields("test_item").Value
        Cells(rowIndex, 5).Value = rs.Fields("status").Value
        Cells(rowIndex, 6).Value = rs.Fields("client_name").Value
        rowIndex = rowIndex + 1
        rs.MoveNext
    Loop

    rs.Close
    conn.Close
End Sub
