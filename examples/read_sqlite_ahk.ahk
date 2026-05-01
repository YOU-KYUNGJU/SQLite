#NoEnv
SetBatchLines, -1

dbPath := "\\192.168.1.7\유해물질시험팀\3. 폼알데히드,pH파트\자동화프로그램 개발\DB\접수현황_수축률DB\접수현황_수축률DB_2026.db"
sql := "SELECT part_name, receipt_no, received_date, test_item FROM receipt_status WHERE part_code = 'shrinkage' LIMIT 10;"

conn := ComObjCreate("ADODB.Connection")
conn.Open("Driver=SQLite3 ODBC Driver;Database=" dbPath ";")

rs := conn.Execute(sql)
result := ""
while !rs.EOF
{
    result .= rs.Fields("part_name").Value " | "
    result .= rs.Fields("receipt_no").Value " | "
    result .= rs.Fields("received_date").Value " | "
    result .= rs.Fields("test_item").Value "`n"
    rs.MoveNext()
}

rs.Close()
conn.Close()

MsgBox, % result
