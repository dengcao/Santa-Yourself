<%set conn=server.CreateObject("adodb.connection")
'连接SQL  conn.connectionstring="driver={sql server};server=127.0.0.1;uid=sa;pwd=123456;database=chinaz_user"
'conn.open
db_path="data/sadan-user#2008-12-24.mdb"
Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.MapPath(db_path)
Conn.Open Connstr%>