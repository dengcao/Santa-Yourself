<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
	<%'读取用户
dim ErrTrue,alerttext
set rs=server.CreateObject("adodb.recordset")
sql="SELECT * FROM Sanda_User WHERE ID="&Trim(Request("user"))
rs.open sql,conn,1,3
rs("photo_scale")=Request("photo_scale")
rs("photo_rotate")=Request("photo_rotate")
rs("photo_x")=Request("photo_x")
rs("photo_y")=Request("photo_y")
rs("message")=Request("message")
rs.update
rs.close
set rs=nothing

conn.close
set conn=nothing%>