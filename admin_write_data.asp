<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<%
'☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
'☆                                                                         ☆
'☆  程 序：圣诞老人你来做 - Santa Yourself                                    ☆
'☆  日 期：2008-12                                                          ☆
'☆  开 发：草札(www.caozha.com)                                              ☆
'☆  鸣 谢：穷店(www.qiongdian.com) 品络(www.pinluo.com)                      ☆
'☆  声 明: 使用本程序源码必须保留此版权声明等相关信息！                            ☆
'☆  Copyright ©2008 www.caozha.com All Rights Reserved.                    ☆
'☆                                                                         ☆
'☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
%>
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