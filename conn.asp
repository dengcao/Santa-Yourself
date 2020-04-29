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

set conn=server.CreateObject("adodb.connection")
'连接SQL  conn.connectionstring="driver={sql server};server=127.0.0.1;uid=sa;pwd=123456;database=chinaz_user"
'conn.open
db_path="data/sadan-user#2008-12-24.mdb"
Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.MapPath(db_path)
Conn.Open Connstr%>