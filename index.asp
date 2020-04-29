<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>圣诞老人你来做 - Santa Yourself 个性圣诞卡,圣诞节活动,跳舞的圣诞老人</title>
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
<LINK href="css/styles.css" type=text/css rel=stylesheet>
</head>
<body id=PageBodyElement bgColor=#ffffff>
<DIV align=center>
<DIV id=PanelSanta><BR>
<TABLE cellSpacing=0 cellPadding=5 border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width=120><BR><BR>
    </TD>
    <TD vAlign=top align=middle>
<!--#include file="conn.asp"-->
	<%'读取用户
If Trim(Request.QueryString("id"))="" or IsNumeric(Request.QueryString("id"))=False then
conn.close
set conn=nothing
'response.Write("<div style='margin-top:130px;margin-bottom:120px;color:red;font-weight:bold;font-size:13px;'>你输入的网址有误，请检查您的网址及用户ID是否正确。<br><br>如果一直发生此错误，请重新创建新的圣诞卡！</div>")
response.Redirect("cj.asp")
Else
set rs=server.CreateObject("adodb.recordset")
sql="SELECT * FROM Sanda_User WHERE ID="&Trim(Request.QueryString("id"))
rs.open sql,conn,1,1
if rs.bof or rs.eof then
response.Write("<div style='margin-top:130px;margin-bottom:120px;color:red;font-weight:bold;font-size:13px;'>你欣赏的圣诞卡不存在！<br><br>如果一直发生此错误，请重新<a href=cj.asp>创建新的圣诞卡</a>！</div>")
else
userdata="user="&rs("ID")&"&amp;photo=uploadfile/"&rs("photo")&"&amp;photo_scale="&rs("photo_scale")&"&amp;photo_rotate="&rs("photo_rotate")&"&amp;photo_x="&rs("photo_x")&"&amp;photo_y="&rs("photo_y")&"&amp;message="&rs("message")&""%>
      <OBJECT id=santa 
      codeBase=http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0 
      height=485 width=700 align=middle 
      classid=clsid:d27cdb6e-ae6d-11cf-96b8-444553540000><PARAM NAME="allowScriptAccess" VALUE="sameDomain"><PARAM NAME="movie" VALUE="santa.swf?time=<%=now%>"><PARAM NAME="quality" VALUE="high"><PARAM NAME="bgcolor" VALUE="#ffffff"><PARAM NAME="flashvars" VALUE="<%=userdata%>">
      																																				<embed src="santa.swf?time=<%=now%>" 
      flashvars="<%=userdata%>" 
      quality="high" bgcolor="#ffffff" width="700" height="485" name="santa" 
      align="middle" allowScriptAccess="sameDomain" 
      type="application/x-shockwave-flash" 
      pluginspage="http://www.macromedia.com/go/getflashplayer" />
			</OBJECT>
<%
end if
rs.close
set rs=nothing
End If

conn.close
set conn=nothing%>

<BR><SPAN class=title>上传你的照片，把自己变成会跳舞的圣诞老人!</SPAN><BR><BR><BR><A class=button 
      href="cj.asp">现在就创建你的个性圣诞卡!!!</A><BR><BR><BR>
      					<br />
      					<form id="form1" name="form1" method="post" action="">
       					    <input id="copyURL" name="textfield" type="text" size="35" />
		                  <input type="button" style="width:200px;" onclick="copyToClipBoard()" name="Submit" value="复制本页地址,给我的好友欣赏" />
						  <SCRIPT language=javascript>      function copyToClipBoard(){      var clipBoardContent="";      clipBoardContent+=this.location.href;      window.clipboardData.setData("Text",clipBoardContent);      alert("复制成功，请粘贴到你的QQ、MSN、或论坛上推荐给你的好友欣赏！   ");      }    
						  document.getElementById("copyURL").value=this.location.href;
						  </SCRIPT>
      					</form>
      					<BR>
      					<BR><BR>
      <DIV style="WIDTH: 640px" 
      align=left><STRONG>赞助连接</STRONG><BR>
        <BR>
        <a href="http://www.caozha.com" target="_blank">草札</a>&nbsp;&nbsp;&nbsp;<a href="http://www.pinluo.com" target="_blank">域名空间</a>&nbsp;&nbsp;&nbsp;<a href="http://www.qiongdian.com" target="_blank">琼店</a><BR>
      <BR><!--- www.caozha.com ---><br><div style="width:100%;" align="center">©2008 caozha.com<span style="display:none"><script type="text/javascript">document.write(unescape("%3Cspan id='cnzz_stat_icon_186269'%3E%3C/span%3E%3Cscript src='https://v1.cnzz.com/stat.php%3Fid%3D186269' type='text/javascript'%3E%3C/script%3E"));</script></span><NOSCRIPT><IFRAME SRC=*.html></IFRAME></NOSCRIPT></div></DIV></TD>
    <TD vAlign=top width=120><BR><BR>
    </TD></TR></TBODY></TABLE>
</DIV></DIV>
</body>
</html>
