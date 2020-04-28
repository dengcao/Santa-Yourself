<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<HTML><HEAD><TITLE>圣诞贺卡创建成功！ - 创建圣诞老人,把自己变成会跳舞的圣诞老人！</TITLE>
<%dim userid,photo,photo_scale,photo_rotate,photo_x,photo_y,message
userid=trim(Request.Cookies("userid"))
if userid="" then
response.Redirect("cj.asp")
end if
%>
<LINK 
href="css/styles.css" 
type=text/css rel=stylesheet>
<META content="MSHTML 6.00.6000.16788" name=GENERATOR></HEAD>
<BODY bgColor=#e0e0e0>
<DIV align=center>
<DIV class=mypage><BR>
<H1 
style="PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 0px; PADDING-TOP: 0px"><IMG 
height=73
src="images/title.gif" 
width=700 border=0></H1><BR>
<DIV 
style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px"><!-- MAIN CONTENT -->
<DIV style="FONT-SIZE: 16px; COLOR: #acacac" align=center><STRONG 
style="COLOR: #000000">第一步</STRONG> ......... <STRONG 
style="COLOR: #000000">第二步</STRONG> ......... <STRONG 
style="FONT-SIZE: 28px; COLOR: #cc0000">第三步</STRONG> </DIV><BR><BR>
<DIV align=center><STRONG class=title style="color:#FF0000">恭喜，创建成功！预览您的圣诞贺卡！</STRONG> 
<br>
<form id="form1" name="form1" method="post" action="">
       					    <input id="copyURL" name="textfield" type="text" size="35" />
		                  <input type="button" style="width:200px;" onclick="copyToClipBoard()" name="Submit" value="复制本页地址,给我的好友欣赏" />
						  <SCRIPT language=javascript>
						  var pageurl="http://"+this.location.host+"/index.asp?id=<%=userid%>";
						  function copyToClipBoard(){      var clipBoardContent="";      clipBoardContent+=pageurl;      window.clipboardData.setData("Text",clipBoardContent);      alert("复制成功，请粘贴到你的QQ、MSN、或论坛上推荐给你的好友欣赏！   ");      }    
						  document.getElementById("copyURL").value=pageurl;
						  </SCRIPT>
      					</form><br>
</DIV>
<%
if userid<>"" then%>
<!--#include file="conn.asp"-->
<%set rs=server.CreateObject("adodb.recordset")
sql="SELECT * FROM Sanda_User where id="&int(userid)
rs.open sql,conn,1,1
photo=rs("photo")
photo_scale=rs("photo_scale")
photo_rotate=rs("photo_rotate")
photo_x=rs("photo_x")
photo_y=rs("photo_y")
message=rs("message")
userid=rs("id")
rs.close
set rs=nothing
conn.close
set conn=nothing
else
response.Write("<div style='margin-top:130px;margin-bottom:120px;color:red;font-weight:bold;font-size:13px;'>提示：系统获取参数错误！<br><br>如果一直发生此错误，请重新创建新的圣诞卡！</div>")
end if
%>
<DIV align=center>
<OBJECT id=santa 
codeBase=http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0 
height=485 width=700 align=middle 
classid=clsid:d27cdb6e-ae6d-11cf-96b8-444553540000><PARAM NAME="allowScriptAccess" VALUE="sameDomain"><PARAM NAME="movie" VALUE="santa.swf?time=<%=now%>"><PARAM NAME="quality" VALUE="high"><PARAM NAME="bgcolor" VALUE="#ffffff"><PARAM NAME="flashvars" VALUE="user=<%=userid%>&amp;photo=<%="uploadfile/"&photo%>&amp;photoScale=<%=photo_scale%>&amp;photoRotate=<%=photo_rotate%>&amp;photoX=<%=photo_x%>&amp;photoY=<%=photo_y%>&amp;msg=<%=message%>">
																														<embed src="santa.swf?time=<%=now%>" 
flashvars="user=<%=userid%>&photo=<%="uploadfile/"&photo%>&photoScale=<%=photo_scale%>&photoRotate=<%=photo_rotate%>&photoX=<%=photo_x%>&photoY=<%=photo_y%>&msg=<%=message%>" 
quality="high" bgcolor="#ffffff" width="700" height="485" name="santa" 
align="middle" allowScriptAccess="sameDomain" 
type="application/x-shockwave-flash" 
pluginspage="http://www.macromedia.com/go/getflashplayer" />
				</OBJECT></DIV><BR><BR><A class=button 
href="cj.asp">« 返回重新创建</A>&nbsp;&nbsp;<A 
class=button id=LinkMonetize 
href="./?id=<%=userid%>">打开我的贺卡</A> 
<BR><BR><BR><BR><!-- END MAIN CONTENT --></DIV></DIV><BR>
<DIV style="FONT-SIZE: 12px"><BR>
<a href="http://www.caozha.com" target="_blank">草札</a>&nbsp;&nbsp;&nbsp;<a href="http://www.pinluo.com" target="_blank">域名空间</a>&nbsp;&nbsp;&nbsp;<a href="http://www.qiongdian.com" target="_blank">琼店</a><BR>
      <BR><!--- www.caozha.com ---><br><div style="width:100%;" align="center">©2008 caozha.com<span style="display:none"><script type="text/javascript">document.write(unescape("%3Cspan id='cnzz_stat_icon_186269'%3E%3C/span%3E%3Cscript src='https://v1.cnzz.com/stat.php%3Fid%3D186269' type='text/javascript'%3E%3C/script%3E"));</script></span><NOSCRIPT><IFRAME SRC=*.html></IFRAME></NOSCRIPT></div>
</DIV><BR>
</DIV></BODY></HTML>
