<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<HTML><HEAD><TITLE>第二步，调整您的照片 - 创建圣诞老人,把自己变成会跳舞的圣诞老人！</TITLE>
<%
'☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
'☆                                                                         ☆
'☆  程 序：圣诞老人你来做 - Santa Yourself                                    ☆
'☆  日 期：2008-12                                                          ☆
'☆  开 发：草札(www.caozha.com)                                              ☆
'☆  鸣 谢：琼店(www.qiongdian.com) 品络(www.pinluo.com)                      ☆
'☆  声 明: 使用本程序源码必须保留此版权声明等相关信息！                            ☆
'☆  Copyright ©2008 www.caozha.com All Rights Reserved.                    ☆
'☆                                                                         ☆
'☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
%>
<%Server.ScriptTimeOut=5000%>
<!--#include FILE="upload_5xsoft.inc"-->
<%
Response.Cookies("userid")="" '去掉可能存在的保存的cookie
dim upload,file,formName,formPath,iCount,FileName
set upload=new upload_5xsoft ''建立上传对象
 formPath="uploadfile/"
 ''在目录后加(/)

 set file=upload.file("FileUpload")  ''生成一个文件对象
 if not(file.FileExt="jpg" or file.FileExt="png") then' or not(file.FileExt="gif" or file.FileExt="bmp") 
 response.Write("<div style='text-align:center;padding-top:150px;color:red;'><b>上传错误！系统只允许上传.jpg、.png格式的图片！<br><br>请返回重新选择图片！——<a href='cj_1.asp'>点击此处返回..</a></div>")
 response.End()
 end if
 if (file.FileSize>512000 or file.FileSize<=0) then
  response.Write("<div style='text-align:center;padding-top:150px;color:red;'><b>上传错误！系统只允许上传小于500Kb的图片！<br><br>请返回重新选择图片！——<a href='cj_1.asp'>点击此处返回..</a></div>")
   response.End()
 end if
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
 'FileName=File.FileName
  FileName=Replace(Replace(Replace(now,":",""),"-","")," ","")&"."&file.FileExt
  file.SaveAs Server.mappath(formPath&FileName)   ''保存文件
  
%>
<!--#include file="conn.asp"-->
<%
'读取用户----------------------------------------
dim userid,photo,photo_scale,photo_rotate,photo_x,photo_y,message
photo=FileName
photo_scale=100
photo_rotate=0
photo_x=0'17
photo_y=0'14
message="Merry+Christmas+and+Happy+New+Year+2009."
set rs=server.CreateObject("adodb.recordset")
sql="SELECT * FROM Sanda_User order by id desc"
rs.open sql,conn,1,3
rs.addnew
rs("photo")=photo
rs("photo_scale")=photo_scale
rs("photo_rotate")=photo_rotate
rs("photo_x")=photo_x
rs("photo_y")=photo_y
rs("message")=message
rs.update
userid=rs("id")
Response.Cookies("userid")=userid
rs.close
set rs=nothing
conn.close
set conn=nothing

'end --------------------------------
  
 end if
 set file=nothing

set upload=nothing  ''删除此对象
if trim(Request.Cookies("userid"))="" then
response.Redirect("cj.asp")
end if
%>
<LINK 
href="css/styles.css" 
type=text/css rel=stylesheet>
<SCRIPT>
		function jsUrl(myUrl) {
			//alert(myUrl);
			window.top.location.href = myUrl;
		}
	</SCRIPT>

<META content="MSHTML 6.00.6000.16788" name=GENERATOR></HEAD>
<BODY bgColor=#e0e0e0>
<DIV align=center>
<DIV class=mypage><BR>
<H1 
style="PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 0px; PADDING-TOP: 0px"><IMG 
height=73 src="images/title.gif" 
width=700 border=0></H1><BR>
<DIV 
style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px"><!-- MAIN CONTENT -->
<DIV style="FONT-SIZE: 16px; COLOR: #acacac" align=center><STRONG 
style="COLOR: #000000">第一步</STRONG> ......... <STRONG 
style="FONT-SIZE: 28px; COLOR: #cc0000">第二步</STRONG> ......... <STRONG>第三步</STRONG> </DIV><BR><BR>
<DIV align=center><STRONG class=title>调整你的照片，使它符合圣诞老人的脸，并输入您的信息！然后单击下一步。<BR>
    <BR><IMG height=75 alt="" 
src="images/newsfeed_1.jpg" 
width=75 border=0>&nbsp;&nbsp;<IMG height=75 alt="" 
src="images/newsfeed_2.jpg" 
width=75 border=0>&nbsp;&nbsp;<IMG height=75 alt="" 
src="images/newsfeed_3.jpg" 
width=75 border=0><BR><BR>
<OBJECT id=admin 
codeBase=http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0 
height=590 width=625 align=middle 
classid=clsid:d27cdb6e-ae6d-11cf-96b8-444553540000><PARAM NAME="allowScriptAccess" VALUE="always"><PARAM NAME="movie" VALUE="admin.swf?time=<%=now%>"><PARAM NAME="quality" VALUE="high"><PARAM NAME="bgcolor" VALUE="#ffffff"><PARAM NAME="flashvars" VALUE="user=<%=userid%>&amp;photo=<%="uploadfile/"&photo%>&amp;photoScale=<%=photo_scale%>&amp;photoRotate=<%=photo_rotate%>&amp;photoX=<%=photo_x%>&amp;photoY=<%=photo_y%>&amp;msg=<%=message%>">
<embed src="admin.swf?time=<%=now%>"
flashvars="user=<%=userid%>&photo=<%="uploadfile/"&photo%>&photoScale=<%=photo_scale%>&photoRotate=<%=photo_rotate%>&photoX=<%=photo_x%>&photoY=<%=photo_y%>&msg=<%=message%>" 
quality="high" bgcolor="#ffffff" width="625" height="590" name="admin" 
align="middle" allowScriptAccess="always" type="application/x-shockwave-flash" 
pluginspage="http://www.macromedia.com/go/getflashplayer" />
	  </OBJECT></DIV><BR><BR><!-- END MAIN CONTENT --></DIV></DIV>
<DIV style="FONT-SIZE: 12px"><BR>
<a href="http://www.caozha.com" target="_blank">草札</a>&nbsp;&nbsp;&nbsp;<a href="http://www.pinluo.com" target="_blank">域名空间</a>&nbsp;&nbsp;&nbsp;<a href="http://www.qiongdian.com" target="_blank">琼店</a><BR>
      <BR><!--- www.caozha.com ---><br><div style="width:100%;" align="center">©2008 caozha.com<span style="display:none"><script type="text/javascript">document.write(unescape("%3Cspan id='cnzz_stat_icon_186269'%3E%3C/span%3E%3Cscript src='https://v1.cnzz.com/stat.php%3Fid%3D186269' type='text/javascript'%3E%3C/script%3E"));</script></span><NOSCRIPT><IFRAME SRC=*.html></IFRAME></NOSCRIPT></div>
</DIV><BR>
</DIV></BODY></HTML>
