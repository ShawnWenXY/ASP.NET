<!--#include virtual="inc_upload.asp"-->

<%
set upload=new upload_5xSoft
set file=upload.file("file1")
formPath="upload/"
if file.filesize>100 then
fileExt=lcase(right(file.filename,3))
if fileExt<>"gif" and fileExt<>"jpg" then
founderr=true
errmsg=errmsg+"<br>"+"<li>文件格式错误！"
end if
end if
randomize
ranNum=int(90000*rnd)+10000
filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
if file.FileSize>0 then 
file.SaveAs Server.mappath(FileName)
end if

Response.write "<body topmargin=0><table width=100% border=0 cellspacing=0 cellpadding=0><td CLASS=chinese>图片上传成功！<script>parent.document.frmAnnounce.photoAddress.value = '" & fileName & "';parent.document.frmAnnounce.submit.disabled = false;</script><a href=upload.asp>[重新上传]</a></td></table></body>"
%>