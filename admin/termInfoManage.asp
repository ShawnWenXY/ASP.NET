<!--#include virtual="/DataBase/conn.asp"-->
<%
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'添加新的学期信息
	if Request("action") = "add" then
	  termBeginYear = CInt(Request("termBeginYear"))
	  termEndYear = CInt(Request("termEndYear"))
	  termUpOrDown = Request("termUpOrDown")
	  if (termEndYear - termBeginYear) <> 1 then
	    Response.Write "<script>alert('你输入的学期年份信息不正确!');</script>"
	  else
	    sqlString = "select * from [termInfo]"
		  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
		  termInfoRs.Open sqlString,conn,1,3
		  termInfoRs.AddNew
		  termInfoRs("termBeginYear") = termBeginYear
		  termInfoRs("termEndYear") = termEndYear
		  termInfoRs("termUpOrDown") = termUpOrDown
		  termInfoRs.Update
		  termInfoRs.Close
		  set termInfoRs = Nothing
		  Response.Write "<script>alert('学期信息添加成功!');</script>"
	  end if
	'删除某个学期信息
	elseif Request("action") = "del" then
	  '取得要删除的学期编号
	  termId = CInt(Request("termId"))
	  sqlString = "delete from [termInfo] where termId=" & termId
	  conn.Execute(sqlString)
	  Response.Write "<script>alert('学期信息删除成功!');</script>"
	end if
%>
<HTML>
<HEAD>
	<Title>学期信息管理</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce">
	 <table width=400 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/list.gif" width=14px height=14px>学期信息管理
			 </td>
	   </tr>
		<tr><td>学期信息</td><td>删除</td></tr>
		<%
		  sqlString = "select * from [termInfo]"
		  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
		  termInfoRs.Open sqlString,conn,1,1
		  while not termInfoRs.EOF
		    Response.Write "<tr><td>" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "年" & termInfoRs("termUpOrDown") & "</td><td><a href='termInfoManage.asp?action=del&termId=" & termInfoRs("termId") & "' onclick=" & """" & "javascript:return confirm('决定删除此记录吗?');" & """" & "><img src='../images/delete.gif' width=12 height=12 border=0>删除</a></td></tr>"
			  termInfoRs.MoveNext
		  wend
		%>
		<tr><td colspan=2>添加新学期信息:<input type=text name=termBeginYear size=5>年-<input type=text name=termEndYear size=5>年
		<select name=termUpOrDown>
		  <option value="上学期">上学期</option>
		  <option value="下学期">下学期</option>
		</select><input type=hidden name=action value="add"><input type="submit" value='添加'></td></tr>
	</table>
</form>
</body>
</html>