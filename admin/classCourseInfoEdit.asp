<!--#include virtual="/DataBase/conn.asp"-->
<%
  'errMessage保存错误信息
  dim errMessage
  errMessage = ""
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'如果管理员修改了课程信息并提交
	if Request("submit") <> "" then
	  '如果没有输入课程名称
	  if Request("courseName") = "" then
	    errMessage = "请输入课程名称!"
	  end if
	  
	  '根据错误消息errMessage内容决定是否执行班级课程信息的修改操作
	  if errMessage = "" then
	    sqlString = "select * from [classCourseInfo] where courseNumber='" & Request("courseNumber") & "'"
		  set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  classCourseInfoRs.Open sqlString,conn,1,3
		  classCourseInfoRs("courseName") = Request("courseName")
		  classCourseInfoRs("classNumber") = Request("classNumber")
		  classCourseInfoRs("termId") = CInt(Request("termId"))
		  classCourseInfoRs("courseScore") = CSng(Request("courseScore"))
		  classCourseInfoRs("courseMemo") = Request("courseMemo")
		  classCourseInfoRs.Update
		  Response.Write "<script>alert('班级课程信息修改成功!');</script>"
	  else
	    Response.Write "<script>alert('" & errMessage & "');</script>"
	  end if
	end if
	
	sqlString = "select * from [classCourseInfo] where courseNumber='" & Request("courseNumber") & "'"
	set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
	classCourseInfoRs.Open sqlString,conn,1,1
%>

<HTML>
<HEAD>
	<Title>班级课程信息修改</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>课程信息管理--&gt;班级课程信息修改
			 </td>
	   </tr>
		 <tr>
		    <td width=100 align="right">课程编号:</td>
		    <td><%=Request("courseNumber")%><input type=hidden name=courseNumber size=20  value='<%=Request("courseNumber")%>'></td>
		  </tr>
		<tr>
		    <td width=100 align="right">所在学期:</td>
		    <td>
			    <select name=termId>
					  <%
						  sqlString = "select * from termInfo"
						  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
						  termInfoRs.Open sqlString,conn,1,1
						  while not termInfoRs.EOF
						    selected = ""
							  if termInfoRs("termId") = classCourseInfoRs("termId") then
							    selected = "selected"
								end if
						    Response.Write "<option value='" & termInfoRs("termId") & "' " & selected & ">" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "年" & termInfoRs("termUpOrDown") & "</option>"
							  termInfoRs.MoveNext
						  wend
					  %>
				  </select>
			  </td>
		  </tr>
		<tr>
		   <td width=100px align="right">所在班级:</td>
			 <td>
			   <select name=classNumber>				
				<%
					  sqlString = "select classNumber,className from [classInfo]"
					  set classInfoRs = Server.CreateObject("ADODB.RecordSet")
					  classInfoRs.Open sqlString,conn,1,1
					  while not classInfoRs.EOF
					    selected = ""
						  if classInfoRs("classNumber") = classCourseInfoRs("classNumber") then
						    selected = "selected"
							end if
					    Response.Write "<option value='" & classInfoRs("classNumber") & "' " & selected & ">" & classInfoRs("className") & "</option>"
						  classInfoRs.MoveNext
					  wend
				  %>
				 </select>
			 </td>
		 </tr>
		  <tr>
		    <td width=100 align="right">课程名称:</td>
		    <td><input type=text name=courseName size=20 value='<%=classCourseInfoRs("courseName")%>'></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">课程学分:</td>
		    <td><input type=text name=courseScore size=5 value=<%=classCourseInfoRs("courseScore")%>>分</td>
		  </tr>
		  <tr>
		    <td width=100 align="right">附加信息:</td>
		    <td><textarea cols=40 rows=5 name=courseMemo><%=classCourseInfoRs("courseMemo")%></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" 确认更新">
		      <input type=button value="返回" onClick="javascript:location.href='classCourseInfoManage.asp'">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>
</HTML>
