<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
%>
<HTML>
<HEAD>
	<Title>班级课程排课</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>排课信息管理--&gt;选择排课的班级
			 </td>
	   </tr><br>
		<tr>
		    <td colspan=2>&nbsp;&nbsp;选择学期:&nbsp;
			    <select name=termId>
				    <option value="">请选择</option>
					  <%
					    dim sqlString
						  sqlString = "select * from termInfo"
						  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
						  termInfoRs.Open sqlString,conn,1,1
						  while not termInfoRs.EOF
						    Response.Write "<option value='" & termInfoRs("termId") & "'>" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "年" & termInfoRs("termUpOrDown") & "</option>"
							  termInfoRs.MoveNext
						  wend
					  %>
				  </select>
			  </td>
		  </tr>
		  <tr>
		    <td colspan=2>&nbsp;&nbsp;选择班级:&nbsp;
			    <select name=classNumber>
				    <option value="">请选择</option>
					  <%
						  sqlString = "select classNumber,className from [classInfo]"
						  set classInfoRs = Server.CreateObject("ADODB.RecordSet")
						  classInfoRs.Open sqlString,conn,1,1
						  while not classInfoRs.EOF
						    Response.Write "<option value='" & classInfoRs("classNumber") & "'>" & classInfoRs("className") & "</option>"
							  classInfoRs.MoveNext
						  wend
					  %>
				  </select>
			  </td>
		  </tr>
		  <tr>
		    <td colspan=2>&nbsp;&nbsp;<input type="submit" name="submit" value="查询班级课程"></td>
		  </tr>
	  </table>
	  <br>
	  <table width=600 border=0 cellpadding=0 cellspacing=0 align="center">
		  <%
		    '如果要求查询某个班级某个学期的课程
		    if Request("submit") <> "" then
			    '判断是否选择了学期
				  if Request("termId") = "" then
					  Response.Write "<script>alert('请选择学期!');</script>"
					elseif Request("classNumber") = "" then
					  Response.Write "<script>alert('请选择班级!');</script>"
					else
					  '查询该学期该班级的所有课程
					  sqlString = "select * from [classCourseInfo] where termId=" & Request("termId") & " and classNumber='" & Request("classNumber") & "'"
					  set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
					  classCourseInfoRs.Open sqlString,conn,1,1
					  if not classCourseInfoRs.EOF then
					    Response.Write "<tr><td colspan=4 style='color:red;' align=center>" & GetClassNameByNumber(Request("classNumber")) & " " & GetTermnameById(Request("termId")) & " 课程信息</td></tr>"
					    Response.Write "<tr><td>课程编号</td><td>课程名称</td><td>课程学分</td><td>操作</td></tr>"
					  else
					    Response.Write "<tr></td><td colspan=4 style='color:red;' align=center>还没有班级课程信息</td></tr>"
					  end if
					  '输出每门课程的信息
					  while not classCourseInfoRs.EOF
					    Response.Write "<tr><td>" & classCourseInfoRs("courseNumber") & "</td><td>" & classCourseInfoRs("courseName") & "</td><td align=center>" & classCourseInfoRs("courseScore") & "</td><td><a href='classCourseTeachMakeSecond.asp?courseNumber=" & classCourseInfoRs("courseNumber") & "'>排课管理</a></td></tr>"
					    classCourseInfoRs.MoveNext
					  wend
					end if
			  end if
		  %>
	  </table>
  </form>
  
</BODY>

</HTML>
