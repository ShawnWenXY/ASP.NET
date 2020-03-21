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
	<Title>公选课程排课</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>排课信息管理--&gt;选择公选课排课的专业
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
		    <td colspan=2>&nbsp;&nbsp;选择专业:&nbsp;
			    <select name=specialFieldNumber>
				    <option value="">请选择</option>
					  <%
						  sqlString = "select * from [specialFieldInfo]"
						  set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
						  specialFieldInfoRs.Open sqlString,conn,1,1
						  while not specialFieldInfoRs.EOF
						    Response.Write "<option value='" & specialFieldInfoRs("specialFieldNumber") & "'>" & specialFieldInfoRs("specialFieldName") & "</option>"
							  specialFieldInfoRs.MoveNext
						  wend
					  %>
				  </select>
			  </td>
		  </tr>
		  <tr>
		    <td colspan=2>&nbsp;&nbsp;<input type="submit" name="submit" value="查询专业选修课程"></td>
		  </tr>
	  </table>
	  <br>
	  <table width=600 border=0 cellpadding=0 cellspacing=0 align="center">
		  <%
		    '如果要求查询某个专业某个学期的选修课程
		    if Request("submit") <> "" then
			    '判断是否选择了学期
				  if Request("termId") = "" then
					  Response.Write "<script>alert('请选择学期!');</script>"
					elseif Request("specialFieldNumber") = "" then
					  Response.Write "<script>alert('请选择专业!');</script>"
					else
					  '查询该学期该专业的所有选修课程
					  sqlString = "select * from [publicCourseInfo] where termId=" & Request("termId") & " and specialFieldNumber='" & Request("specialFieldNumber") & "'"
					  set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
					  publicCourseInfoRs.Open sqlString,conn,1,1
					  if not publicCourseInfoRs.EOF then
					    Response.Write "<tr><td colspan=4 style='color:red;' align=center>" & GetSpecialFieldNameByNumber(Request("specialFieldNumber")) & " " & GetTermnameById(Request("termId")) & " 选修课程信息</td></tr>"
					    Response.Write "<tr><td>课程编号</td><td>课程名称</td><td>课程学分</td><td>操作</td></tr>"
					  else
					    Response.Write "<tr></td><td colspan=4 style='color:red;' align=center>还没有专业选修课程信息</td></tr>"
					  end if
					  '输出每门课程的信息
					  while not publicCourseInfoRs.EOF
					    Response.Write "<tr><td>" & publicCourseInfoRs("courseNumber") & "</td><td>" & publicCourseInfoRs("courseName") & "</td><td align=center>" & publicCourseInfoRs("courseScore") & "</td><td><a href='publicCourseTeachMakeSecond.asp?courseNumber=" & publicCourseInfoRs("courseNumber") & "'>排课管理</a></td></tr>"
					    publicCourseInfoRs.MoveNext
					  wend
					end if
			  end if
		  %>
	  </table>
  </form>
  
</BODY>

</HTML>
