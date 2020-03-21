<!--#include virtual="/Database/conn.asp"-->
<%
  'errMessage保存错误信息
  dim errMessage,sqlString
  errMessage = ""
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'如果管理员添加了新的公选课程信息并提交
	if Request("submit") <> "" then
	  '如果没有选择所在的学期
	  if Request("termId") = "" then
	    errMessage = "请选择所在的学期!"
	  end if
	  '如果没有选择专业信息
	  if Request("specialFieldNumber") = "" then
	    errMessage = "请选择专业!"
	  end if
	  '如果没有输入课程编号
	  if Request("courseNumber") = "" then
	    errMessage = "请输入课程编号!"
	  end if
	  '如果没有输入课程名称
	  if Request("courseName") = "" then
	    errMessage = "请输入课程名称!"
	  '检查该学期该专业该课程信息是否已经存在
	  else
	    sqlString = "select * from [publicCourseInfo] where specialFieldNumber='" & Request("specialFieldNumber") & "' and termId=" & Request("termId") & " and courseName='" & Request("courseName") & "'"
		  set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  publicCourseInfoRs.Open sqlString,conn,1,1
		  if not publicCourseInfoRs.EOF then
		    errMessage = "该学期该班级已经存在该课程名称信息"
		  end if
		  publicCourseInfoRs.Close
		  set publicCourseInfoRs = nothing
	  end if
	  '根据错误消息errMessage内容决定是否执行新公选课程信息的添加操作
	  if errMessage = "" then
	    sqlString = "select * from [publicCourseInfo]"
		  set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  publicCourseInfoRs.Open sqlString,conn,1,3
		  publicCourseInfoRs.AddNew
		  publicCourseInfoRs("courseNumber") = Request("courseNumber")
		  publicCourseInfoRs("courseName") = Request("courseName")
		  publicCourseInfoRs("courseType") = "选修课"
		  publicCourseInfoRs("specialFieldNumber") = Request("specialFieldNumber")
		  publicCourseInfoRs("termId") = CInt(Request("termId"))
		  publicCourseInfoRs("courseScore") = CSng(Request("courseScore"))
		  publicCourseInfoRs("courseMemo") = Request("courseMemo")
		  publicCourseInfoRs.Update
		  Response.Write "<script>alert('选修课程信息添加成功!');</script>"
	  else
	    Response.Write "<script>alert('" & errMessage & "');</script>"
	  end if
	end if
%>
<HTML>
<HEAD>
	<Title>专业公选课程信息添加</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>课程信息管理--&gt;专业公选信息添加
			 </td>
	   </tr><br>
		<tr>
		    <td width=100 align="right">所在学期:</td>
		    <td>
			    <select name=termId>
				    <option value="">请选择</option>
					  <%
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
		   <td width=100px align="right">选择专业:</td>
			 <td>
			   <select name=specialFieldNumber>
				   <option value="">请选择</option>
				  <%
					  sqlString = "select * from [specialFieldInfo]"
					  set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
					  specialFieldInfoRs.Open sqlString,conn,1,1
					  while not specialFieldInfoRs.Eof
					    Response.Write "<option value='" & specialFieldInfoRs("specialFieldNumber") & "'>" & specialFieldInfoRs("specialFieldName") & "</option>"
						  specialFieldInfoRs.MoveNext
					  wend
					  specialFieldInfoRs.Close
				  %>
				 </select>
			 </td>
		 </tr>
		 <tr>
		    <td width=100 align="right">课程编号:</td>
		    <td><input type=text name=courseNumber size=20></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">课程名称:</td>
		    <td><input type=text name=courseName size=20></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">课程学分:</td>
		    <td><input type=text name=courseScore size=5>分</td>
		  </tr>
		  <tr>
		    <td width=100 align="right">附加信息:</td>
		    <td><textarea cols=40 rows=5 name=courseMemo></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" 确认添加 ">
		      <input type="reset" value=" 重新填写 ">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>

</HTML>
