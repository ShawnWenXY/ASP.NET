<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  dim sqlString,teachId,teachInfoRs,courseName,termName,teacherName
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'取得上课信息编号
	teachId = CInt(Request("teachId"))
	sqlString = "select * from [publicCourseTeach] where teachId=" & teachId
	set teachInfoRs = Server.CreateObject("ADODB.RecordSet")
	teachInfoRs.Open sqlString,conn,1,1
	if not teachInfoRs.EOF then
	  '取得上课的课程的名称
	  courseName = GetClassCourseNameByNumber(teachInfoRs("courseNumber"))
	  '取得该上课信息所在的学期信息
	  termName = GetTermnameById(teachInfoRs("termId"))
	  '取得该上课信息的授课教师的姓名
	  teacherName = GetTeacherNameByNumber(teachInfoRs("teacherNumber"))
	end if
%>
<HTML>
<HEAD>
	<Title>课程排课详细信息查看</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/list.gif" width=14px height=14px>排课信息管理--&gt;选修课程详细排课信息查看
			 </td>
	   </tr>
		 <tr>
		   <td width=100 align="right">课程名称:</td>
			 <td><%=courseName%></td>
		 </tr>
		 <tr>
		   <td width=100px align="right">上课学期:</td>
			 <td>
			   <%=termName%>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">上课地点:</td>
			 <td>
			   <%=teachInfoRs("teachClassRoom")%>
			 </td>
		 </tr>
		  <tr>
		   <td width=100px align="right">上课时间:</td>
			 <td>星期
			  <%=teachInfoRs("teachDay")%>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">详细上课信息</td>
			 <td>
			   &nbsp;<input type="checkbox" name=MorningOne value="1" <%if teachInfoRs("MorningOne") = True then Response.Write "checked" end if%>>上午第一节
				 &nbsp;<input type="checkbox" name=MorningTwo value="1" <% if CInt(teachInfoRs("MorningTwo")) = True then Response.Write "checked" end if%>>上午第二节<br>
				 &nbsp;<input type="checkbox" name=MorningThree value="1" <% if CInt(teachInfoRs("MorningThree")) = True then Response.Write "checked" end if%>>上午第三节
				 &nbsp;<input type="checkbox" name=MorningFour value="1" <% if CInt(teachInfoRs("MorningFour")) = True then Response.Write "checked" end if%>>上午第四节<br>
				 &nbsp;<input type="checkbox" name=MorningFive value="1" <% if CInt(teachInfoRs("MorningFive")) = True then Response.Write "checked" end if%>>上午第五节
				 &nbsp;<input type="checkbox" name=AfternoonOne value="1" <% if CInt(teachInfoRs("AfternoonOne")) = True then Response.Write "checked" end if%>>下午第一节<br>
				 &nbsp;<input type="checkbox" name=AfternoonTwo value="1" <% if CInt(teachInfoRs("AfternoonTwo")) = True then Response.Write "checked" end if%>>下午第二节
				 &nbsp;<input type="checkbox" name=AfternoonThree value="1" <% if CInt(teachInfoRs("AfternoonThree")) = True then Response.Write "checked" end if%>>下午第三节<br>
				 &nbsp;<input type="checkbox" name=AfternoonFour value="1" <% if CInt(teachInfoRs("AfternoonFour")) = True then Response.Write "checked" end if%>>下午第四节
				 &nbsp;<input type="checkbox" name=EveningOne value="1" <% if CInt(teachInfoRs("EveningOne")) = True then Response.Write "checked" end if%>>晚上第一节<br>
				 &nbsp;<input type="checkbox" name=EveningTwo value="1" <% if CInt(teachInfoRs("EveningTwo")) = True then Response.Write "checked" end if%>>晚上第二节
				 &nbsp;<input type="checkbox" name=EveningThree value="1" <% if CInt(teachInfoRs("EveningThree")) = True then Response.Write "checked" end if%>>晚上第三节
			 </td>
		 </tr>
		 <tr bgcolor="#ffffff">
        <td height="30" colspan="2" align="center">
				 <input type="button" value="返回" onClick="javascript:location.href='publicCourseTeachMakeSecond.asp?courseNumber=<%=teachInfoRs("courseNumber")%>'">
		    </td>
      </tr>
	 </table>
 </form>
</body>
</html>