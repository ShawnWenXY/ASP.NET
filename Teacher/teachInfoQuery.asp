<!--#include file="../DataBase/conn.asp"-->
<!--#include file="../System/function.asp"-->
<%
  dim sqlString,courseNumber,termId
  '如果教师还没有登陆
  if session("teacherNumber")="" and  session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'取得查询的课程编号的关键字
	courseNumber = Request("courseNumber")
	'取得查询的学期信息
	termId = Request("termId")
	'从班级必修课上课信息表中进行查询的sql
	sqlString = "select * from [classCourseTeach] where teacherNumber='" & session("teacherNumber") & "'"
	if courseNumber <> "" then
	  sqlString = sqlString & " and courseNumber like '%" & courseNumber & "%'"
	end if
	if termId = "" then
	  termId = 0
	end if
	if termId <> 0 then
	  sqlString = sqlString & " and termId=" & termId
	end if
	set classCourseTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	classCourseTeachInfoRs.Open sqlString,conn,1,1
	'从选修课上课细心你表中进行查询的sql语句
	sqlString = "select * from [publicCourseTeach] where teacherNumber='" & Session("teacherNumber") & "'"
	if courseNumber <> "" then
	  sqlString = sqlString & " and courseNumber like '%" & courseNumber & "%'"
	end if
	if termId <> 0 then
	  sqlString = sqlString & " and termId=" & termId
	end if
	set publicCourseTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	publicCourseTeachInfoRs.Open sqlString,conn,1,1
%>
<html>
<head>
   <title>教师授课信息查询</title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	<table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#FFFFFF'>
    <form name="form1" method="post" action="teachInfoQuery.asp">
      <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=8 align="center">
		      <img src="../images/list.gif" width=14px height=14px>选课信息管理--&gt;授课信息查询
			 </td>
	   </tr>
     <tr>
	     <td  align="left" height="22" colspan="7" bgcolor="#ffffff"> 
	       课程编号:<input type="text" name=courseNumber size=18 value='<%=courseNumber%>'>&nbsp;
	       所在学期:
	       <select name=termId>
	        <option value="">选择学期</option>
		    <%
		    sqlString = "select * from [termInfo]"
			  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
			  termInfoRs.Open sqlString,conn,1,1
			  while not termInfoRs.EOF
			    selected = ""
				  if termId = "" then
				    termId = 0
					else
					  termId = CInt(termId)
					end if
				  if termInfoRs("termId") = termId then
				    selected = "selected"
					end if
			    Response.Write "<option value='" & termInfoRs("termId") & "' " & selected & ">" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "年" & termInfoRs("termUpOrDown") & "</option>"
					termInfoRs.MoveNext
			  wend
		  %>
	      </select>
	      <input type="submit" value=" 检索 " class="button1">
     </td>
    </tr>
	  <%
	    if classCourseTeachInfoRs.EOF and publicCourseTeachInfoRs.EOF then
		    Response.Write "<tr><td colspan=6 align=center>对不起,还没有对应的授课信息</td></tr>"
		  else
		    Response.Write "<tr><td>课程编号</td><td>所在学期</td><td>所在班级(或专业)</td><td>课程类型</td><td>上课教室</td><td>上课时间</td><td>详细</td></tr>"
		  end if
		  while not classCourseTeachInfoRs.EOF
		    Response.Write "<tr><td>" & classCourseTeachInfoRs("courseNumber") & "</td><td>" & GetTermnameById(classCourseTeachInfoRs("termId")) & "</td>"
			  Response.Write "<td>" & GetClassNameByNumber(classCourseTeachInfoRs("classNumber")) & "</td><td>必修课</td><td>" & classCourseTeachInfoRs("teachClassRoom") & "</td>"
			  Response.Write "<td>星期" & classCourseTeachInfoRs("teachDay") & "</td><td><a href='classCourseTeachDetail.asp?termId=" & termId & "&courseNumber=" & courseNumber & "&teachId=" & classCourseTeachInfoRs("teachId") & "'>详细</a></td></tr>"
		    classCourseTeachInfoRs.MoveNext
		  wend
		  while not publicCourseTeachInfoRs.EOF
		    Response.Write "<tr><td>" & publicCourseTeachInfoRs("courseNumber") & "</td><td>" & GetTermnameById(publicCourseTeachInfoRs("termId")) & "</td>"
			  Response.Write "<td>" & GetSpecialFieldNameByNumber(publicCourseTeachInfoRs("specialFieldNumber")) & "</td><td>选修课</td><td>" & publicCourseTeachInfoRs("teachClassRoom") & "</td>"
			  Response.Write "<td>星期" & publicCourseTeachInfoRs("teachDay") & "</td><td><a href='publicCourseTeachDetail.asp?termId=" & termId & "&courseNumber=" & courseNumber & "&teachId=" & publicCourseTeachInfoRs("teachId") & "'>详细</a></td></tr>"
		    publicCourseTeachInfoRs.MoveNext
		  wend
	  %>
    </form>
	</table>
</body>
</html>