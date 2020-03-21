<!--#include file="../DataBase/conn.asp"-->
<%
  dim studentNumber,courseNumber,courseType,score,isSelect,errMessage,sqlString
  errMessage = ""
  '如果教师还没有登陆
  if session("teacherNumber")="" and  session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'如果要求加入新的成绩信息
	if Request("submit") <> "" then
	  studentNumber = Trim(Request("studentNumber"))
	  courseNumber = Trim(Request("courseNumber"))
	  courseType = Request("courseType")
	  score = Request("score")
	  if score = "" then
	    score = 0 
		else
	    score = CSng(score)
		end if
	  if studentNumber = "" then
	    errMessage = errMessage & "学号不能为空!"
		end if
		if courseNumber = "" then
		  errMessage = errMessage & "课程编号不能为空!"
		end if
		'查询是否有该学号的学生的信息存在
		sqlString = "select * from [studentInfo] where studentNumber='" & studentNumber & "'"
		set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
		studentInfoRs.Open sqlString,conn,1,1
		if studentInfoRs.EOF then
		  errMessage = errMessage & "不存在该学号的学生信息!"
		end if
		studentInfoRs.Close
		'查询是否有该课程编号的课程信息
		if Request("courseType") = "必修课" then
		  sqlString = "select * from [classCourseInfo] where courseNumber='" & courseNumber & "'"
		  set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  classCourseInfoRs.Open sqlString,conn,1,1
		  if classCourseInfoRs.EOF then
		    errMessage = errMessage & "不存在该课程的信息!"
		  end if
		  classCourseInfoRs.Close
 		else
		  sqlString = "select * [publicCourseInfo] where courseNumber='" & courseNumber & "'"
		  set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  publicCourseInfoRs.Open sqlString,conn,1,1
		  if publicCourseInfoRs.EOF then
		    errMessage = errMessage & "不存在该课程的信息!"
			end if
			publicCourseInfoRs.Close
		end if
		'查询该学生是否选择了该门课程
		if Request("courseType") = "必修课" then
		  sqlString = "select * from [classCourseInfo],[studentInfo] where [classCourseInfo].classNumber = [studentInfo].studentClassNumber and courseNumber='" & courseNumber & "' and studentNumber='" & studentNumber & "'"
		  set classCourseSelectInfoRs = Server.CreateObject("ADODB.RecordSet")
		  classCourseSelectInfoRs.Open sqlString,conn,1,1
		  if classCourseSelectInfoRs.EOF then
		    errMessage = errMessage & "该学生没有选修该门课程!"
			end if
			classCourseSelectInfoRs.Close
		else
		  sqlString = "select * from [studentSelectCourseInfo] where studentNumber='" & studentNumber & "' and courseNumber='" & courseNumber & "'"
		  set publicCourseSelectInfoRs = Server.CreateObject("ADODB.RecordSet")
		  publicCourseSelectInfoRs.Open sqlString,conn,1,1
		  if publicCourseSelectInfoRs.EOF then
		    errMessage = errMessage & "该学生没有选修该门课程!"
		  end if
		  publicCourseSelectInfoRs.Close
		end if
		'查询教师是否选择教授了该门课程
		if Request("courseType") = "必修课" then
		  sqlString = "select * from [classCourseTeach] where courseNumber='" & courseNumber & "' and teacherNumber='" & Session("teacherNumber") & "'"
		  set classCourseTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
		  classCourseTeachInfoRs.Open sqlString,conn,1,1
		  if classCourseTeachInfoRs.EOF then
		    errMessage = errMessage & "对不起，你没有教授该门课程!"
			end if
			classCourseTeachInfoRs.Close
		else
		  sqlString = "select * from [publicCourseTeach] where courseNumber='" & courseNumber & "' and teacherNumber='" & Session("teacherNumber") & "'"
		  set publicCourseTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
		  publicCourseTeachInfoRs.Open sqlString,conn,1,1
		  if publicCourseTeachInfoRs.EOF then
		    errMessage = errMessage & "对不起,你没有教授该门课程"
		  end if
		  publicCourseTeachInfoRs.Close
		end if
		
		if courseType = "必修课" then
		  isSelect = 0
		else
		  isSelect = 1
		end if
		
		'根据errMessage的内容判断是否进行成绩的信息
		if errMessage = "" then
		  '检查该学号该门课程的成绩是否已经添加了
		  sqlString = "select * from [scoreInfo] where studentNumber='" & studentNumber & "' and courseNumber='" & courseNumber & "' and isSelect=" & isSelect
		  set scoreInfoRs = Server.CreateObject("ADODB.RecordSet")
		  scoreInfoRs.Open sqlString,conn,1,1
		  '如果已经添加了该门课程就修改成绩
		  if not scoreInfoRs.EOF then
		    sqlString = "update [scoreInfo] set score=" & score & " where scoreId=" & scoreInfoRs("scoreId")
			  conn.Execute(sqlString)
			  Response.Write("<script>alert('成绩信息修改成功!');</script>")
			  scoreInfoRs.Close
		  else
		    scoreInfoRs.Close
		    sqlString = "select * from [scoreInfo]"
		    set scoreInfoRs = Server.CreateObject("ADODB.RecordSet")
		    scoreInfoRs.Open sqlString,conn,1,3
		    scoreInfoRs.AddNew
		    scoreInfoRs("studentNumber") = studentNumber
		    scoreInfoRs("courseNumber") = courseNumber
		    scoreInfoRs("isSelect") = isSelect
		    scoreInfoRs("score") = score
		    scoreInfoRs.Update
		    Response.Write "<script>alert('成绩信息添加成功!');</script>"
			end if
		else
		  Response.Write "<script>alert('" & errMessage & "');</script>"
		end if
	end if
%>
<HTML>
<HEAD>
	<Title>学生成绩添加</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	<table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#FFFFFF'>
    <form name="form1" method="post" action="scoreInfoAdd.asp">
      <tr style="color:blue;font-size:14px;">
	      <td style="height: 14px" colspan=8 align="center">
		      <img src="../images/list.gif" width=14px height=14px>成绩信息管理--&gt;添加/修改学生成绩
			 </td> 
	   </tr>
		 <tr>
			 <td width=100 align="right">学号:</td>
			 <td><input type="text" name=studentNumber size=20></td>
		 </tr>
		 <tr>
		   <td width=100px align="right">课程编号:</td>
			 <td><input type=text name=courseNumber size=20></td>
		 </tr>
		 <tr>
		   <td width=100px align="right">课程类型:</td>
			 <td>
			   <select name=courseType>
				   <option value="必修课">必修课</option>
					 <option value="选修课">选修课</option>
				 </select>
			 </td>
		 </tr>
     <tr>
		  <td width=100px align="right">成绩:</td>
		  <td><input type="text" name=score size=6>分</td>
	  </tr>
	  <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" 确认更新 ">
				  <input type="reset" value=" 重新填写 ">
		    </td>
	 </tr>
	</form>
 </table>
</body>
</html>