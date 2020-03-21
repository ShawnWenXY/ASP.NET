<!--#include virtual="/Database/conn.asp"-->
<%
  dim teacherNumber,sqlString
  teacherNumber = Request("teacherNumber")
  '首先取得该教师所教的所有必修课程
  sqlString = "select courseNumber from [classCourseTeach] where teacherNumber='" & teacherNumber & "'"
  set teacherTeachCourseRs = Server.CreateObject("ADODB.RecordSet")
  teacherTeachCourseRs.Open sqlString,conn,1,1
  while not teacherTeachCourseRs.EOF
    '删除该必须课程的信息记录
    sqlString = "delete from [classCourseInfo] where courseNumber='" & teacherTeachCourseRs("courseNumber") & "'"
	  conn.Execute(sqlString)
	  '删除该必修课程的成绩信息
	  sqlString = "delete from [scoreInfo] where courseNumber='" & teacherTeachCourseRs("courseNumber") & "' and isSelect=0"
	  conn.Execute(sqlString)
	  teacherTeachCourseRs.MoveNext
  wend 
  '删除该必修课程的授课信息
  sqlString = "delete from [classCourseTeach] where teacherNumber='" & teacherNumber & "'"
  conn.Execute(sqlString)
  teacherTeachCourseRs.Close
  '然后得到该教师所教的选修课程
  sqlString = "select courseNumber from [publicCourseTeach] where teacherNumber='" & teacherNumber & "'"
  'set teacherTeachCourseRs = Server.CreateObject("ADODB.RecordSet")
  teacherTeachCourseRs.Open sqlString,conn,1,1
  'set teacherTeachCourseRs = conn.Execute(sqlString)
  while not teacherTeachCourseRs.EOF
    '删除该选修课程的信息记录
	  sqlString = "delete from [publicCourseInfo] where courseNumber='" & teacherTeachCourseRs("courseNumber") & "'"
	  conn.Execute(sqlString)
	  '删除该选修课程的成绩信息
	  sqlString = "delete from [scoreInfo] where courseNumber='" & teacherTeachCourseRs("courseNumber") & "' and isSelect=1"
	  conn.Execute(sqlString)
	  '删除该选修课的学生选课信息
	  sqlString = "delete from [studentSelectCourseInfo] where courseNumber='" & teacherTeachCourseRs("courseNumber") & "'"
	  conn.Execute(sqlString)
	  teacherTeachCourseRs.MoveNext
  wend
  '删除该选修课的授课信息
  sqlString = "delete from [publicCourseTeach] where teacherNumber='" & teacherNumber & "'"
  conn.Execute(sqlString)
  teacherTeachCourseRs.Close
  '最后删除教师的个人信息记录
  sqlString = "delete from [teacherInfo] where teacherNumber='" & teacherNumber & "'"
  conn.Execute(sqlString)
  
  Response.Write "<script>alert('教师信息删除成功!');location.href='teacherInfoManage.asp';</script>"
%>
