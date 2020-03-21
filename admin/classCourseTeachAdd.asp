<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  dim  sqlString,courseNumber,courseName,classNumber,className,termId,termName
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'取得排课的班级课程的编号
	courseNumber = Request("courseNumber")
	'取得该课程的班级编号和学期编号等信息
	sqlString = "select * from [classCourseInfo] where courseNumber='" & courseNumber & "'"
	set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
	classCourseInfoRs.Open sqlString,conn,1,1
	if not classCourseInfoRs.EOF then
	  courseName = classCourseInfoRs("courseName")
	  classNumber = classCourseInfoRs("classNumber")
	  className = GetClassNameByNumber(classNumber)
	  termId = classCourseInfoRs("termId")
	  termName = GetTermnameById(termId)
	end if

'判断在某天某个教室的该节课是否已经被排课了
Function GetConflictString(timeString)
  dim sqlString,OneClassTeachInfoRs,conflictString,timeFormatString
  select case timeString
    case "MorningOne"
	     timeFormatString = "上午第一节课"
		case "MorningTwo"
		   timeFormatString = "上午第二节课"
		case "MorningThree"
		   timeFormatString = "上午第三节课"
		case "MorningFour"
		   timeFormatString = "上午第四节课"
		case "MorningFive"
		   timeFormatString = "上午第五节课"
		case "AfternoonOne"
			 timeFormatString = "下午第一节课"
	  case "AfternoonTwo"
	     timeFormatString = "下午第二节课"
		case "AfternoonThree"
		   timeFormatString = "下午第三节课"
		case "AfternoonFour"
		   timeFormatString = "下午第四节课"
		case "EveningOne"
		   timeFormatString = "晚上第一节课"
		case "EveningTwo"
			 timeFormatString = "晚上第二节课"
		case "EveningThree"
		  timeFormatString = "晚上第三节课"
		case else
		  timeFormatString = ""
  end select
  conflictString = ""
  '下面的代码查询在当天当节课该地点是否被其他课程排课了
  sqlString = "select * from [classCourseTeach] where termId=" & CInt(termId) & " and teachClassRoom='" & Request("teachClassRoom") & "' and teachDay=" & CInt(Request("teachDay")) & " and " & timeString & "=1"
	set OneClassTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	OneClassTeachInfoRs.Open sqlString,conn,1,1
	if not OneClassTeachInfoRs.EOF then
	  conflictString = conflictString & "教室" & Request("teachClassRoom") & "星期" & Request("teachDay") & timeFormatString & "已被占用!"
	end if
	OneClassTeachInfoRs.Close
	Set OneClassTeachInfoRs = nothing
	sqlString = "select * from [publicCourseTeach] where termId=" & CInt(termId) & " and teachClassRoom='" & Request("teachClassRoom") & "' and teachDay=" & CInt(Request("teachDay")) & " and " & timeString & "=1"
	set OneClassTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	OneClassTeachInfoRs.Open sqlString,conn,1,1
	if not OneClassTeachInfoRs.EOF then
	  conflictString = conflictString & "教室" & Request("teachClassRoom") & "星期" & Request("teachDay") & timeFormatString & "已被占用!"
	end if
	OneClassTeachInfoRs.Close
	Set OneClassTeachInfoRs = nothing
	'下面查询某个班级排课的时间是否发生了冲突
	sqlString = "select * from [classCourseTeach] where termId=" &CInt(termId) & " and classNumber='" & classNumber & "' and teachDay=" & CInt(Request("teachDay")) & " and " & timeString & "=1"
	set OneClassTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	OneClassTeachInfoRs.Open sqlString,conn,1,1
	if not OneClassTeachInfoRs.EOF then
	  conflictString = conflictString & "该班级在星期" & Request("teachDay") & timeFormatString & "已经排了课程了!"
	end if
	GetConflictString = conflictString
End Function
	
	'如果需要向该课程添加新的排课信息
	if Request("submit") <> "" then
	  '检查是否选择了教师
	  if Request("teacherNumber") = "" then
	    Response.Write "<script>alert('请选择授课教师!');</script>"
		elseif Request("teachClassRoom") = "" then
		  Response.Write "<script>alert('请选择上课地点!');</script>"
		elseif Request("teachDay") = "" then
		  Response.Write "<script>alert('请选择上课时间!');</script>"
		else
		  '验证通过,首先取得上课时间的详细信息
		  dim MorningOne,MorningTwo,MorningThree,MorningFour,MorningFive
			dim AfternoonOne,AfternoonTwo,AfternoonThree,AfternoonFour
	    dim EveningOne,EvenigTwo,EveningThree
	    if Request("MorningOne") = "1" then
	      MorningOne = 1
		  else
		    MorningOne = 0
		  end if
		  if Request("MorningTwo") = "1" then
		    MorningTwo = 1
			else
			  MorningTwo = 0
			end if
			if Request("MorningThree") = "1" then
			  MorningThree = 1
			else
			  MorningThree = 0
			end if
			if Request("MorningFour") = "1" then
			  MorningFour = 1
			else
			  MorningFour = 0
			end if
			if Request("MorningFive") = "1" then
			  MorningFive = 1
			else
			  MorningFive = 0
			end if
			if Request("AfternoonOne") = "1" then
			  AfternoonOne = 1
			else
			  AfternoonOne = 0
			end if
			if Request("AfternoonTwo") = "1" then
			  AfternoonTwo = 1
			else
			  AfternoonTwo = 0
			end if
			if Request("AfternoonThree") = "1" then
			  AfternoonThree = 1
			else
			  AfternoonThree = 0
			end if
			if Request("AfternoonFour") = "1" then
			  AfternoonFour = 1
			else
			  AfternoonFour = 0
			end if
			if Request("EveningOne") = "1" then
			  EveningOne = 1
			else
			  EveningOne = 0
			end if
			if Request("EveningTwo") = "1" then
			  EveningTwo = 1
			else
			  EveningTwo = 0
			end if
			if Request("EveningThree") = "1" then
			  EveningThree = 1
			else
			  EveningThree = 0
		  end if
		  '保存排课冲突的信息
		  dim conflictMessage,conflictString
		  conflictMessage = ""
		  '如果选择了上午第一节课需要判断是否排课冲突
		  if MorningOne = 1 then
		    conflictString = GetConflictString("MorningOne")
			  if conflictString <> "" then
			    conflictMessage = conflictMessage & conflictString
				end if
		  end if
		  if MorningTwo = 1 then
		    conflictString = GetConflictString("MorningTwo")
			  if conflictString <> "" then
			    conflictMessage = conflictMessage & conflictString
				end if
		  end if
		   if MorningThree = 1 then
		    conflictString = GetConflictString("MorningThree")
			  if conflictString <> "" then
			    conflictMessage = conflictMessage & conflictString
				end if
		  end if
		  if MorningFour = 1 then
		    conflictString = GetConflictString("MorningFour")
			  if conflictString <> "" then
			    conflictMessage = conflictMessage & conflictString
				end if
		  end if
		  if MorningFive = 1 then
		    conflictString = GetConflictString("MorningFive")
			  if conflictString <> "" then
			    conflictMessage = conflictMessage & conflictString
				end if
		  end if
		  if AfternoonOne = 1 then
		    conflictString = GetConflictString("AfternoonOne")
			  if conflictString <> "" then
			    conflictMessage = conflictMessage & conflictString
				end if
		  end if
		  if AfternoonTwo = 1 then
		    conflictString = GetConflictString("AfternoonTwo")
			  if conflictString <> "" then
			    conflictMessage = conflictMessage & conflictString
				end if
		  end if
		  if AfternoonThree = 1 then
		    conflictString = GetConflictString("AfternoonThree")
			  if conflictString <> "" then
			    conflictMessage = conflictMessage & conflictString
				end if
		  end if
		  if AfternoonFour = 1 then
		    conflictString = GetConflictString("AfternoonFour")
			  if conflictString <> "" then
			    conflictMessage = conflictMessage & conflictString
				end if
		  end if
		  if EveningOne = 1 then
		    conflictString = GetConflictString("EveningOne")
			  if conflictString <> "" then
			    conflictMessage = conflictMessage & conflictString
				end if
		  end if
		  if EveningTwo = 1 then
		    conflictString = GetConflictString("EveningTwo")
			  if conflictString <> "" then
			    conflictMessage = conflictMessage & conflictString
				end if
		  end if
		  if EveningThree = 1 then
		    conflictString = GetConflictString("EveningThree")
			  if conflictString <> "" then
			    conflictMessage = conflictMessage & conflictString
				end if
		  end if
		  '如果没有排课冲突
		  if conflictMessage <> "" then
		    Response.Write "<script>alert('" & conflictMessage & "');</script>"
			else
		    '下面开始将新的上课信息加入到数据库中
		    sqlString = "select * from [classCourseTeach]"
		    set classCourseTeachRs = Server.CreateObject("ADODB.RecordSet")
		    classCourseTeachRs.Open sqlString,conn,1,3
		    classCourseTeachRs.AddNew
		    classCourseTeachRs("courseNumber") = courseNumber
		    classCourseTeachRs("termId") = CInt(termId)
			  classCourseTeachRs("classNumber") = classNumber
		    classCourseTeachRs("teacherNumber") = Request("teacherNumber")
		    classCourseTeachRs("teachClassRoom") = Request("teachClassRoom")
		    classCourseTeachRs("teachDay") = CInt(Request("teachDay"))
		    classCourseTeachRs("MorningOne") = MorningOne
		    classCourseTeachRs("MorningTwo") = MorningTwo
		    classCourseTeachRs("MorningThree") = MorningThree
		    classCourseTeachRs("MorningFour") = MorningFour
		    classCourseTeachRs("MorningFive") = MorningFive
		    classCourseTeachRs("AfternoonOne") = AfternoonOne
		    classCourseTeachRs("AfternoonTwo") = AfternoonTwo
		    classCourseTeachRs("AfternoonThree") = AfternoonThree
		    classCourseTeachRs("AfternoonFour") = AfternoonFour
		    classCourseTeachRs("EveningOne") = EveningOne
		    classCourseTeachRs("EveningTwo") = EveningTwo
		    classCourseTeachRs("EveningThree") = EveningThree
		    classCourseTeachRs.Update
		    Response.Write "<script>alert('课程上课信息添加成功!');</script>"
			end if
		end if
	end if
%>
<HTML>
<HEAD>
	<Title>课程排课信息添加</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
   <input type="hidden" name=courseNumber value=<%=Request("courseNumber")%>>
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>排课信息管理--&gt;班级排课信息添加
			 </td>
	   </tr>
		 <tr>
		   <td width=100 align="right">课程信息:</td>
			 <td><% Response.Write className & " " & termName & " " & courseName%></td>
		 </tr>
		 <tr>
		   <td width=100px align="right">选择教师:</td>
			 <td>
			   <select name=teacherNumber>
				   <option value="">请选择</option>
				   <%
					   sqlString = "select teacherNumber,teacherName from [teacherInfo] order by teacherNumber"
						 set teacherInfoRs = Server.CreateObject("ADODB.RecordSet")
						 teacherInfoRs.Open sqlString,conn,1,1
						 while not teacherInfoRs.EOF
						   Response.Write "<option value='" & teacherInfoRs("teacherNumber") & "'>" & teacherInfoRs("teacherNumber") & "--" & teacherInfoRs("teacherName") & "</option>"
							 teacherInfoRs.MoveNext
						 wend
					 %>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">选择上课地点:</td>
			 <td>
			   <select name=teachClassRoom>
				   <option value="">请选择</option>
				   <%
					   sqlString = "select classRoomName from [classRoomInfo] order by classRoomName"
						 set classRoomInfoRs = Server.CreateObject("ADODB.RecordSet")
						 classRoomInfoRs.Open sqlString,conn,1,1
						 while not classRoomInfoRs.EOF
						   Response.Write "<option value='" & classRoomInfoRs("classRoomName") & "'>" & classRoomInfoRs("classRoomName") & "</option>"
							 classRoomInfoRs.MoveNext
						 wend
						 classRoomInfoRs.Close
					 %>
				 </select>
			 </td>
		 </tr>
		  <tr>
		   <td width=100px align="right">上课时间:</td>
			 <td>
			   <select name=teachDay>
				   <option value="">请选择</option>
				   <option value="1">星期一</option>
					 <option value="2">星期二</option>
					 <option value="3">星期三</option>
					 <option value="4">星期四</option>
					 <option value="5">星期五</option>
					 <option value="6">星期六</option>
					 <option value="7">星期日</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">详细上课信息</td>
			 <td>
			   &nbsp;<input type="checkbox" name=MorningOne value="1">上午第一节
				 &nbsp;<input type="checkbox" name=MorningTwo value="1">上午第二节<br>
				 &nbsp;<input type="checkbox" name=MorningThree value="1">上午第三节
				 &nbsp;<input type="checkbox" name=MorningFour value="1">上午第四节<br>
				 &nbsp;<input type="checkbox" name=MorningFive value="1">上午第五节
				 &nbsp;<input type="checkbox" name=AfternoonOne value="1">下午第一节<br>
				 &nbsp;<input type="checkbox" name=AfternoonTwo value="1">下午第二节
				 &nbsp;<input type="checkbox" name=AfternoonThree value="1">下午第三节<br>
				 &nbsp;<input type="checkbox" name=AfternoonFour value="1">下午第四节
				 &nbsp;<input type="checkbox" name=EveningOne value="1">晚上第一节<br>
				 &nbsp;<input type="checkbox" name=EveningTwo value="1">晚上第二节
				 &nbsp;<input type="checkbox" name=EveningThree value="1">晚上第三节
			 </td>
		 </tr>
		 <tr bgcolor="#ffffff">
        <td height="30" colspan="2" align="center">
		      <input name="submit"  type="submit" value=" 确认添加 "> &nbsp;
				 <input type="button" value="返回" onClick="javascript:location.href='classCourseTeachMakeSecond.asp?courseNumber=<%=courseNumber%>'">
		    </td>
      </tr>
	 </table>
 </form>
</body>
</html>