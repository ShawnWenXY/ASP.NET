<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  dim  sqlString,courseNumber,courseName,classNumber,className,termId,termName
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'ȡ���ſεİ༶�γ̵ı��
	courseNumber = Request("courseNumber")
	'ȡ�øÿγ̵İ༶��ź�ѧ�ڱ�ŵ���Ϣ
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

'�ж���ĳ��ĳ�����ҵĸýڿ��Ƿ��Ѿ����ſ���
Function GetConflictString(timeString)
  dim sqlString,OneClassTeachInfoRs,conflictString,timeFormatString
  select case timeString
    case "MorningOne"
	     timeFormatString = "�����һ�ڿ�"
		case "MorningTwo"
		   timeFormatString = "����ڶ��ڿ�"
		case "MorningThree"
		   timeFormatString = "��������ڿ�"
		case "MorningFour"
		   timeFormatString = "������Ľڿ�"
		case "MorningFive"
		   timeFormatString = "�������ڿ�"
		case "AfternoonOne"
			 timeFormatString = "�����һ�ڿ�"
	  case "AfternoonTwo"
	     timeFormatString = "����ڶ��ڿ�"
		case "AfternoonThree"
		   timeFormatString = "��������ڿ�"
		case "AfternoonFour"
		   timeFormatString = "������Ľڿ�"
		case "EveningOne"
		   timeFormatString = "���ϵ�һ�ڿ�"
		case "EveningTwo"
			 timeFormatString = "���ϵڶ��ڿ�"
		case "EveningThree"
		  timeFormatString = "���ϵ����ڿ�"
		case else
		  timeFormatString = ""
  end select
  conflictString = ""
  '����Ĵ����ѯ�ڵ��쵱�ڿθõص��Ƿ������γ��ſ���
  sqlString = "select * from [classCourseTeach] where termId=" & CInt(termId) & " and teachClassRoom='" & Request("teachClassRoom") & "' and teachDay=" & CInt(Request("teachDay")) & " and " & timeString & "=1"
	set OneClassTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	OneClassTeachInfoRs.Open sqlString,conn,1,1
	if not OneClassTeachInfoRs.EOF then
	  conflictString = conflictString & "����" & Request("teachClassRoom") & "����" & Request("teachDay") & timeFormatString & "�ѱ�ռ��!"
	end if
	OneClassTeachInfoRs.Close
	Set OneClassTeachInfoRs = nothing
	sqlString = "select * from [publicCourseTeach] where termId=" & CInt(termId) & " and teachClassRoom='" & Request("teachClassRoom") & "' and teachDay=" & CInt(Request("teachDay")) & " and " & timeString & "=1"
	set OneClassTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	OneClassTeachInfoRs.Open sqlString,conn,1,1
	if not OneClassTeachInfoRs.EOF then
	  conflictString = conflictString & "����" & Request("teachClassRoom") & "����" & Request("teachDay") & timeFormatString & "�ѱ�ռ��!"
	end if
	OneClassTeachInfoRs.Close
	Set OneClassTeachInfoRs = nothing
	'�����ѯĳ���༶�ſε�ʱ���Ƿ����˳�ͻ
	sqlString = "select * from [classCourseTeach] where termId=" &CInt(termId) & " and classNumber='" & classNumber & "' and teachDay=" & CInt(Request("teachDay")) & " and " & timeString & "=1"
	set OneClassTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	OneClassTeachInfoRs.Open sqlString,conn,1,1
	if not OneClassTeachInfoRs.EOF then
	  conflictString = conflictString & "�ð༶������" & Request("teachDay") & timeFormatString & "�Ѿ����˿γ���!"
	end if
	GetConflictString = conflictString
End Function
	
	'�����Ҫ��ÿγ�����µ��ſ���Ϣ
	if Request("submit") <> "" then
	  '����Ƿ�ѡ���˽�ʦ
	  if Request("teacherNumber") = "" then
	    Response.Write "<script>alert('��ѡ���ڿν�ʦ!');</script>"
		elseif Request("teachClassRoom") = "" then
		  Response.Write "<script>alert('��ѡ���Ͽεص�!');</script>"
		elseif Request("teachDay") = "" then
		  Response.Write "<script>alert('��ѡ���Ͽ�ʱ��!');</script>"
		else
		  '��֤ͨ��,����ȡ���Ͽ�ʱ�����ϸ��Ϣ
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
		  '�����ſγ�ͻ����Ϣ
		  dim conflictMessage,conflictString
		  conflictMessage = ""
		  '���ѡ���������һ�ڿ���Ҫ�ж��Ƿ��ſγ�ͻ
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
		  '���û���ſγ�ͻ
		  if conflictMessage <> "" then
		    Response.Write "<script>alert('" & conflictMessage & "');</script>"
			else
		    '���濪ʼ���µ��Ͽ���Ϣ���뵽���ݿ���
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
		    Response.Write "<script>alert('�γ��Ͽ���Ϣ��ӳɹ�!');</script>"
			end if
		end if
	end if
%>
<HTML>
<HEAD>
	<Title>�γ��ſ���Ϣ���</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
   <input type="hidden" name=courseNumber value=<%=Request("courseNumber")%>>
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�ſ���Ϣ����--&gt;�༶�ſ���Ϣ���
			 </td>
	   </tr>
		 <tr>
		   <td width=100 align="right">�γ���Ϣ:</td>
			 <td><% Response.Write className & " " & termName & " " & courseName%></td>
		 </tr>
		 <tr>
		   <td width=100px align="right">ѡ���ʦ:</td>
			 <td>
			   <select name=teacherNumber>
				   <option value="">��ѡ��</option>
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
		   <td width=100px align="right">ѡ���Ͽεص�:</td>
			 <td>
			   <select name=teachClassRoom>
				   <option value="">��ѡ��</option>
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
		   <td width=100px align="right">�Ͽ�ʱ��:</td>
			 <td>
			   <select name=teachDay>
				   <option value="">��ѡ��</option>
				   <option value="1">����һ</option>
					 <option value="2">���ڶ�</option>
					 <option value="3">������</option>
					 <option value="4">������</option>
					 <option value="5">������</option>
					 <option value="6">������</option>
					 <option value="7">������</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">��ϸ�Ͽ���Ϣ</td>
			 <td>
			   &nbsp;<input type="checkbox" name=MorningOne value="1">�����һ��
				 &nbsp;<input type="checkbox" name=MorningTwo value="1">����ڶ���<br>
				 &nbsp;<input type="checkbox" name=MorningThree value="1">���������
				 &nbsp;<input type="checkbox" name=MorningFour value="1">������Ľ�<br>
				 &nbsp;<input type="checkbox" name=MorningFive value="1">��������
				 &nbsp;<input type="checkbox" name=AfternoonOne value="1">�����һ��<br>
				 &nbsp;<input type="checkbox" name=AfternoonTwo value="1">����ڶ���
				 &nbsp;<input type="checkbox" name=AfternoonThree value="1">���������<br>
				 &nbsp;<input type="checkbox" name=AfternoonFour value="1">������Ľ�
				 &nbsp;<input type="checkbox" name=EveningOne value="1">���ϵ�һ��<br>
				 &nbsp;<input type="checkbox" name=EveningTwo value="1">���ϵڶ���
				 &nbsp;<input type="checkbox" name=EveningThree value="1">���ϵ�����
			 </td>
		 </tr>
		 <tr bgcolor="#ffffff">
        <td height="30" colspan="2" align="center">
		      <input name="submit"  type="submit" value=" ȷ����� "> &nbsp;
				 <input type="button" value="����" onClick="javascript:location.href='classCourseTeachMakeSecond.asp?courseNumber=<%=courseNumber%>'">
		    </td>
      </tr>
	 </table>
 </form>
</body>
</html>