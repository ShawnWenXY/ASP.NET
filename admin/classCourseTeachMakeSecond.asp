<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  dim sqlString,courseNumber,courseName,classNumber,termId,className,termName
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'ȡ��Ҫ�ſεĿγ̱��
	courseNumber = Request("courseNumber")
	sqlString = "select termId,classNumber,courseName from [classCourseInfo] where courseNumber='" & courseNumber & "'"
	set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
	classCourseInfoRs.Open sqlString,conn,1,1
	'ȡ�øÿγ̵Ŀγ�����,�༶��ź�����ѧ�ڱ��
	if not classCourseInfoRs.EOF then
	  courseName = classCourseInfoRs("courseName")
	  classNumber = classCourseInfoRs("classNumber")
	  termId = CInt(classCourseInfoRs("termId"))
	end if
	'���ݰ༶���ȡ�ð༶����
	className = GetClassNameByNumber(classNumber)
	'����ѧ�ڱ��ȡ��ѧ������
	termName = GetTermnameById(termId)
%>
<HTML>
<HEAD>
	<Title>�༶�γ��ſ�</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>�ſ���Ϣ����--&gt;�γ��ſ���Ϣ�б�
			 </td>
	   </tr>
		<tr>
		  <td colspan=2 style="color:red;">
		  <%
		     Response.Write className & " " & termName & " " &  courseName & "�γ̵��ſ���Ϣ"
		  %>
		  </td>
		</tr>
		<%
		  '���ݿγ̱�ŵõ��ÿγ̵��ſ���Ϣ
		  sqlString = "select * from [classCourseTeach] where courseNumber='" & courseNumber & "'"
		  set classCourseTeachRs = Server.CreateObject("ADODB.RecordSet")
		  classCourseTeachRs.Open sqlString,conn,1,1
		  if not classCourseTeachRs.EOF then
		    Response.Write "<tr><td>���ν�ʦ</td><td>�Ͽεص�</td><td>�Ͽ�ʱ��</td><td>����</td></tr>"
		  else
		    Response.Write "<tr><td align=center colspan=4>�Բ���,��û�иÿγ̵��ſ���Ϣ!</td></tr>"
		  end if
		  while not classCourseTeachRs.EOF
		    Response.Write "<tr><td>" & GetTeacherNameByNumber(classCourseTeachRs("teacherNumber")) & "</td><td>" & classCourseTeachRs("teachClassRoom") & "</td><td>����" & classCourseTeachRs("teachDay") & vbCrlf
			  Response.Write "</td><td><a href='classCourseTeachDetail.asp?teachId=" & classCourseTeachRs("teachId") & "'><img src='../images/edit.gif' height=12 width=12 border=0>��ϸ</a>&nbsp;<a href='classCourseTeachDel.asp?teachId=" & classCourseTeachRs("teachId") & "' onclick=" & """" & "javascript:return confirm('���ȷ��ɾ���˼�¼��?')" & """" & "><img src='../images/delete.gif' height=12 width=12 border=0>ɾ��</a></td></tr>" & vbCrlf
			  classCourseTeachRs.MoveNext
		  wend
		  classCourseTeachRs.Close
		%>
	  <tr><td colspan=4><input type="button" value="����µ��ſ���Ϣ" onClick="javascript:location.href='classCourseTeachAdd.asp?courseNumber=<%=courseNumber%>'">
	      &nbsp;<input type="button" value="����ѡ��༶�γ���Ϣ"  onClick="javascript:location.href='classCourseTeachMakeFirst.asp?termId=<%=termId%>&classNumber=<%=classNumber%>&submit=true'">
		  </td></tr>
	</table>
</BODY>
</HTML>


