<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  dim sqlString,courseNumber,courseName,specialFieldNumber,termId,specialFieldName,termName
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'ȡ��Ҫ�ſεĿγ̱��
	courseNumber = Request("courseNumber")
	sqlString = "select termId,specialFieldNumber,courseName from [publicCourseInfo] where courseNumber='" & courseNumber & "'"
	set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
	publicCourseInfoRs.Open sqlString,conn,1,1
	'ȡ�øÿγ̵Ŀγ�����,רҵ��ź�����ѧ�ڱ��
	if not publicCourseInfoRs.EOF then
	  courseName = publicCourseInfoRs("courseName")
	  specialFieldNumber = publicCourseInfoRs("specialFieldNumber")
	  termId = CInt(publicCourseInfoRs("termId"))
	end if
	'����רҵ���ȡ��רҵ����
	specialFieldName = GetSpecialFieldNameByNumber(specialFieldNumber)
	'����ѧ�ڱ��ȡ��ѧ������
	termName = GetTermnameById(termId)
%>
<HTML>
<HEAD>
	<Title>רҵ��ѡ�γ��ſ�</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>�ſ���Ϣ����--&gt;ѡ�޿γ��ſ���Ϣ�б�
			 </td>
	   </tr>
		<tr>
		  <td colspan=2 style="color:red;">
		  <%
		     Response.Write specialFieldName & " רҵ " & termName & " " &  courseName & "�γ̵��ſ���Ϣ"
		  %>
		  </td>
		</tr>
		<%
		  '���ݿγ̱�ŵõ��ÿγ̵��ſ���Ϣ
		  sqlString = "select * from [publicCourseTeach] where courseNumber='" & courseNumber & "'"
		  set publicCourseTeachRs = Server.CreateObject("ADODB.RecordSet")
		  publicCourseTeachRs.Open sqlString,conn,1,1
		  if not publicCourseTeachRs.EOF then
		    Response.Write "<tr><td>���ν�ʦ</td><td>�Ͽεص�</td><td>�Ͽ�ʱ��</td><td>����</td></tr>"
		  else
		    Response.Write "<tr><td align=center colspan=4>�Բ���,��û�иÿγ̵��ſ���Ϣ!</td></tr>"
		  end if
		  while not publicCourseTeachRs.EOF
		    Response.Write "<tr><td>" & GetTeacherNameByNumber(publicCourseTeachRs("teacherNumber")) & "</td><td>" & publicCourseTeachRs("teachClassRoom") & "</td><td>����" & publicCourseTeachRs("teachDay") & vbCrlf
			  Response.Write "</td><td><a href='publicCourseTeachDetail.asp?teachId=" & publicCourseTeachRs("teachId") & "'><img src='../images/edit.gif' height=12 width=12 border=0>��ϸ</a>&nbsp;<a href='publicCourseTeachDel.asp?teachId=" & publicCourseTeachRs("teachId") & "' onclick=" & """" & "javascript:return confirm('���ȷ��ɾ���˼�¼��?')" & """" & "><img src='../images/delete.gif' height=12 width=12 border=0>ɾ��</a></td></tr>" & vbCrlf
			  publicCourseTeachRs.MoveNext
		  wend
		  publicCourseTeachRs.Close
		%>
	  <tr><td colspan=4><input type="button" value="����µ��ſ���Ϣ" onClick="javascript:location.href='publicCourseTeachAdd.asp?courseNumber=<%=courseNumber%>'">
	      &nbsp;<input type="button" value="����ѡ��༶�γ���Ϣ"  onClick="javascript:location.href='publicCourseTeachMakeFirst.asp?termId=<%=termId%>&specialFieldNumber=<%=specialFieldNumber%>&submit=true'">
		  </td></tr>
	</table>
</BODY>
</HTML>


