<!--#include file="../DataBase/conn.asp"-->
<!--#include file="../System/config.asp"-->
<!--#include file="../System/function.asp"-->
<%
  '���ѧ����û�е�½
  if session("studentNumber")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'�жϹ���Ա�Ƿ񿪷���ѡ��
	if canSelect = 0 then
	  Response.Write "<script>alert('�Բ���,����û�п���ѡ�ι���!');location.href='../System/systemInfo.asp';</script>"
	else
	  if Now < selectStartTime or Now > selectEndTime then
	    Response.Write "<script>alert('����ѡ��ʱ���ڽ���ѡ��!');location.href='../System/systemInfo.asp';</script>"
	  end if
	end if
	'���Ҫ�����ѡ�α�
	if  Request("submit") <> "" then
	  '����ɾ����ѧ����רҵ��ѧ�ڵ�����ѡ�޿γ̵�ѡ�޿γ�
	  sqlString = "delete from [studentSelectCourseInfo] where studentNumber='" & session("studentNumber") & "' and courseNumber in (select courseNumber from [publicCourseInfo] where specialFieldNumber='" & GetSpecialFieldNumberByStudentNumber(session("studentNumber")) & "' and termId=" & termId & ")"
	  conn.Execute(sqlString)
	  'ȡ��ѡ�޿εı��
	  courseNumbers = Request("courseNumbers")
	  if courseNumbers <> "" then
	    courseNumbers = Split(courseNumbers,",")
		  for each courseNumber in courseNumbers
		    sqlString = "insert into [studentSelectCourseInfo] values ('" & session("studentNumber") & "','" & Trim(courseNumber) & "')"
			  conn.Execute(sqlString)
		  next
	  end if
	  Response.Write "<script>alert('ѡ�γɹ�!');</script>"
	end if
%>
<HTML>
<HEAD>
	<Title>ѧ��ѡ��</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	<table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#FFFFFF'>
    <form name="form1" method="post" action="courseSelect.asp">
      <tr style="color:blue;font-size:14px;">
	      <td style="height: 14px" colspan=5 align="center">
		      <img src="../images/list.gif" width=14px height=14px>ѡ����Ϣ����--&gt;ѧ��ѡ��
			 </td>
	   </tr>
		 <tr><td colspan=5><%=GetTermnameById(termId)%>&nbsp;&nbsp;ѡ�޿γ���Ϣ</td></tr>
		 <%
		   sqlString = "select specialFieldNumber from [specialFieldInfo],[classInfo],[studentInfo] where [specialFieldInfo].specialFieldNumber = [classInfo].classSpecialFieldNumber and [classInfo].classNumber = [studentInfo].studentClassNumber and [studentInfo].studentNumber='" & Session("studentNumber") & "'"
			 set studentSpecialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
			 studentSpecialFieldInfoRs.Open sqlString,conn,1,1
			 specialFieldNumber = studentSpecialFieldInfoRs("specialFieldNumber")
			 studentSpecialFieldInfoRs.Close
			 '�õ���רҵ��ѧ�ڵ�����ѡ�޿γ���Ϣ
			 sqlString = "select * from [publicCourseInfo] where specialFieldNumber='" & specialFieldNumber & "' and termId=" & termId
			 set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
			 publicCourseInfoRs.Open sqlString,conn,1,1
			 if not publicCourseInfoRs.EOF then
			   Response.Write "<tr><td>ѡ��</td><td>�γ̱��</td><td>�γ�����</td><td>�γ�ѧ��</td><td>�Ͽ���ʦ</td></tr>"
			 else
			   Response.Write "<tr><td colspan=5>�Բ���,��û��ѡ�޿γ���Ϣ</td></tr>"
			 end if
			 while not publicCourseInfoRs.EOF
			   Response.Write "<tr><td><input type=checkbox name=courseNumbers value='" & publicCourseInfoRs("courseNumber") & "'></td><td>" & publicCourseInfoRs("courseNumber") & "</td><td>" & publicCourseInfoRs("courseName") & "</td><td>" & publicCourseInfoRs("courseScore") & "</td><td>" & GetTeacherNameByPublicCourseNumber(publicCourseInfoRs("courseNumber")) & "</td></tr>"
			   publicCourseInfoRs.MoveNext
			wend
			publicCourseInfoRs.Close
		 %>
	  <tr bgcolor="#ffffff">
        <td height="30" colspan="5" align="center">
		      <input name="submit"  type="submit" value=" ѡ�� ">
				  <input type="reset" value=" ������д ">
		    </td>
	 </tr>
	</form>
 </table>
 <table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#FFFFFF'>
      <tr style="color:blue;font-size:14px;">
	      <td style="height: 14px" colspan=5 align="center">
		      <img src="../images/list.gif" width=14px height=14px><%=GetTermNameById(termId)%>&nbsp;���ѡ����Ϣ
			 </td>
	   </tr>
		<%
		  sqlString = "select [studentSelectCourseInfo].courseNumber,[publicCourseInfo].courseName from [studentSelectCourseInfo],[publicCourseInfo] where [studentSelectCourseInfo].studentNumber='" & session("studentNumber") & "' and [studentSelectCourseInfo].courseNumber=[publicCourseInfo].courseNumber and [publicCourseInfo].termId=" & termId
		  set selectCourseRs = Server.CreateObject("ADODB.RecordSet")
		  selectCourseRs.Open sqlString,conn,1,1
		  if selectCourseRs.EOF then
		    Response.Write "<tr><td colspan=2 align=center>�㻹û��ѡ��</td></tr>"
			end if
		  while not selectCourseRs.EOF
		    Response.Write "<tr><td>" & selectCourseRs("courseNumber") & "</td><td>" & selectCourseRs("courseName") & "</td></tr>"
	      selectCourseRs.MoveNext	 
		  wend
		%>
</table>
</body>
</html>