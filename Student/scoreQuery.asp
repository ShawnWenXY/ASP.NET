<!--#include file="../DataBase/conn.asp"-->
<!--#include file="../System/function.asp"-->
<%
  '���ѧ����û�е�½
  if session("studentNumber")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
%>
<HTML>
<HEAD>
	<Title>ѧ���ɼ���ѯ</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	<table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#FFFFFF'>
    <form name="form1" method="post" action="scoreQuery.asp">
      <tr style="color:blue;font-size:14px;">
	      <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/list.gif" width=14px height=14px>�ɼ���Ϣ����--&gt;ѧ���ɼ���ѯ
			 </td>
	   </tr>
		 <tr>
		   <td>��ѡ��ѧ��:</td>
			 <td>
			   <select name=termId>
				    <option value="">��ѡ��</option>
					  <%
					    dim sqlString
						  sqlString = "select * from [termInfo]"
						  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
						  termInfoRs.Open sqlString,conn,1,1
						  while not termInfoRs.EOF
						    Response.Write "<option value='" & termInfoRs("termId") & "'>" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "��" & termInfoRs("termUpOrDown") & "</option>"
							  termInfoRs.MoveNext
						  wend
					  %>
				  </select>&nbsp;<input type="submit" name="submit" value="��ѯ�ɼ�">
				</td>
		 </tr>
		 <tr>
		   <td colspan=2>
			   <table width=100% border=0 bordercolor="maroon" cellspacing=0>
				  <%
				     if Request("termId") <> "" then
						    Response.Write "<tr><td colspan=8 align=center>" & GetTermNameById(Request("termId")) & " �ɼ���Ϣ</td></tr>"
							  Response.Write "<tr><td>�γ̱��</td><td>�γ�����</td><td>�γ�����</td><td>�ɼ�</td></tr>"
					 
						   '��ѯ���޿εĳɼ�
						   sqlString = "select [scoreInfo].courseNumber,[scoreInfo].score,[classCourseInfo].courseName from [scoreInfo],[classCourseInfo] where [scoreInfo].studentNumber='" & session("studentNumber") & "' and [scoreInfo].isSelect=0 and [scoreInfo].courseNumber=[classCourseInfo].courseNumber and [classCourseInfo].termId=" & Request("termId")
						   set scoreInfoRs = Server.CreateObject("ADODB.RecordSet")
						   scoreInfoRs.Open sqlString,conn,1,1
						   while not scoreInfoRs.EOF
						     Response.Write "<tr><td>" & scoreInfoRs("courseNumber") & "</td><td>" & scoreInfoRs("courseName") & "</td><td>���޿�</td><td>" & scoreInfoRs("score") & "</td></tr>"
							   scoreInfoRs.MoveNext
						   wend
						   scoreInfoRs.Close
						   '��ѯѡ�޿εĳɼ�
						   sqlString = "select [scoreInfo].courseNumber,[scoreInfo].score,[publicCourseInfo].courseName from [scoreInfo],[publicCourseInfo] where [scoreInfo].studentNumber='" & session("studentNumber") & "' and [scoreInfo].isSelect=1 and [scoreInfo].courseNumber=[publicCourseInfo].courseNumber and [publicCourseInfo].termId=" & Request("termId")
						   scoreInfoRs.Open sqlString,conn,1,1
						   while not scoreInfoRs.EOF
						     Response.Write "<tr><td>" & scoreInfoRs("courseNumber") & "</td><td>" & scoreInfoRs("courseName") & "</td><td>ѡ�޿�</td><td>" & scoreInfoRs("score") & "</td></tr>"
							   scoreInfoRs.MoveNext
						   wend
						   scoreInfoRs.Close
						end if
				  %>
				 </table>
		   </td>
		 </tr>
		</form>
	</table>
</body>
</html>
		