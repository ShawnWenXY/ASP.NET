<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
%>
<HTML>
<HEAD>
	<Title>��ѡ�γ��ſ�</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�ſ���Ϣ����--&gt;ѡ��ѡ���ſε�רҵ
			 </td>
	   </tr><br>
		<tr>
		    <td colspan=2>&nbsp;&nbsp;ѡ��ѧ��:&nbsp;
			    <select name=termId>
				    <option value="">��ѡ��</option>
					  <%
					    dim sqlString
						  sqlString = "select * from termInfo"
						  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
						  termInfoRs.Open sqlString,conn,1,1
						  while not termInfoRs.EOF
						    Response.Write "<option value='" & termInfoRs("termId") & "'>" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "��" & termInfoRs("termUpOrDown") & "</option>"
							  termInfoRs.MoveNext
						  wend
					  %>
				  </select>
			  </td>
		  </tr>
		  <tr>
		    <td colspan=2>&nbsp;&nbsp;ѡ��רҵ:&nbsp;
			    <select name=specialFieldNumber>
				    <option value="">��ѡ��</option>
					  <%
						  sqlString = "select * from [specialFieldInfo]"
						  set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
						  specialFieldInfoRs.Open sqlString,conn,1,1
						  while not specialFieldInfoRs.EOF
						    Response.Write "<option value='" & specialFieldInfoRs("specialFieldNumber") & "'>" & specialFieldInfoRs("specialFieldName") & "</option>"
							  specialFieldInfoRs.MoveNext
						  wend
					  %>
				  </select>
			  </td>
		  </tr>
		  <tr>
		    <td colspan=2>&nbsp;&nbsp;<input type="submit" name="submit" value="��ѯרҵѡ�޿γ�"></td>
		  </tr>
	  </table>
	  <br>
	  <table width=600 border=0 cellpadding=0 cellspacing=0 align="center">
		  <%
		    '���Ҫ���ѯĳ��רҵĳ��ѧ�ڵ�ѡ�޿γ�
		    if Request("submit") <> "" then
			    '�ж��Ƿ�ѡ����ѧ��
				  if Request("termId") = "" then
					  Response.Write "<script>alert('��ѡ��ѧ��!');</script>"
					elseif Request("specialFieldNumber") = "" then
					  Response.Write "<script>alert('��ѡ��רҵ!');</script>"
					else
					  '��ѯ��ѧ�ڸ�רҵ������ѡ�޿γ�
					  sqlString = "select * from [publicCourseInfo] where termId=" & Request("termId") & " and specialFieldNumber='" & Request("specialFieldNumber") & "'"
					  set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
					  publicCourseInfoRs.Open sqlString,conn,1,1
					  if not publicCourseInfoRs.EOF then
					    Response.Write "<tr><td colspan=4 style='color:red;' align=center>" & GetSpecialFieldNameByNumber(Request("specialFieldNumber")) & " " & GetTermnameById(Request("termId")) & " ѡ�޿γ���Ϣ</td></tr>"
					    Response.Write "<tr><td>�γ̱��</td><td>�γ�����</td><td>�γ�ѧ��</td><td>����</td></tr>"
					  else
					    Response.Write "<tr></td><td colspan=4 style='color:red;' align=center>��û��רҵѡ�޿γ���Ϣ</td></tr>"
					  end if
					  '���ÿ�ſγ̵���Ϣ
					  while not publicCourseInfoRs.EOF
					    Response.Write "<tr><td>" & publicCourseInfoRs("courseNumber") & "</td><td>" & publicCourseInfoRs("courseName") & "</td><td align=center>" & publicCourseInfoRs("courseScore") & "</td><td><a href='publicCourseTeachMakeSecond.asp?courseNumber=" & publicCourseInfoRs("courseNumber") & "'>�ſι���</a></td></tr>"
					    publicCourseInfoRs.MoveNext
					  wend
					end if
			  end if
		  %>
	  </table>
  </form>
  
</BODY>

</HTML>
