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
	<Title>�༶�γ��ſ�</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�ſ���Ϣ����--&gt;ѡ���ſεİ༶
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
		    <td colspan=2>&nbsp;&nbsp;ѡ��༶:&nbsp;
			    <select name=classNumber>
				    <option value="">��ѡ��</option>
					  <%
						  sqlString = "select classNumber,className from [classInfo]"
						  set classInfoRs = Server.CreateObject("ADODB.RecordSet")
						  classInfoRs.Open sqlString,conn,1,1
						  while not classInfoRs.EOF
						    Response.Write "<option value='" & classInfoRs("classNumber") & "'>" & classInfoRs("className") & "</option>"
							  classInfoRs.MoveNext
						  wend
					  %>
				  </select>
			  </td>
		  </tr>
		  <tr>
		    <td colspan=2>&nbsp;&nbsp;<input type="submit" name="submit" value="��ѯ�༶�γ�"></td>
		  </tr>
	  </table>
	  <br>
	  <table width=600 border=0 cellpadding=0 cellspacing=0 align="center">
		  <%
		    '���Ҫ���ѯĳ���༶ĳ��ѧ�ڵĿγ�
		    if Request("submit") <> "" then
			    '�ж��Ƿ�ѡ����ѧ��
				  if Request("termId") = "" then
					  Response.Write "<script>alert('��ѡ��ѧ��!');</script>"
					elseif Request("classNumber") = "" then
					  Response.Write "<script>alert('��ѡ��༶!');</script>"
					else
					  '��ѯ��ѧ�ڸð༶�����пγ�
					  sqlString = "select * from [classCourseInfo] where termId=" & Request("termId") & " and classNumber='" & Request("classNumber") & "'"
					  set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
					  classCourseInfoRs.Open sqlString,conn,1,1
					  if not classCourseInfoRs.EOF then
					    Response.Write "<tr><td colspan=4 style='color:red;' align=center>" & GetClassNameByNumber(Request("classNumber")) & " " & GetTermnameById(Request("termId")) & " �γ���Ϣ</td></tr>"
					    Response.Write "<tr><td>�γ̱��</td><td>�γ�����</td><td>�γ�ѧ��</td><td>����</td></tr>"
					  else
					    Response.Write "<tr></td><td colspan=4 style='color:red;' align=center>��û�а༶�γ���Ϣ</td></tr>"
					  end if
					  '���ÿ�ſγ̵���Ϣ
					  while not classCourseInfoRs.EOF
					    Response.Write "<tr><td>" & classCourseInfoRs("courseNumber") & "</td><td>" & classCourseInfoRs("courseName") & "</td><td align=center>" & classCourseInfoRs("courseScore") & "</td><td><a href='classCourseTeachMakeSecond.asp?courseNumber=" & classCourseInfoRs("courseNumber") & "'>�ſι���</a></td></tr>"
					    classCourseInfoRs.MoveNext
					  wend
					end if
			  end if
		  %>
	  </table>
  </form>
  
</BODY>

</HTML>
