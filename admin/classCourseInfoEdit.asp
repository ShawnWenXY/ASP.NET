<!--#include virtual="/DataBase/conn.asp"-->
<%
  'errMessage���������Ϣ
  dim errMessage
  errMessage = ""
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'�������Ա�޸��˿γ���Ϣ���ύ
	if Request("submit") <> "" then
	  '���û������γ�����
	  if Request("courseName") = "" then
	    errMessage = "������γ�����!"
	  end if
	  
	  '���ݴ�����ϢerrMessage���ݾ����Ƿ�ִ�а༶�γ���Ϣ���޸Ĳ���
	  if errMessage = "" then
	    sqlString = "select * from [classCourseInfo] where courseNumber='" & Request("courseNumber") & "'"
		  set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  classCourseInfoRs.Open sqlString,conn,1,3
		  classCourseInfoRs("courseName") = Request("courseName")
		  classCourseInfoRs("classNumber") = Request("classNumber")
		  classCourseInfoRs("termId") = CInt(Request("termId"))
		  classCourseInfoRs("courseScore") = CSng(Request("courseScore"))
		  classCourseInfoRs("courseMemo") = Request("courseMemo")
		  classCourseInfoRs.Update
		  Response.Write "<script>alert('�༶�γ���Ϣ�޸ĳɹ�!');</script>"
	  else
	    Response.Write "<script>alert('" & errMessage & "');</script>"
	  end if
	end if
	
	sqlString = "select * from [classCourseInfo] where courseNumber='" & Request("courseNumber") & "'"
	set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
	classCourseInfoRs.Open sqlString,conn,1,1
%>

<HTML>
<HEAD>
	<Title>�༶�γ���Ϣ�޸�</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�γ���Ϣ����--&gt;�༶�γ���Ϣ�޸�
			 </td>
	   </tr>
		 <tr>
		    <td width=100 align="right">�γ̱��:</td>
		    <td><%=Request("courseNumber")%><input type=hidden name=courseNumber size=20  value='<%=Request("courseNumber")%>'></td>
		  </tr>
		<tr>
		    <td width=100 align="right">����ѧ��:</td>
		    <td>
			    <select name=termId>
					  <%
						  sqlString = "select * from termInfo"
						  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
						  termInfoRs.Open sqlString,conn,1,1
						  while not termInfoRs.EOF
						    selected = ""
							  if termInfoRs("termId") = classCourseInfoRs("termId") then
							    selected = "selected"
								end if
						    Response.Write "<option value='" & termInfoRs("termId") & "' " & selected & ">" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "��" & termInfoRs("termUpOrDown") & "</option>"
							  termInfoRs.MoveNext
						  wend
					  %>
				  </select>
			  </td>
		  </tr>
		<tr>
		   <td width=100px align="right">���ڰ༶:</td>
			 <td>
			   <select name=classNumber>				
				<%
					  sqlString = "select classNumber,className from [classInfo]"
					  set classInfoRs = Server.CreateObject("ADODB.RecordSet")
					  classInfoRs.Open sqlString,conn,1,1
					  while not classInfoRs.EOF
					    selected = ""
						  if classInfoRs("classNumber") = classCourseInfoRs("classNumber") then
						    selected = "selected"
							end if
					    Response.Write "<option value='" & classInfoRs("classNumber") & "' " & selected & ">" & classInfoRs("className") & "</option>"
						  classInfoRs.MoveNext
					  wend
				  %>
				 </select>
			 </td>
		 </tr>
		  <tr>
		    <td width=100 align="right">�γ�����:</td>
		    <td><input type=text name=courseName size=20 value='<%=classCourseInfoRs("courseName")%>'></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">�γ�ѧ��:</td>
		    <td><input type=text name=courseScore size=5 value=<%=classCourseInfoRs("courseScore")%>>��</td>
		  </tr>
		  <tr>
		    <td width=100 align="right">������Ϣ:</td>
		    <td><textarea cols=40 rows=5 name=courseMemo><%=classCourseInfoRs("courseMemo")%></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" ȷ�ϸ���">
		      <input type=button value="����" onClick="javascript:location.href='classCourseInfoManage.asp'">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>
</HTML>
