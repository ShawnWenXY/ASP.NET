<!--#include virtual="/Database/conn.asp"-->
<%
  'errMessage���������Ϣ
  dim errMessage
  errMessage = ""
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'�������Ա������µĿγ���Ϣ���ύ
	if Request("submit") <> "" then
	  '���û��ѡ�����ڵ�ѧ��
	  if Request("termId") = "" then
	    errMessage = "��ѡ�����ڵ�ѧ��!"
	  end if
	  '���û��ѡ��༶��Ϣ
	  if Request("classNumber") = "" then
	    errMessage = "��ѡ��γ����õİ༶!"
	  end if
	  '���û������γ̱��
	  if Request("courseNumber") = "" then
	    errMessage = "������γ̱��!"
	  end if
	  '���û������γ�����
	  if Request("courseName") = "" then
	    errMessage = "������γ�����!"
	  '����ѧ�ڸð༶�ÿγ���Ϣ�Ƿ��Ѿ�����
	  else
	    sqlString = "select * from [classCourseInfo] where classNumber='" & Request("classNumber") & "' and termId=" & Request("termId") & " and courseName='" & Request("courseName") & "'"
		  set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  classCourseInfoRs.Open sqlString,conn,1,1
		  if not classCourseInfoRs.EOF then
		    errMessage = "��ѧ�ڸð༶�Ѿ����ڸÿγ�������Ϣ"
		  end if
		  classCourseInfoRs.Close
		  set classCourseInfoRs = nothing
	  end if
	  '���ݴ�����ϢerrMessage���ݾ����Ƿ�ִ���°༶�γ���Ϣ����Ӳ���
	  if errMessage = "" then
	    sqlString = "select * from [classCourseInfo]"
		  set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  classCourseInfoRs.Open sqlString,conn,1,3
		  classCourseInfoRs.AddNew
		  classCourseInfoRs("courseNumber") = Request("courseNumber")
		  classCourseInfoRs("courseName") = Request("courseName")
		  classCourseInfoRs("courseType") = "���޿�"
		  classCourseInfoRs("classNumber") = Request("classNumber")
		  classCourseInfoRs("termId") = CInt(Request("termId"))
		  classCourseInfoRs("courseScore") = CSng(Request("courseScore"))
		  classCourseInfoRs("courseMemo") = Request("courseMemo")
		  classCourseInfoRs.Update
		  Response.Write "<script>alert('�༶�γ���Ϣ��ӳɹ�!');</script>"
	  else
	    Response.Write "<script>alert('" & errMessage & "');</script>"
	  end if
	end if
%>
<HTML>
<HEAD>
	<Title>�༶�γ���Ϣ���</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�γ���Ϣ����--&gt;�༶�γ���Ϣ���
			 </td>
	   </tr><br>
		<tr>
		    <td width=100 align="right">����ѧ��:</td>
		    <td>
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
		   <td width=100px align="right">���ڰ༶:</td>
			 <td>
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
		    <td width=100 align="right">�γ̱��:</td>
		    <td><input type=text name=courseNumber size=20></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">�γ�����:</td>
		    <td><input type=text name=courseName size=20></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">�γ�ѧ��:</td>
		    <td><input type=text name=courseScore size=5>��</td>
		  </tr>
		  <tr>
		    <td width=100 align="right">������Ϣ:</td>
		    <td><textarea cols=40 rows=5 name=courseMemo></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" ȷ����� ">
				  <input type="reset" value=" ������д ">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>

</HTML>
