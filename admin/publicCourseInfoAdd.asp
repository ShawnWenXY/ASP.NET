<!--#include virtual="/Database/conn.asp"-->
<%
  'errMessage���������Ϣ
  dim errMessage,sqlString
  errMessage = ""
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'�������Ա������µĹ�ѡ�γ���Ϣ���ύ
	if Request("submit") <> "" then
	  '���û��ѡ�����ڵ�ѧ��
	  if Request("termId") = "" then
	    errMessage = "��ѡ�����ڵ�ѧ��!"
	  end if
	  '���û��ѡ��רҵ��Ϣ
	  if Request("specialFieldNumber") = "" then
	    errMessage = "��ѡ��רҵ!"
	  end if
	  '���û������γ̱��
	  if Request("courseNumber") = "" then
	    errMessage = "������γ̱��!"
	  end if
	  '���û������γ�����
	  if Request("courseName") = "" then
	    errMessage = "������γ�����!"
	  '����ѧ�ڸ�רҵ�ÿγ���Ϣ�Ƿ��Ѿ�����
	  else
	    sqlString = "select * from [publicCourseInfo] where specialFieldNumber='" & Request("specialFieldNumber") & "' and termId=" & Request("termId") & " and courseName='" & Request("courseName") & "'"
		  set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  publicCourseInfoRs.Open sqlString,conn,1,1
		  if not publicCourseInfoRs.EOF then
		    errMessage = "��ѧ�ڸð༶�Ѿ����ڸÿγ�������Ϣ"
		  end if
		  publicCourseInfoRs.Close
		  set publicCourseInfoRs = nothing
	  end if
	  '���ݴ�����ϢerrMessage���ݾ����Ƿ�ִ���¹�ѡ�γ���Ϣ����Ӳ���
	  if errMessage = "" then
	    sqlString = "select * from [publicCourseInfo]"
		  set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  publicCourseInfoRs.Open sqlString,conn,1,3
		  publicCourseInfoRs.AddNew
		  publicCourseInfoRs("courseNumber") = Request("courseNumber")
		  publicCourseInfoRs("courseName") = Request("courseName")
		  publicCourseInfoRs("courseType") = "ѡ�޿�"
		  publicCourseInfoRs("specialFieldNumber") = Request("specialFieldNumber")
		  publicCourseInfoRs("termId") = CInt(Request("termId"))
		  publicCourseInfoRs("courseScore") = CSng(Request("courseScore"))
		  publicCourseInfoRs("courseMemo") = Request("courseMemo")
		  publicCourseInfoRs.Update
		  Response.Write "<script>alert('ѡ�޿γ���Ϣ��ӳɹ�!');</script>"
	  else
	    Response.Write "<script>alert('" & errMessage & "');</script>"
	  end if
	end if
%>
<HTML>
<HEAD>
	<Title>רҵ��ѡ�γ���Ϣ���</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>�γ���Ϣ����--&gt;רҵ��ѡ��Ϣ���
			 </td>
	   </tr><br>
		<tr>
		    <td width=100 align="right">����ѧ��:</td>
		    <td>
			    <select name=termId>
				    <option value="">��ѡ��</option>
					  <%
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
		   <td width=100px align="right">ѡ��רҵ:</td>
			 <td>
			   <select name=specialFieldNumber>
				   <option value="">��ѡ��</option>
				  <%
					  sqlString = "select * from [specialFieldInfo]"
					  set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
					  specialFieldInfoRs.Open sqlString,conn,1,1
					  while not specialFieldInfoRs.Eof
					    Response.Write "<option value='" & specialFieldInfoRs("specialFieldNumber") & "'>" & specialFieldInfoRs("specialFieldName") & "</option>"
						  specialFieldInfoRs.MoveNext
					  wend
					  specialFieldInfoRs.Close
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
