<!--#include file="../DataBase/conn.asp"-->
<!--#include file="../System/md5.asp"-->
<%
  '���������Ϣ,���û�д�����Ϊ��
  dim errMessage
  errMessage = ""
  '���ѧ����û�е�½
  if session("studentNumber")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'���Ҫ���޸�Ϊ�µ�����
	if Request("submit") <> "" then
	  oldPassword = Trim(Request("oldPassword"))
	  newPassword = Trim(Request("newPassword"))
	  newPasswordAgain = Trim(Request("newPasswordAgain"))
	  if oldPassword = "" then
	    errMessage = errMessage & "������ԭ��������!"
		end if
		if newPassword = "" then
		  errMessage = errMessage & "�������µ�����"
	  end if
	  if newPassword <> newPasswordAgain then
	    errMessage = errMessage & "��������ĵ������벻һ��"
		end if
		if errMessage = "" then
		  sqlString = "select * from [studentInfo] where studentNumber='" & session("studentNumber") & "'"
		  set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
		  studentInfoRs.Open sqlString,conn,1,3
		  '�ж������ԭ�������Ƿ���ȷ
		  if studentInfoRs("studentPassword") <> md5(oldPassword) then
		    Response.Write "<script>alert('������ľ����벻��ȷ!');</script>"
			else
			  studentInfoRs("studentPassword") = md5(newPassword)
			  studentInfoRs.Update
			  Response.Write "<script>alert('�����޸ĳɹ�!');</script>"
			end if
			studentInfoRs.Close
		else
		  Response.Write "<script>alert('" & errMessage & "');</script>"
		end if
	end if
%>
<HTML>
<HEAD>
	<Title>��½��������</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>ϵͳ����--&gt;��������
			 </td>
	   </tr>
		 <tr>
		   <td width=100 align="right">ԭ������:</td>
			 <td><input type=password name=oldPassword size=20></td>
		</tr>
		<tr>
		  <td width=100 align="right">������:</td>
		  <td><input type=password name=newPassword size=20></td>
		</tr>
		<tr>
		  <td width=100 align="right">������ȷ��:</td>
		  <td><input type=password name=newPasswordAgain></td>
		</tr>
		<tr>
		  <td colspan=2 align="center">
		    <input type="submit" name=submit value="�޸�����">&nbsp;&nbsp;
			  <input type="reset" value="ȡ��">
		  </td>
		</tr>
	</table>
</form>
</body>
</html>
