<!--#include file="Database/conn.asp"-->
<!--#include file="System/md5.asp"-->
<%
	'ȡ���û���,����,�����Ϣ
  dim username,password,identity,sqlString
  username = Trim(CStr(Request("username")))
  password = Trim(CStr(Request("password")))
  identity = Trim(CStr(Request("identity")))
  '�������½��Ϣ��صĽ����¼��
  set rsLogin = Server.CreateObject("adodb.recordset")
  '�����ѧ�����
  if identity = "student" then
    sqlString = "select * from [studentInfo] where studentNumber='" & username & "' and studentPassword='" & md5(password) & "'"
	  rsLogin.Open sqlString,conn,1,1
	  '���ѧ��ѧ�ź����붼������ȷ
	  if not rsLogin.EOF then
	    Session("studentNumber") = username
		  Response.Redirect "Student/index.asp"
		  Response.End
		'�����ѧ�ź�����ļ�¼������
	  else
	    Response.Write "<script>alert('ѧ�Ż������������!');location.href='login.asp';</script>"
		  Response.End
		end if
	'����ǽ�ʦ���
	elseif identity = "teacher" then
	  sqlString = "select * from [teacherInfo] where teacherNumber='" & username & "' and teacherPassword='" & md5(password) & "'"
	  rsLogin.Open sqlString,conn,1,1
	  '�����ʦ�̹��ź����붼������ȷ
	  if not rsLogin.EOF then
	    Session("teacherNumber") = username
		  Response.Redirect "Teacher/index.asp"
		'��������ڸý̹��ŵĽ�ʦ��Ϣ
		else
		  Response.Write "<script>alert('��ʦ��½�ʺŻ��������!');location.href='login.asp';</script>"
		  Response.End
		end if
	'����ǹ���Ա��ݵ�½
	else
	  sqlString = "select * from [admin] where adminUsername='" & username & "' and adminPassword='" & password & "'"
	  rsLogin.Open sqlString,conn,1,1
	  '�������Ա�ʺź����붼������ȷ
	  if not rsLogin.EOF then
	    Session("adminUsername") = username
		  Response.Redirect "Admin/index.asp"
		'�������Ա����Ϣ������
		else
		  Response.Write "<script>alert('�ù���Ա����Ϣ������!');location.href='login.asp';</script>"
		end if
	end if
%>
