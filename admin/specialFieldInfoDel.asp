<!--#include virtual="/Database/conn.asp"-->
<%
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'ȡ��Ҫɾ����רҵ���
	specialFieldNumber = Trim(Request("specialFieldNumber"))
	'����רҵ���Ƿ���ڰ༶��Ϣ,��������ڰ༶��Ҫ��ɾ���༶�����ִ��רҵ��Ϣ��ɾ������
	sqlString = "select * from [classInfo] where classSpecialFieldNumber='" & specialFieldNumber & "'"
	set classInfoRs = Server.CreateObject("ADODB.RecordSet")
	classInfoRs.Open sqlString,conn,1,1
	if not classInfoRs.EOF then
	  Response "<script>alert('��רҵ�»����ڰ༶,����ɾ���༶��Ϣ!');location.href='specialFieldInfoManage.asp';</script>"	
	else  '��������ڰ༶��Ϣ��
	  '��ִ�и�רҵѡ�޿γ̵�ɾ������(ÿ���༶�ı��޿γ��ڰ༶��Ϣɾ������ʱִ��ɾ������)
	  sqlString = "delete from [publicCourseInfo] where specialFieldNumber='" & specialFieldNumber & "'"
	  conn.Execute(sqlString)
	  '��ִ�и�רҵ��Ϣ��ɾ��
	  sqlString = "delete from [specialFieldInfo] where specialFieldNumber='" & specialFieldNumber & "'"
	  conn.Execute(sqlString)
	  Response.Write "<script>alert('רҵ��Ϣɾ���ɹ�!');location.href='specialFieldInfoManage.asp';</script>"
	end if
%>