<!--#include virtual="/Database/conn.asp"-->
<%
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  'ȡ��ɾ���İ༶�γ̱��
  courseNumber = Request.QueryString("courseNumber")
  '����ɾ���ÿγ̵ĳɼ���Ϣ
  sqlString = "delete from [scoreInfo] where courseNumber='" & courseNumber & "' and isSelect=0"
  conn.Execute(sqlString)
  'Ȼ��ɾ���ÿγ̵��Ͽ���Ϣ
  sqlString = "delete from [classCourseTeach] where courseNumber='" & courseNumber & "'"
  conn.Execute(sqlString)
  '���ɾ���ð༶�γ̵���Ϣ
  sqlString = "delete from [classCourseInfo] where courseNumber='" & courseNumber & "'"
  conn.Execute(sqlString)
  
  Response.Write "<script>alert('�༶�γ���Ϣɾ���ɹ�!');location.href='classCourseManage.asp';</script>"
%>