<!--#include virtual="/Database/conn.asp"-->
<%
  dim teacherNumber,sqlString
  teacherNumber = Request("teacherNumber")
  '����ȡ�øý�ʦ���̵����б��޿γ�
  sqlString = "select courseNumber from [classCourseTeach] where teacherNumber='" & teacherNumber & "'"
  set teacherTeachCourseRs = Server.CreateObject("ADODB.RecordSet")
  teacherTeachCourseRs.Open sqlString,conn,1,1
  while not teacherTeachCourseRs.EOF
    'ɾ���ñ���γ̵���Ϣ��¼
    sqlString = "delete from [classCourseInfo] where courseNumber='" & teacherTeachCourseRs("courseNumber") & "'"
	  conn.Execute(sqlString)
	  'ɾ���ñ��޿γ̵ĳɼ���Ϣ
	  sqlString = "delete from [scoreInfo] where courseNumber='" & teacherTeachCourseRs("courseNumber") & "' and isSelect=0"
	  conn.Execute(sqlString)
	  teacherTeachCourseRs.MoveNext
  wend 
  'ɾ���ñ��޿γ̵��ڿ���Ϣ
  sqlString = "delete from [classCourseTeach] where teacherNumber='" & teacherNumber & "'"
  conn.Execute(sqlString)
  teacherTeachCourseRs.Close
  'Ȼ��õ��ý�ʦ���̵�ѡ�޿γ�
  sqlString = "select courseNumber from [publicCourseTeach] where teacherNumber='" & teacherNumber & "'"
  'set teacherTeachCourseRs = Server.CreateObject("ADODB.RecordSet")
  teacherTeachCourseRs.Open sqlString,conn,1,1
  'set teacherTeachCourseRs = conn.Execute(sqlString)
  while not teacherTeachCourseRs.EOF
    'ɾ����ѡ�޿γ̵���Ϣ��¼
	  sqlString = "delete from [publicCourseInfo] where courseNumber='" & teacherTeachCourseRs("courseNumber") & "'"
	  conn.Execute(sqlString)
	  'ɾ����ѡ�޿γ̵ĳɼ���Ϣ
	  sqlString = "delete from [scoreInfo] where courseNumber='" & teacherTeachCourseRs("courseNumber") & "' and isSelect=1"
	  conn.Execute(sqlString)
	  'ɾ����ѡ�޿ε�ѧ��ѡ����Ϣ
	  sqlString = "delete from [studentSelectCourseInfo] where courseNumber='" & teacherTeachCourseRs("courseNumber") & "'"
	  conn.Execute(sqlString)
	  teacherTeachCourseRs.MoveNext
  wend
  'ɾ����ѡ�޿ε��ڿ���Ϣ
  sqlString = "delete from [publicCourseTeach] where teacherNumber='" & teacherNumber & "'"
  conn.Execute(sqlString)
  teacherTeachCourseRs.Close
  '���ɾ����ʦ�ĸ�����Ϣ��¼
  sqlString = "delete from [teacherInfo] where teacherNumber='" & teacherNumber & "'"
  conn.Execute(sqlString)
  
  Response.Write "<script>alert('��ʦ��Ϣɾ���ɹ�!');location.href='teacherInfoManage.asp';</script>"
%>
