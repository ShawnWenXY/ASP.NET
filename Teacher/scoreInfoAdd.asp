<!--#include file="../DataBase/conn.asp"-->
<%
  dim studentNumber,courseNumber,courseType,score,isSelect,errMessage,sqlString
  errMessage = ""
  '�����ʦ��û�е�½
  if session("teacherNumber")="" and  session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'���Ҫ������µĳɼ���Ϣ
	if Request("submit") <> "" then
	  studentNumber = Trim(Request("studentNumber"))
	  courseNumber = Trim(Request("courseNumber"))
	  courseType = Request("courseType")
	  score = Request("score")
	  if score = "" then
	    score = 0 
		else
	    score = CSng(score)
		end if
	  if studentNumber = "" then
	    errMessage = errMessage & "ѧ�Ų���Ϊ��!"
		end if
		if courseNumber = "" then
		  errMessage = errMessage & "�γ̱�Ų���Ϊ��!"
		end if
		'��ѯ�Ƿ��и�ѧ�ŵ�ѧ������Ϣ����
		sqlString = "select * from [studentInfo] where studentNumber='" & studentNumber & "'"
		set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
		studentInfoRs.Open sqlString,conn,1,1
		if studentInfoRs.EOF then
		  errMessage = errMessage & "�����ڸ�ѧ�ŵ�ѧ����Ϣ!"
		end if
		studentInfoRs.Close
		'��ѯ�Ƿ��иÿγ̱�ŵĿγ���Ϣ
		if Request("courseType") = "���޿�" then
		  sqlString = "select * from [classCourseInfo] where courseNumber='" & courseNumber & "'"
		  set classCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  classCourseInfoRs.Open sqlString,conn,1,1
		  if classCourseInfoRs.EOF then
		    errMessage = errMessage & "�����ڸÿγ̵���Ϣ!"
		  end if
		  classCourseInfoRs.Close
 		else
		  sqlString = "select * [publicCourseInfo] where courseNumber='" & courseNumber & "'"
		  set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  publicCourseInfoRs.Open sqlString,conn,1,1
		  if publicCourseInfoRs.EOF then
		    errMessage = errMessage & "�����ڸÿγ̵���Ϣ!"
			end if
			publicCourseInfoRs.Close
		end if
		'��ѯ��ѧ���Ƿ�ѡ���˸��ſγ�
		if Request("courseType") = "���޿�" then
		  sqlString = "select * from [classCourseInfo],[studentInfo] where [classCourseInfo].classNumber = [studentInfo].studentClassNumber and courseNumber='" & courseNumber & "' and studentNumber='" & studentNumber & "'"
		  set classCourseSelectInfoRs = Server.CreateObject("ADODB.RecordSet")
		  classCourseSelectInfoRs.Open sqlString,conn,1,1
		  if classCourseSelectInfoRs.EOF then
		    errMessage = errMessage & "��ѧ��û��ѡ�޸��ſγ�!"
			end if
			classCourseSelectInfoRs.Close
		else
		  sqlString = "select * from [studentSelectCourseInfo] where studentNumber='" & studentNumber & "' and courseNumber='" & courseNumber & "'"
		  set publicCourseSelectInfoRs = Server.CreateObject("ADODB.RecordSet")
		  publicCourseSelectInfoRs.Open sqlString,conn,1,1
		  if publicCourseSelectInfoRs.EOF then
		    errMessage = errMessage & "��ѧ��û��ѡ�޸��ſγ�!"
		  end if
		  publicCourseSelectInfoRs.Close
		end if
		'��ѯ��ʦ�Ƿ�ѡ������˸��ſγ�
		if Request("courseType") = "���޿�" then
		  sqlString = "select * from [classCourseTeach] where courseNumber='" & courseNumber & "' and teacherNumber='" & Session("teacherNumber") & "'"
		  set classCourseTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
		  classCourseTeachInfoRs.Open sqlString,conn,1,1
		  if classCourseTeachInfoRs.EOF then
		    errMessage = errMessage & "�Բ�����û�н��ڸ��ſγ�!"
			end if
			classCourseTeachInfoRs.Close
		else
		  sqlString = "select * from [publicCourseTeach] where courseNumber='" & courseNumber & "' and teacherNumber='" & Session("teacherNumber") & "'"
		  set publicCourseTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
		  publicCourseTeachInfoRs.Open sqlString,conn,1,1
		  if publicCourseTeachInfoRs.EOF then
		    errMessage = errMessage & "�Բ���,��û�н��ڸ��ſγ�"
		  end if
		  publicCourseTeachInfoRs.Close
		end if
		
		if courseType = "���޿�" then
		  isSelect = 0
		else
		  isSelect = 1
		end if
		
		'����errMessage�������ж��Ƿ���гɼ�����Ϣ
		if errMessage = "" then
		  '����ѧ�Ÿ��ſγ̵ĳɼ��Ƿ��Ѿ������
		  sqlString = "select * from [scoreInfo] where studentNumber='" & studentNumber & "' and courseNumber='" & courseNumber & "' and isSelect=" & isSelect
		  set scoreInfoRs = Server.CreateObject("ADODB.RecordSet")
		  scoreInfoRs.Open sqlString,conn,1,1
		  '����Ѿ�����˸��ſγ̾��޸ĳɼ�
		  if not scoreInfoRs.EOF then
		    sqlString = "update [scoreInfo] set score=" & score & " where scoreId=" & scoreInfoRs("scoreId")
			  conn.Execute(sqlString)
			  Response.Write("<script>alert('�ɼ���Ϣ�޸ĳɹ�!');</script>")
			  scoreInfoRs.Close
		  else
		    scoreInfoRs.Close
		    sqlString = "select * from [scoreInfo]"
		    set scoreInfoRs = Server.CreateObject("ADODB.RecordSet")
		    scoreInfoRs.Open sqlString,conn,1,3
		    scoreInfoRs.AddNew
		    scoreInfoRs("studentNumber") = studentNumber
		    scoreInfoRs("courseNumber") = courseNumber
		    scoreInfoRs("isSelect") = isSelect
		    scoreInfoRs("score") = score
		    scoreInfoRs.Update
		    Response.Write "<script>alert('�ɼ���Ϣ��ӳɹ�!');</script>"
			end if
		else
		  Response.Write "<script>alert('" & errMessage & "');</script>"
		end if
	end if
%>
<HTML>
<HEAD>
	<Title>ѧ���ɼ����</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	<table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#FFFFFF'>
    <form name="form1" method="post" action="scoreInfoAdd.asp">
      <tr style="color:blue;font-size:14px;">
	      <td style="height: 14px" colspan=8 align="center">
		      <img src="../images/list.gif" width=14px height=14px>�ɼ���Ϣ����--&gt;���/�޸�ѧ���ɼ�
			 </td> 
	   </tr>
		 <tr>
			 <td width=100 align="right">ѧ��:</td>
			 <td><input type="text" name=studentNumber size=20></td>
		 </tr>
		 <tr>
		   <td width=100px align="right">�γ̱��:</td>
			 <td><input type=text name=courseNumber size=20></td>
		 </tr>
		 <tr>
		   <td width=100px align="right">�γ�����:</td>
			 <td>
			   <select name=courseType>
				   <option value="���޿�">���޿�</option>
					 <option value="ѡ�޿�">ѡ�޿�</option>
				 </select>
			 </td>
		 </tr>
     <tr>
		  <td width=100px align="right">�ɼ�:</td>
		  <td><input type="text" name=score size=6>��</td>
	  </tr>
	  <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" ȷ�ϸ��� ">
				  <input type="reset" value=" ������д ">
		    </td>
	 </tr>
	</form>
 </table>
</body>
</html>