<!--#include file="../DataBase/conn.asp"-->
<!--#include file="../System/function.asp"-->
<%
  dim sqlString,courseNumber,termId
  '�����ʦ��û�е�½
  if session("teacherNumber")="" and  session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'ȡ�ò�ѯ�Ŀγ̱�ŵĹؼ���
	courseNumber = Request("courseNumber")
	'ȡ�ò�ѯ��ѧ����Ϣ
	termId = Request("termId")
	'�Ӱ༶���޿��Ͽ���Ϣ���н��в�ѯ��sql
	sqlString = "select * from [classCourseTeach] where teacherNumber='" & session("teacherNumber") & "'"
	if courseNumber <> "" then
	  sqlString = sqlString & " and courseNumber like '%" & courseNumber & "%'"
	end if
	if termId = "" then
	  termId = 0
	end if
	if termId <> 0 then
	  sqlString = sqlString & " and termId=" & termId
	end if
	set classCourseTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	classCourseTeachInfoRs.Open sqlString,conn,1,1
	'��ѡ�޿��Ͽ�ϸ������н��в�ѯ��sql���
	sqlString = "select * from [publicCourseTeach] where teacherNumber='" & Session("teacherNumber") & "'"
	if courseNumber <> "" then
	  sqlString = sqlString & " and courseNumber like '%" & courseNumber & "%'"
	end if
	if termId <> 0 then
	  sqlString = sqlString & " and termId=" & termId
	end if
	set publicCourseTeachInfoRs = Server.CreateObject("ADODB.RecordSet")
	publicCourseTeachInfoRs.Open sqlString,conn,1,1
%>
<html>
<head>
   <title>��ʦ�ڿ���Ϣ��ѯ</title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	<table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#FFFFFF'>
    <form name="form1" method="post" action="teachInfoQuery.asp">
      <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=8 align="center">
		      <img src="../images/list.gif" width=14px height=14px>ѡ����Ϣ����--&gt;�ڿ���Ϣ��ѯ
			 </td>
	   </tr>
     <tr>
	     <td  align="left" height="22" colspan="7" bgcolor="#ffffff"> 
	       �γ̱��:<input type="text" name=courseNumber size=18 value='<%=courseNumber%>'>&nbsp;
	       ����ѧ��:
	       <select name=termId>
	        <option value="">ѡ��ѧ��</option>
		    <%
		    sqlString = "select * from [termInfo]"
			  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
			  termInfoRs.Open sqlString,conn,1,1
			  while not termInfoRs.EOF
			    selected = ""
				  if termId = "" then
				    termId = 0
					else
					  termId = CInt(termId)
					end if
				  if termInfoRs("termId") = termId then
				    selected = "selected"
					end if
			    Response.Write "<option value='" & termInfoRs("termId") & "' " & selected & ">" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "��" & termInfoRs("termUpOrDown") & "</option>"
					termInfoRs.MoveNext
			  wend
		  %>
	      </select>
	      <input type="submit" value=" ���� " class="button1">
     </td>
    </tr>
	  <%
	    if classCourseTeachInfoRs.EOF and publicCourseTeachInfoRs.EOF then
		    Response.Write "<tr><td colspan=6 align=center>�Բ���,��û�ж�Ӧ���ڿ���Ϣ</td></tr>"
		  else
		    Response.Write "<tr><td>�γ̱��</td><td>����ѧ��</td><td>���ڰ༶(��רҵ)</td><td>�γ�����</td><td>�Ͽν���</td><td>�Ͽ�ʱ��</td><td>��ϸ</td></tr>"
		  end if
		  while not classCourseTeachInfoRs.EOF
		    Response.Write "<tr><td>" & classCourseTeachInfoRs("courseNumber") & "</td><td>" & GetTermnameById(classCourseTeachInfoRs("termId")) & "</td>"
			  Response.Write "<td>" & GetClassNameByNumber(classCourseTeachInfoRs("classNumber")) & "</td><td>���޿�</td><td>" & classCourseTeachInfoRs("teachClassRoom") & "</td>"
			  Response.Write "<td>����" & classCourseTeachInfoRs("teachDay") & "</td><td><a href='classCourseTeachDetail.asp?termId=" & termId & "&courseNumber=" & courseNumber & "&teachId=" & classCourseTeachInfoRs("teachId") & "'>��ϸ</a></td></tr>"
		    classCourseTeachInfoRs.MoveNext
		  wend
		  while not publicCourseTeachInfoRs.EOF
		    Response.Write "<tr><td>" & publicCourseTeachInfoRs("courseNumber") & "</td><td>" & GetTermnameById(publicCourseTeachInfoRs("termId")) & "</td>"
			  Response.Write "<td>" & GetSpecialFieldNameByNumber(publicCourseTeachInfoRs("specialFieldNumber")) & "</td><td>ѡ�޿�</td><td>" & publicCourseTeachInfoRs("teachClassRoom") & "</td>"
			  Response.Write "<td>����" & publicCourseTeachInfoRs("teachDay") & "</td><td><a href='publicCourseTeachDetail.asp?termId=" & termId & "&courseNumber=" & courseNumber & "&teachId=" & publicCourseTeachInfoRs("teachId") & "'>��ϸ</a></td></tr>"
		    publicCourseTeachInfoRs.MoveNext
		  wend
	  %>
    </form>
	</table>
</body>
</html>