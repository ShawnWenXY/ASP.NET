<!--#include virtual="/DataBase/conn.asp"-->
<%
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'����µ�ѧ����Ϣ
	if Request("action") = "add" then
	  termBeginYear = CInt(Request("termBeginYear"))
	  termEndYear = CInt(Request("termEndYear"))
	  termUpOrDown = Request("termUpOrDown")
	  if (termEndYear - termBeginYear) <> 1 then
	    Response.Write "<script>alert('�������ѧ�������Ϣ����ȷ!');</script>"
	  else
	    sqlString = "select * from [termInfo]"
		  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
		  termInfoRs.Open sqlString,conn,1,3
		  termInfoRs.AddNew
		  termInfoRs("termBeginYear") = termBeginYear
		  termInfoRs("termEndYear") = termEndYear
		  termInfoRs("termUpOrDown") = termUpOrDown
		  termInfoRs.Update
		  termInfoRs.Close
		  set termInfoRs = Nothing
		  Response.Write "<script>alert('ѧ����Ϣ��ӳɹ�!');</script>"
	  end if
	'ɾ��ĳ��ѧ����Ϣ
	elseif Request("action") = "del" then
	  'ȡ��Ҫɾ����ѧ�ڱ��
	  termId = CInt(Request("termId"))
	  sqlString = "delete from [termInfo] where termId=" & termId
	  conn.Execute(sqlString)
	  Response.Write "<script>alert('ѧ����Ϣɾ���ɹ�!');</script>"
	end if
%>
<HTML>
<HEAD>
	<Title>ѧ����Ϣ����</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce">
	 <table width=400 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/list.gif" width=14px height=14px>ѧ����Ϣ����
			 </td>
	   </tr>
		<tr><td>ѧ����Ϣ</td><td>ɾ��</td></tr>
		<%
		  sqlString = "select * from [termInfo]"
		  set termInfoRs = Server.CreateObject("ADODB.RecordSet")
		  termInfoRs.Open sqlString,conn,1,1
		  while not termInfoRs.EOF
		    Response.Write "<tr><td>" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "��" & termInfoRs("termUpOrDown") & "</td><td><a href='termInfoManage.asp?action=del&termId=" & termInfoRs("termId") & "' onclick=" & """" & "javascript:return confirm('����ɾ���˼�¼��?');" & """" & "><img src='../images/delete.gif' width=12 height=12 border=0>ɾ��</a></td></tr>"
			  termInfoRs.MoveNext
		  wend
		%>
		<tr><td colspan=2>�����ѧ����Ϣ:<input type=text name=termBeginYear size=5>��-<input type=text name=termEndYear size=5>��
		<select name=termUpOrDown>
		  <option value="��ѧ��">��ѧ��</option>
		  <option value="��ѧ��">��ѧ��</option>
		</select><input type=hidden name=action value="add"><input type="submit" value='���'></td></tr>
	</table>
</form>
</body>
</html>