<!--#include file="../database/conn.asp"-->
<!--#include file="../system/function.asp"-->
<%
  '���ѧ����û�е�½
  if session("studentnumber")="" then
    response.write "<script>top.location.href='../login.asp';</script>"
	end if
	dim classname(12),classfieldname(12)
	classname(1) = "�����һ��"
	classname(2) = "����ڶ���"
	classname(3) = "���������"
	classname(4) = "������Ľ�"
	classname(5) = "��������"
	classname(6) = "�����һ��"
	classname(7) = "����ڶ���"
	classname(8) = "���������"
	classname(9) = "������Ľ�"
	classname(10) = "���ϵ�һ��"
	classname(11) = "���ϵڶ���"
	classname(12) = "���ϵ�����"
	classfieldname(1) = "morningone"
	classfieldname(2) = "morningtwo"
	classfieldname(3) = "morningthree"
	classfieldname(4) = "morningfour"
	classfieldname(5) = "morningfive"
	classfieldname(6) = "afternoonone"
	classfieldname(7) = "afternoontwo"
	classfieldname(8) = "afternoonthree"
	classfieldname(9) = "afternoonfour"
	classfieldname(10) = "eveningone"
	classfieldname(11) = "eveningtwo"
	classfieldname(12) = "eveningthree"
%>
<html>
<head>
	<title>ѧ��ѡ��</title>
	<meta http-equiv="content-type" content="text/html; charset=gb2312">
	<link href="../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
	<table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#ffffff'>
    <form name="form1" method="post" action="coursetablemake.asp">
      <tr style="color:blue;font-size:14px;">
	      <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/list.gif" width=14px height=14px>ѡ����Ϣ����--&gt;���ɿα�
			 </td>
	   </tr>
		 <tr>
		   <td>��ѡ��ѧ��:</td>
			 <td>
			   <select name=termid>
				    <option value="">��ѡ��</option>
					  <%
					    dim sqlstring
						  sqlstring = "select * from [terminfo]"
						  set terminfors = server.createobject("adodb.recordset")
						  terminfors.open sqlstring,conn,1,1
						  while not terminfors.eof
						    response.write "<option value='" & terminfors("termid") & "'>" & terminfors("termbeginyear") & "-" & terminfors("termendyear") & "��" & terminfors("termupordown") & "</option>"
							  terminfors.movenext
						  wend
					  %>
				  </select>&nbsp;<input type="submit" name="submit" value="���ɿα�">
				</td>
		 </tr>
	  <tr>
	    <td height="30">�α���:</td>
			<td>
				<table width=100% border=1 cellspacing=0 bordercolor="green">
				  <%
						if request("termid") <> "" then
						   response.write "<tr><td colspan=8 align=center>" & gettermnamebyid(request("termid")) & " �α�</td></tr>"
						end if
				  %>
				  <tr><td width="12.5%">�α�</td><td width="12.5%">����һ</td><td width="12.5%">���ڶ�</td><td width="12.5%">������</td><td width="12.5%">������</td><td width="12.5%">������</td><td width="12.5%">������</td><td width="12.5%">������</td></tr>
			  <%
			      for i = 1 to 12 
						   response.write "<tr><td>" & classname(i) & "</td>"
					     for j = 1 to 7
						      response.write "<td align=center id='class" & i & j & "'>&nbsp;</td>"
						   next
							 response.write "</tr>"
						next
			  %>
			  </table>
			</td>
	 </tr>
	</form>
 </table>
 <script language="javascript">
 <%
   '���ѡ����Ҫ����ĳ��ѧ�ڵĿα�
   if request("termid") <> "" then
	   '�õ���ѧ���İ༶���
	   classnumber = getclassnumberbystudentnumber(session("studentnumber"))
		 '�õ���ѧ����רҵ���
		 specialfieldnumber = getspecialfieldnumberbystudentnumber(session("studentnumber"))
		 '��ѯ��ѧ�ڸð༶�ı��޿γ��Ͽ���Ϣ
		 sqlstring = "select * from [classcourseteachview] where classnumber='" & classnumber & "' and termid=" & request("termid")
		 set classcourseteachrs = server.createobject("adodb.recordset")
		 classcourseteachrs.open sqlstring,conn,1,1
		 while not classcourseteachrs.eof
		   for i = 1 to 12
			   if classcourseteachrs(classfieldname(i)) = true then
				   '���ĳ�ڿδ����Ͽ���Ϣ���������Ӧ��λ��
				   response.write "document.all.class" & i & classcourseteachrs("teachday") & ".innerhtml=" & """" & classcourseteachrs("coursename") & "(" & classcourseteachrs("teachclassroom") & ")" & """" & ";" & vbcrlf
				 end if
			 next
			 classcourseteachrs.movenext
		 wend
		 classcourseteachrs.close
		 '��ѯ��ѧ�ڸ�ѧ��ѡ�޿γ��Ͽ���Ϣ
		 sqlstring = "select * from [publiccourseteachview] where studentnumber='" & session("studentnumber") & "' and termid=" & request("termid")
		 set publiccourseteachrs = server.createobject("adodb.recordset")
		 publiccourseteachrs.open sqlstring,conn,1,1
		 while not publiccourseteachrs.eof
			 for i = 1 to 12
			   if publiccourseteachrs(classfieldname(i)) = true then
				   '���ĳ�ڿδ����Ͽ���Ϣ���������Ӧ��λ��
				   response.write "document.all.class" & i & publiccourseteachrs("teachday") & ".innerhtml+=" & """" & publiccourseteachrs("coursename") & "(" & publiccourseteachrs("teachclassroom") & ")" & """" & ";" & vbcrlf
				 end if
			 next
			 publiccourseteachrs.movenext
		 wend
	end if
 %>
 </script>
 </body>
 </html>