<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/config.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  dim sqlString
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  '��ѯ�༶�γ̵�sql���
  sqlString = "select * from [publicCourseInfo] where 1=1"
  '�жϲ�ѯ�Ŀγ̱���Ƿ�Ϊ�����Ʋ�ѯsql���
  if Request("courseNumber") <> "" then
    sqlString = sqlString & " and courseNumber like '%" & Trim(Request("courseNumber")) & "%'"
  end if
  '�жϲ�ѯ�Ŀγ������Ƿ�Ϊ�����Ʋ�ѯsql���
  if Request("courseName") <> "" then
    sqlString = sqlString & " and courseName like '%" & Trim(Request("courseName")) & "%'"
  end if
  '�жϲ�ѯ��רҵ��Ϣ�Ƿ�������Ʋ�ѯ��sql���
  if Request("specialFieldNumber") <> "" then
    sqlString = sqlString & " and specialFieldNumber='" & Request("specialFieldNumber") & "'"
  end if
  '�жϲ�ѯ��ѧ����Ϣ�Ƿ�������Ʋ�ѯ��sql���
  if Request("termId") <> "" then
    sqlString = sqlString & " and termId=" & CInt(Request("termId"))
  end if
  set publicCourseInfoRs = Server.CreateObject("ADODB.RecordSet")
  publicCourseInfoRs.Open sqlString,conn,1,3
  If Request("Page") = "" Then
		intPage = 1
	Else
		intPage = Clng(Request("Page"))
	End If
	'����ÿҳ��ʾ�ļ�¼��
	publicCourseInfoRs.PageSize = pageSize
	If intPage > publicCourseInfoRs.PageCount Then
		intPage = publicCourseInfoRs.PageCount
	End If
	If intPage <= 0 Then
		intPage = 1
	End If
	If Not publicCourseInfoRs.EOF Then
		publicCourseInfoRs.AbsolutePage = intPage
	End If
%>
<HTML>
<HEAD>
	<Title>ѡ�޿γ���Ϣ����</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript">
	function formsubmit(page)
	{
		str=document.form1;
		str.page.value=page;
		str.submit();
	}
	function changepage()
	{
		str=document.form1;
		str.page.value=str.selectpage.value;
		str.submit();
	}
	</script>
</HEAD>
<BODY>
	<table width=700px border="1" align="center" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#FFFFFF'>
 <form name="form1" method="post" action="publicCourseInfoManage.asp">
 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=7 align="center">
		      <img src="../images/list.gif" width=14px height=14px>�γ���Ϣ����--&gt;ѡ�޿γ���Ϣ�б�
			 </td>
	   </tr><br>
 <tr>
	<td  align="left" height="22" colspan="7" bgcolor="#ffffff"> 
	���γ̱��:<input type=text name=courseNumber size=8>&nbsp;
	  �γ�����:<input type=text name=courseName size=8>&nbsp;
	  רҵ:<select name=specialFieldNumber>
	             <option value="">��ѡ��</option>
					   <%
						   sqlString = "select specialFieldNumber,specialFieldName from [specialFieldInfo]"
							 set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
							 specialFieldInfoRs.Open sqlString,conn,1,1
							 while not specialFieldInfoRs.EOF
							   Response.Write "<option value='" & specialFieldInfoRs("specialFieldNumber") & "'>" & specialFieldInfoRs("specialFieldName") & "</option>"
								 specialFieldInfoRs.MoveNext
							 wend
						 %>
						 </select>
		ѧ��:<select name=termId>
		       <option value="">��ѡ��</option>
				   <%
					   sqlString = "select * from [termInfo]"
						 set termInfoRs = Server.CreateObject("ADODB.RecordSet")
						 termInfoRs.Open sqlString,conn,1,1
						  while not termInfoRs.EOF
						    Response.Write "<option value='" & termInfoRs("termId") & "'>" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "��" & termInfoRs("termUpOrDown") & "</option>"
							  termInfoRs.MoveNext
						  wend
						  termInfoRs.Close
					 %>
			   </select>
		<input type="submit" value=" ���� " class="button1">
	<td>
 </tr>
  <tr> 
		<td  nowrap> 
      <div align="center">�γ̱��</div>
    </td>
	 <td>
	   <div align="center">�γ�����</div>
		</td>
		<td>
	   <div align="center">����רҵ</div>
		</td>
		<td>
	   <div align="center">����ѧ��</div>
		</td>
		<td>
	   <div align="center">�γ�ѧ��</div>
		</td>
		<td>
		  <div align="center">����</div>
		</td>
  </tr>
  <%
    for i=0 to publicCourseInfoRs.PageSize-1
	    if not publicCourseInfoRs.EOF then
  %>
  <tr align="center"> 
    <td nowrap>&nbsp;<%=publicCourseInfoRs("courseNumber")%></td>
    <td nowrap>&nbsp;<%=publicCourseInfoRs("courseName")%></td>
    <td nowrap>&nbsp;<%=GetSpecialFieldNameByNumber(publicCourseInfoRs("specialFieldNumber"))%></td>
	  <td nowrap>&nbsp;<%=GetTermnameById(publicCourseInfoRs("termId"))%></td>
	  <td nowrap>&nbsp;<%=publicCourseInfoRs("courseScore")%></td>
   <td nowrap>&nbsp;<a href="publicCourseInfoEdit.asp?courseNumber=<%=publicCourseInfoRs("courseNumber")%>"><img src="../images/edit.gif" border=0 height=12 width=12>�༭</a>&nbsp;&nbsp;<a href="publicCourseInfoDel.asp?courseNumber=<%=publicCourseInfoRs("courseNumber")%>" onClick="javascript:return confirm('��ľ���ɾ���˼�¼��?');"><img src="../images/delete.gif" border=0 height=12 width=12>ɾ��</a></td>
  </tr>
  <%
        publicCourseInfoRs.MoveNext
		  End If
	  Next
	%>
  <tr> 
		      <td  align="right" height="22" colspan="7" bgcolor="#ffffff"> 
		        <%
		        If intPage = 1 Or intPage = 0 Then
					Response.Write "ǰһҳ"
		        Else
		        %>
		        <a href="javascript:formsubmit('<%=intPage - 1%>');">ǰһҳ</a> 
		        <%
		        End If
		        %>
		        &nbsp;&nbsp; 
		        <%
		        If intPage = publicCourseInfoRs.PageCount or publicCourseInfoRs.PageCount=0 Then
					Response.Write "��һҳ"
		        Else
		        %>
		        <a href="javascript:formsubmit('<%=intPage + 1%>');">��һҳ</a> 
		        <%
		        End If
		        %>
		        &nbsp; ת�� 
		        <select name="selectpage" onchange="changepage();">
		          <%
					If publicCourseInfoRs.PageCount <= 0 Then
						Response.Write "<option value=''>��</option>"
					Else
						For intLoop = 1 To publicCourseInfoRs.PageCount
							If intPage = intLoop Then
								Response.Write "<option value='" & intLoop & "' selected>" & intLoop & "</option>"
							Else
								Response.Write "<option value='" & intLoop & "'>" & intLoop & "</option>"
							End If
						Next
					End If
					%>
		        </select>ҳ
		      </td>
		    </tr>
	    <input type="hidden" name="page" value="">
    </form>
</table>
</BODY>
</HTML>
