<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/config.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  sqlString = "select * from [classInfo] where 1=1"
  'ȡ�ò�ѯ�İ༶��Źؼ�����Ϣ
  classNumber = Trim(Request("classNumber"))
  if classNumber <> "" then
    sqlString = sqlString & " and classNumber like '%" & classNumber & "%'"
  end if
  'ȡ�ò�ѯ�İ༶���ƵĹؼ�����Ϣ
  className = Trim(Request("className"))
  if className <> "" then
    sqlString = sqlString & " and className like '%" & className & "%'"
  end if
  'ȡ�ò�ѯ��רҵ�����Ϣ
  classSpecialFieldNumber = Request("classSpecialFieldNumber")
  if classSpecialFieldNumber <> "" then
    sqlString = sqlString & " and classSpecialFieldNumber='" & classSpecialFieldNumber & "'"
  end if
  set classInfoRs = Server.CreateObject("ADODB.RecordSet")
  classInfoRs.Open sqlString,conn,1,3
  If Request("Page") = "" Then
		intPage = 1
	Else
		intPage = Clng(Request("Page"))
	End If
	'����ÿҳ��ʾ�ļ�¼��
	classInfoRs.PageSize = pageSize
	If intPage > classInfoRs.PageCount Then
		intPage = classInfoRs.PageCount
	End If
	If intPage <= 0 Then
		intPage = 1
	End If
	If Not classInfoRs.EOF Then
		classInfoRs.AbsolutePage = intPage
	End If
%>
<HTML>
<HEAD>
	<Title>�༶��Ϣ����</Title>
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
 <form name="form1" method="post" action="classInfoManage.asp">
 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=8 align="center">
		      <img src="../images/list.gif" width=14px height=14px>�༶��Ϣ����--&gt;�༶��Ϣ�б�
			 </td>
	   </tr><br>
 <tr>
	<td  align="left" height="22" colspan="8" bgcolor="#ffffff"> 
	  �༶���:&nbsp;<input type=text name=classNumber size=15>&nbsp;&nbsp;
	  �༶����:&nbsp;<input type=text name=className size=10>&nbsp;&nbsp;
		�༶����רҵ:<select name=classSpecialFieldNumber>
		       <option value="">��ѡ��</option>
				   <%
					   sqlString = "select * from [specialFieldInfo]"
						 set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
						 specialFieldInfoRs.Open sqlString,conn,1,1
						 while not specialFieldInfoRs.EOF
						   Response.Write "<option value='" & specialFieldInfoRs("specialFieldNumber") & "'>" & specialFieldInfoRs("specialFieldName") & "</option>"
							 specialFieldInfoRs.MoveNext
						 wend
					 %>
			   </select>
		<input type="submit" value=" ���� " class="button1">
	<td>
 </tr>
  <tr> 
		<td  nowrap> 
      <div align="center">�༶���</div>
    </td>
	 <td>
	   <div align="center">�༶����</div>
		</td>
		<td>
	   <div align="center">����רҵ</div>
		</td>
		<td>
	   <div align="center">����ʱ��</div>
		</td>
		<td>
	   <div align="center">����������</div>
		</td>
		<td>
	   <div align="center">ѧ��</div>
		</td>
		<td>
		  <div align="center">����</div>
		</td>
  </tr>
  <%
    for i=0 to classInfoRs.PageSize-1
	    if not classInfoRs.EOF then
  %>
  <tr align="center"> 
    <td nowrap>&nbsp;<%=classInfoRs("classNumber")%></td>
    <td nowrap>&nbsp;<%=classInfoRs("className")%></td>
    <td nowrap>&nbsp;<%=GetSpecialFieldNameByNumber(classInfoRs("classSpecialFieldNumber"))%></td>
	  <td nowrap>&nbsp;<%=classInfoRs("classBeginTime")%></td>
	  <td nowrap>&nbsp;<%=classInfoRs("classTeacherCharge")%></td>
	  <td nowrap>&nbsp;<%=classInfoRs("classYearsTime")%></td>
   <td nowrap>&nbsp;<a href="classInfoEdit.asp?classNumber=<%=classInfoRs("classNumber")%>"><img src="../images/edit.gif" border=0 height=12 width=12>��ϸ</a>&nbsp;&nbsp;<a href="classInfoDel.asp?classNumber=<%=classInfoRs("classNumber")%>" onClick="javascript:return confirm('��ľ���ɾ���˼�¼��?');"><img src="../images/delete.gif" border=0 height=12 width=12>ɾ��</a></td>
  </tr>
  <%
        classInfoRs.MoveNext
		  End If
	  Next
	%>
  <tr> 
		      <td  align="right" height="22" colspan="8" bgcolor="#ffffff"> 
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
		        If intPage = classInfoRs.PageCount or classInfoRs.PageCount=0 Then
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
					If classInfoRs.PageCount <= 0 Then
						Response.Write "<option value=''>��</option>"
					Else
						For intLoop = 1 To classInfoRs.PageCount
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
