<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/config.asp"-->
<%
  '�������Ա��û�е�½
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  sqlString = "select * from [teacherInfo] where 1=1"
  teacherName = Request("teacherName")
  teacherNumber = Request("teacherNumberNumber")
  '���ݲ�ͬ�������������sql���Ĺ���
  if teacherName <> "" then
    sqlString = sqlString & " and teacherName like '%" & teacherName & "%'"
	end if
	if teacherNumber <> "" then
	  sqlString = sqlString & " and teacherNumber like '%" & teacherNumber & "%'"
	end if
	set teacherInfoRs = Server.CreateObject("ADODB.RecordSet")
	teacherInfoRs.Open sqlString,conn,1,1
  If Request("Page") = "" Then
		intPage = 1
	Else
		intPage = Clng(Request("Page"))
	End If
	'����ÿҳ��ʾ�ļ�¼��
	teacherInfoRs.PageSize = pageSize
	If intPage > teacherInfoRs.PageCount Then
		intPage = teacherInfoRs.PageCount
	End If
	If intPage <= 0 Then
		intPage = 1
	End If
	If Not teacherInfoRs.EOF Then
		teacherInfoRs.AbsolutePage = intPage
	End If
%>
<HTML>
<HEAD>
	<Title>��ʦ��Ϣ����</Title>
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
 <form name="form1" method="post" action="teacherInfoManage.asp">
 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=7 align="center">
		      <img src="../images/list.gif" width=14px height=14px>��ʦ��Ϣ����--&gt;��ʦ��Ϣ�б�
			 </td>
	   </tr><br>
 <tr>
	<td  align="left" height="22" colspan="7" bgcolor="#ffffff"> 
	  ��ְ�����:<input type="text" name=teacherNumber size=18 value="<%=teacherNumber%>">&nbsp;
		����:<input type="text" name="teacherName" size="15"  value="<%=teacherName%>">&nbsp;
		<input type="submit" value=" ���� " class="button1">
	<td>
 </tr>
  <tr> 
		<td  nowrap> 
      <div align="center">��ְ�����</div>
    </td>
	 <td>
	   <div align="center">����</div>
		</td>
		<td>
	   <div align="center">�Ա�</div>
		</td>
		<td>
	   <div align="center">����</div>
		</td>
		<td>
		  <div align="center">��Уʱ��</div>
		</td>
		<td>
		  <div align="center">��ʦ�绰</div>
		</td>
		<td>
		  <div align="center">����</div>
		</td>
  </tr>
  <%
    for i=0 to teacherInfoRs.PageSize-1
	    if not teacherInfoRs.EOF then
  %>
  <tr align="center"> 
    <td nowrap>&nbsp;<%=teacherInfoRs("teacherNumber")%></td>
    <td nowrap>&nbsp;<%=teacherInfoRs("teacherName")%></td>
    <td nowrap>&nbsp;<%=teacherInfoRs("teacherSex")%></td>
	  <td nowrap>&nbsp;<%=teacherInfoRs("teacherBirthday")%></td>
	  <td nowrap>&nbsp;<%=teacherInfoRs("teacherArriveTime")%></td>
	  <td nowrap>&nbsp;<%=teacherInfoRs("teacherPhone")%></td>
    <td nowrap>&nbsp;<a href="teacherInfoDetail.asp?teacherNumber=<%=teacherInfoRs("teacherNumber")%>"><img src="../images/edit.gif" border=0 height=12 width=12>��ϸ</a>&nbsp;&nbsp;<a href="teacherInfoDel.asp?teacherNumber=<%=teacherInfoRs("teacherNumber")%>" onClick="javascript:return confirm('��ľ���ɾ���˼�¼��?');"><img src="../images/delete.gif" border=0 height=12 width=12>ɾ��</a></td>
  </tr>
  <%
        teacherInfoRs.MoveNext
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
		        If intPage = teacherInfoRs.PageCount or teacherInfoRs.PageCount=0 Then
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
					If teacherInfoRs.PageCount <= 0 Then
						Response.Write "<option value=''>��</option>"
					Else
						For intLoop = 1 To teacherInfoRs.PageCount
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
