<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/config.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  '�ж��Ƿ��Ѿ���¼�����û�е�¼����ת����¼ҳ��
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if

  sqlString = "select * from [specialFieldInfo] where 1=1"
  '��ȡ���ݹ����Ĳ�ѯ����
  specialFieldNumber = Request("specialFieldNumber")
  specialFieldName = Request("specialFieldName")
  specialFieldCollegeNumber = Request("specialFieldCollegeNumber")
  '�����ѯ��רҵ�����Ϣ�ؼ��ֲ�Ϊ����ƴ�ӳɲ�ѯ����
  if specialFieldNumber <> "" then
    sqlString = sqlString & " and specialFieldNumber like '%" & specialFieldNumber & "%'"
  end if
  if specialFieldName <> "" then
    sqlString = sqlString & " and specialFieldName like '%" & specialFieldName & "%'"
  end if
  '�����ѯ��ѧԺ��Ϣ��Ϊ�վͽ��������ӵ�sql�����
  if specialFieldCollegeNumber <> "" then
    sqlString = sqlString & " and  specialCollegeNumber='" & specialFieldCollegeNumber & "'"
  end if
  '�������ݼ�
  set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
  '��ѯ���ݿ� �������������ݼ�
 ' response.Write(sqlString)
  'response.End()
  specialFieldInfoRs.Open sqlString,conn,1,3
  
  If Request("Page") = "" Then
		intPage = 1
	Else
		intPage = Clng(Request("Page"))
	End If
	'����ÿҳ��ʾ�ļ�¼��
	specialFieldInfoRs.PageSize = pageSize
	If intPage > specialFieldInfoRs.PageCount Then
		intPage = specialFieldInfoRs.PageCount
	End If
	If intPage <= 0 Then
		intPage = 1
	End If
	If Not specialFieldInfoRs.EOF Then
		'���õ�ǰΪ�ڼ�ҳ
		specialFieldInfoRs.AbsolutePage = intPage
	End If
%>
<HTML>
<HEAD>
	<Title>רҵ��Ϣ����</Title>
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
 <form name="form1" method="post" action="specialFieldInfoManage.asp">
 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=4 align="center">
		      <img src="../images/list.gif" width=14px height=14px>�༶��Ϣ����--&gt;רҵ��Ϣ�б�
			 </td>
	   </tr><br>
 <tr>
	<td  align="left" height="22" colspan="4" bgcolor="#ffffff"> 
	  רҵ���:&nbsp;<input type=text name=specialFieldNumber size=10>&nbsp;
	  רҵ����:&nbsp;<input type=text name=specialFieldName size=10>&nbsp;
		����ѧԺ:<select name="specialFieldCollegeNumber">
		       <option value="">��ѡ��</option>
				   <%
					   sqlString = "select * from [collegeInfo]"
						 set collegeInfoRs = Server.CreateObject("ADODB.RecordSet")
						 collegeInfoRs.Open sqlString,conn,1,1
						 while not collegeInfoRs.EOF
						   Response.Write "<option value='" & collegeInfoRs("collegeNumber") & "'>" & collegeInfoRs("collegeName") & "</option>"
							 collegeInfoRs.MoveNext
						 wend 
					 %>
			   </select>
		<input type="submit" value=" ���� " class="button1">
	<td>
 </tr>
  <tr> 
		<td  nowrap> 
      <div align="center">רҵ���</div>
    </td>
	 <td>
	   <div align="center">רҵ����</div>
		</td>
		<td>
	   <div align="center">����ѧԺ</div>
		</td>
		<td>
		  <div align="center">ɾ��</div>
		</td>
  </tr>
  <%
    for i=0 to specialFieldInfoRs.PageSize-1
	    if not specialFieldInfoRs.EOF then
  %>
  <tr align="center"> 
    <td nowrap>&nbsp;<%=specialFieldInfoRs("specialFieldNumber")%></td>
    <td nowrap>&nbsp;<%=specialFieldInfoRs("specialFieldName")%></td>
    <td nowrap>&nbsp;<%=GetCollegeNameByNumber(specialFieldInfoRs("specialCollegeNumber"))%></td>
    <td nowrap>&nbsp;<a href="specialFieldInfoDel.asp?specialFieldNumber=<%=specialFieldInfoRs("specialFieldNumber")%>" onClick="javascript:return confirm('��ľ���ɾ���˼�¼��?');"><img src="../images/delete.gif" border=0 height=12 width=12>ɾ��</a></td>
  </tr>
  <%
        specialFieldInfoRs.MoveNext
		  End If
	  Next
	%>
  <tr> 
		      <td  align="right" height="22" colspan="4" bgcolor="#ffffff"> 
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
		        If intPage = specialFieldInfoRs.PageCount or specialFieldInfoRs.PageCount=0 Then
					Response.Write "��һҳ"
		        Else
		        %>
		        <a href="javascript:formsubmit('<%=intPage + 1%>');">��һҳ</a> 
		        <%
		        End If
		        %>
		        &nbsp; ת�� 
		        <select name="selectpage" onChange="changepage();">
		          <%
					If specialFieldInfoRs.PageCount <= 0 Then
						Response.Write "<option value=''>��</option>"
					Else
						For intLoop = 1 To specialFieldInfoRs.PageCount
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
