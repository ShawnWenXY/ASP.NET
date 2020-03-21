<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/config.asp"-->
<%
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  sqlString = "select * from [teacherInfo] where 1=1"
  teacherName = Request("teacherName")
  teacherNumber = Request("teacherNumberNumber")
  '根据不同的组合条件进行sql语句的构造
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
	'设置每页显示的记录数
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
	<Title>教师信息管理</Title>
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
		      <img src="../images/list.gif" width=14px height=14px>教师信息管理--&gt;教师信息列表
			 </td>
	   </tr><br>
 <tr>
	<td  align="left" height="22" colspan="7" bgcolor="#ffffff"> 
	  教职工编号:<input type="text" name=teacherNumber size=18 value="<%=teacherNumber%>">&nbsp;
		姓名:<input type="text" name="teacherName" size="15"  value="<%=teacherName%>">&nbsp;
		<input type="submit" value=" 检索 " class="button1">
	<td>
 </tr>
  <tr> 
		<td  nowrap> 
      <div align="center">教职工编号</div>
    </td>
	 <td>
	   <div align="center">姓名</div>
		</td>
		<td>
	   <div align="center">性别</div>
		</td>
		<td>
	   <div align="center">生日</div>
		</td>
		<td>
		  <div align="center">入校时间</div>
		</td>
		<td>
		  <div align="center">教师电话</div>
		</td>
		<td>
		  <div align="center">操作</div>
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
    <td nowrap>&nbsp;<a href="teacherInfoDetail.asp?teacherNumber=<%=teacherInfoRs("teacherNumber")%>"><img src="../images/edit.gif" border=0 height=12 width=12>详细</a>&nbsp;&nbsp;<a href="teacherInfoDel.asp?teacherNumber=<%=teacherInfoRs("teacherNumber")%>" onClick="javascript:return confirm('真的决定删除此记录吗?');"><img src="../images/delete.gif" border=0 height=12 width=12>删除</a></td>
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
					Response.Write "前一页"
		        Else
		        %>
		        <a href="javascript:formsubmit('<%=intPage - 1%>');">前一页</a> 
		        <%
		        End If
		        %>
		        &nbsp;&nbsp; 
		        <%
		        If intPage = teacherInfoRs.PageCount or teacherInfoRs.PageCount=0 Then
					Response.Write "下一页"
		        Else
		        %>
		        <a href="javascript:formsubmit('<%=intPage + 1%>');">下一页</a> 
		        <%
		        End If
		        %>
		        &nbsp; 转向 
		        <select name="selectpage" onchange="changepage();">
		          <%
					If teacherInfoRs.PageCount <= 0 Then
						Response.Write "<option value=''>无</option>"
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
		        </select>页
		      </td>
		    </tr>
	    <input type="hidden" name="page" value="">
    </form>
</table>
</BODY>
</HTML>
