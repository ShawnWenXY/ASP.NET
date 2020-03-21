<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/config.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  dim sqlString
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  '查询班级课程的sql语句
  sqlString = "select * from [publicCourseInfo] where 1=1"
  '判断查询的课程编号是否为空完善查询sql语句
  if Request("courseNumber") <> "" then
    sqlString = sqlString & " and courseNumber like '%" & Trim(Request("courseNumber")) & "%'"
  end if
  '判断查询的课程名称是否为空完善查询sql语句
  if Request("courseName") <> "" then
    sqlString = sqlString & " and courseName like '%" & Trim(Request("courseName")) & "%'"
  end if
  '判断查询的专业信息是否存在完善查询的sql语句
  if Request("specialFieldNumber") <> "" then
    sqlString = sqlString & " and specialFieldNumber='" & Request("specialFieldNumber") & "'"
  end if
  '判断查询的学期信息是否存在完善查询的sql语句
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
	'设置每页显示的记录数
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
	<Title>选修课程信息管理</Title>
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
		      <img src="../images/list.gif" width=14px height=14px>课程信息管理--&gt;选修课程信息列表
			 </td>
	   </tr><br>
 <tr>
	<td  align="left" height="22" colspan="7" bgcolor="#ffffff"> 
	　课程编号:<input type=text name=courseNumber size=8>&nbsp;
	  课程名称:<input type=text name=courseName size=8>&nbsp;
	  专业:<select name=specialFieldNumber>
	             <option value="">请选择</option>
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
		学期:<select name=termId>
		       <option value="">请选择</option>
				   <%
					   sqlString = "select * from [termInfo]"
						 set termInfoRs = Server.CreateObject("ADODB.RecordSet")
						 termInfoRs.Open sqlString,conn,1,1
						  while not termInfoRs.EOF
						    Response.Write "<option value='" & termInfoRs("termId") & "'>" & termInfoRs("termBeginYear") & "-" & termInfoRs("termEndYear") & "年" & termInfoRs("termUpOrDown") & "</option>"
							  termInfoRs.MoveNext
						  wend
						  termInfoRs.Close
					 %>
			   </select>
		<input type="submit" value=" 检索 " class="button1">
	<td>
 </tr>
  <tr> 
		<td  nowrap> 
      <div align="center">课程编号</div>
    </td>
	 <td>
	   <div align="center">课程名称</div>
		</td>
		<td>
	   <div align="center">开设专业</div>
		</td>
		<td>
	   <div align="center">开设学期</div>
		</td>
		<td>
	   <div align="center">课程学分</div>
		</td>
		<td>
		  <div align="center">操作</div>
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
   <td nowrap>&nbsp;<a href="publicCourseInfoEdit.asp?courseNumber=<%=publicCourseInfoRs("courseNumber")%>"><img src="../images/edit.gif" border=0 height=12 width=12>编辑</a>&nbsp;&nbsp;<a href="publicCourseInfoDel.asp?courseNumber=<%=publicCourseInfoRs("courseNumber")%>" onClick="javascript:return confirm('真的决定删除此记录吗?');"><img src="../images/delete.gif" border=0 height=12 width=12>删除</a></td>
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
					Response.Write "前一页"
		        Else
		        %>
		        <a href="javascript:formsubmit('<%=intPage - 1%>');">前一页</a> 
		        <%
		        End If
		        %>
		        &nbsp;&nbsp; 
		        <%
		        If intPage = publicCourseInfoRs.PageCount or publicCourseInfoRs.PageCount=0 Then
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
					If publicCourseInfoRs.PageCount <= 0 Then
						Response.Write "<option value=''>无</option>"
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
		        </select>页
		      </td>
		    </tr>
	    <input type="hidden" name="page" value="">
    </form>
</table>
</BODY>
</HTML>
