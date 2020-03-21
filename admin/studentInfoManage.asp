<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/config.asp"-->
<%
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  sqlString = "select * from [studentInfoView] where 1=1"
  studentName = Request("studentName")
  studentNumber = Request("studentNumber")
  studentClass = Request("studentClass")
  '根据不同的组合条件进行sql语句的构造
  if studentName <> "" then
    sqlString = sqlString & " and studentName like '%" & studentName & "%'"
	end if
	if studentNumber <> "" then
	  sqlString = sqlString & " and studentNumber like '%" & studentNumber & "%'"
	end if
	if studentClass <> "" then
    sqlString = sqlString & " and classNumber like '%" & studentClass & "%'"
	end if
	set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
	studentInfoRs.Open sqlString,conn,1,1
  If Request("Page") = "" Then
		intPage = 1
	Else
		intPage = Clng(Request("Page"))
	End If
	'设置每页显示的记录数
	studentInfoRs.PageSize = pageSize 
	If intPage > studentInfoRs.PageCount Then
		intPage = studentInfoRs.PageCount
	End If
	If intPage <= 0 Then
		intPage = 1
	End If
	If Not studentInfoRs.EOF Then
		studentInfoRs.AbsolutePage = intPage
	End If
%>
<HTML>
<HEAD>
	<Title>学生信息管理</Title>
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
	<table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#FFFFFF'>
 <form name="form1" method="post" action="studentInfoManage.asp">
 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=8 align="center">
		      <img src="../images/list.gif" width=14px height=14px>学生信息管理--&gt;学生信息列表
			 </td>
	   </tr><br>
 <tr>
	<td  align="left" height="22" colspan="8" bgcolor="#ffffff"> 
	  学号:<input type="text" name=studentNumber size=18 value="<%=studentNumber%>">&nbsp;
		姓名:<input type="text" name="studentName" size="15"  value="<%=studentName%>">&nbsp;
		班级:<select name=studentClass>
		       <option value="">请选择学生所在班级</option>
				   <%
					   sqlString = "select classNumber,className from [classInfo]"
						 set classInfoRs = Server.CreateObject("ADODB.RecordSet")
						 classInfoRs.Open sqlString,conn,1,1
						 while not classInfoRs.EOF
						   Response.Write "<option value='" & classInfoRs("classNumber") & "'>" & classInfoRs("className") & "</option>"
							 classInfoRs.MoveNext
						 wend
					 %>
			   </select>
		<input type="submit" value=" 检索 " class="button1">
	<td>
 </tr>
  <tr> 
		<td  nowrap> 
      <div align="center">学号</div>
    </td>
	 <td>
	   <div align="center">姓名</div>
		</td>
		<td>
	   <div align="center">性别</div>
		</td>
		<td>
	   <div align="center">政治面貌</div>
		</td>
		<td>
		  <div align="center">所在学院</div>
		</td>
		<td>
		  <div align="center">所在专业</div>
		</td>
		<td>
		  <div align="center">所在班级</div>
		</td>
		<td>
		  <div align="center">操作</div>
		</td>
  </tr>
  <%
    for i=0 to studentInfoRs.PageSize-1
	    if not studentInfoRs.EOF then
  %>
  <tr align="center"> 
    <td nowrap>&nbsp;<%=studentInfoRs("studentNumber")%></td>
    <td nowrap>&nbsp;<%=studentInfoRs("studentName")%></td>
    <td nowrap>&nbsp;<%=studentInfoRs("studentSex")%></td>
	  <td nowrap>&nbsp;<%=studentInfoRs("studentState")%></td>
	  <td nowrap>&nbsp;<%=studentInfoRs("collegeName")%></td>
	  <td nowrap>&nbsp;<%=studentInfoRs("specialFieldName")%></td>
    <td nowrap>&nbsp;<%=studentInfoRs("className")%></td>
    <td nowrap>&nbsp;<a href="studentInfoDetail.asp?studentNumber=<%=studentInfoRs("studentNumber")%>"><img src="../images/edit.gif" border=0 height=12 width=12>详细</a>&nbsp;&nbsp;<a href="studentInfoDel.asp?studentNumber=<%=studentInfoRs("studentNumber")%>" onClick="javascript:return confirm('真的决定删除此记录吗?');"><img src="../images/delete.gif" border=0 height=12 width=12>删除</a></td>
  </tr>
  <%
        studentInfoRs.MoveNext
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
		        If intPage = studentInfoRs.PageCount or studentInfoRs.PageCount=0 Then
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
					If studentInfoRs.PageCount <= 0 Then
						Response.Write "<option value=''>无</option>"
					Else
						For intLoop = 1 To studentInfoRs.PageCount
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
