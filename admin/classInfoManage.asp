<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/config.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  sqlString = "select * from [classInfo] where 1=1"
  '取得查询的班级编号关键字信息
  classNumber = Trim(Request("classNumber"))
  if classNumber <> "" then
    sqlString = sqlString & " and classNumber like '%" & classNumber & "%'"
  end if
  '取得查询的班级名称的关键字信息
  className = Trim(Request("className"))
  if className <> "" then
    sqlString = sqlString & " and className like '%" & className & "%'"
  end if
  '取得查询的专业编号信息
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
	'设置每页显示的记录数
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
	<Title>班级信息管理</Title>
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
		      <img src="../images/list.gif" width=14px height=14px>班级信息管理--&gt;班级信息列表
			 </td>
	   </tr><br>
 <tr>
	<td  align="left" height="22" colspan="8" bgcolor="#ffffff"> 
	  班级编号:&nbsp;<input type=text name=classNumber size=15>&nbsp;&nbsp;
	  班级名称:&nbsp;<input type=text name=className size=10>&nbsp;&nbsp;
		班级所在专业:<select name=classSpecialFieldNumber>
		       <option value="">请选择</option>
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
		<input type="submit" value=" 检索 " class="button1">
	<td>
 </tr>
  <tr> 
		<td  nowrap> 
      <div align="center">班级编号</div>
    </td>
	 <td>
	   <div align="center">班级名称</div>
		</td>
		<td>
	   <div align="center">所在专业</div>
		</td>
		<td>
	   <div align="center">成立时间</div>
		</td>
		<td>
	   <div align="center">班主任姓名</div>
		</td>
		<td>
	   <div align="center">学制</div>
		</td>
		<td>
		  <div align="center">操作</div>
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
   <td nowrap>&nbsp;<a href="classInfoEdit.asp?classNumber=<%=classInfoRs("classNumber")%>"><img src="../images/edit.gif" border=0 height=12 width=12>详细</a>&nbsp;&nbsp;<a href="classInfoDel.asp?classNumber=<%=classInfoRs("classNumber")%>" onClick="javascript:return confirm('真的决定删除此记录吗?');"><img src="../images/delete.gif" border=0 height=12 width=12>删除</a></td>
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
					Response.Write "前一页"
		        Else
		        %>
		        <a href="javascript:formsubmit('<%=intPage - 1%>');">前一页</a> 
		        <%
		        End If
		        %>
		        &nbsp;&nbsp; 
		        <%
		        If intPage = classInfoRs.PageCount or classInfoRs.PageCount=0 Then
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
					If classInfoRs.PageCount <= 0 Then
						Response.Write "<option value=''>无</option>"
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
		        </select>页
		      </td>
		    </tr>
	    <input type="hidden" name="page" value="">
    </form>
</table>
</BODY>
</HTML>
