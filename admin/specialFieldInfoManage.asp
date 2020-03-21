<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/config.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  '判断是否已经登录，如果没有登录则跳转到登录页面
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if

  sqlString = "select * from [specialFieldInfo] where 1=1"
  '获取传递过来的查询参数
  specialFieldNumber = Request("specialFieldNumber")
  specialFieldName = Request("specialFieldName")
  specialFieldCollegeNumber = Request("specialFieldCollegeNumber")
  '如果查询的专业编号信息关键字不为空则拼接成查询条件
  if specialFieldNumber <> "" then
    sqlString = sqlString & " and specialFieldNumber like '%" & specialFieldNumber & "%'"
  end if
  if specialFieldName <> "" then
    sqlString = sqlString & " and specialFieldName like '%" & specialFieldName & "%'"
  end if
  '如果查询的学院信息不为空就将条件附加到sql语句中
  if specialFieldCollegeNumber <> "" then
    sqlString = sqlString & " and  specialCollegeNumber='" & specialFieldCollegeNumber & "'"
  end if
  '声明数据集
  set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
  '查询数据库 将数据填充进数据集
 ' response.Write(sqlString)
  'response.End()
  specialFieldInfoRs.Open sqlString,conn,1,3
  
  If Request("Page") = "" Then
		intPage = 1
	Else
		intPage = Clng(Request("Page"))
	End If
	'设置每页显示的记录数
	specialFieldInfoRs.PageSize = pageSize
	If intPage > specialFieldInfoRs.PageCount Then
		intPage = specialFieldInfoRs.PageCount
	End If
	If intPage <= 0 Then
		intPage = 1
	End If
	If Not specialFieldInfoRs.EOF Then
		'设置当前为第几页
		specialFieldInfoRs.AbsolutePage = intPage
	End If
%>
<HTML>
<HEAD>
	<Title>专业信息管理</Title>
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
		      <img src="../images/list.gif" width=14px height=14px>班级信息管理--&gt;专业信息列表
			 </td>
	   </tr><br>
 <tr>
	<td  align="left" height="22" colspan="4" bgcolor="#ffffff"> 
	  专业编号:&nbsp;<input type=text name=specialFieldNumber size=10>&nbsp;
	  专业名称:&nbsp;<input type=text name=specialFieldName size=10>&nbsp;
		所在学院:<select name="specialFieldCollegeNumber">
		       <option value="">请选择</option>
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
		<input type="submit" value=" 检索 " class="button1">
	<td>
 </tr>
  <tr> 
		<td  nowrap> 
      <div align="center">专业编号</div>
    </td>
	 <td>
	   <div align="center">专业名称</div>
		</td>
		<td>
	   <div align="center">所在学院</div>
		</td>
		<td>
		  <div align="center">删除</div>
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
    <td nowrap>&nbsp;<a href="specialFieldInfoDel.asp?specialFieldNumber=<%=specialFieldInfoRs("specialFieldNumber")%>" onClick="javascript:return confirm('真的决定删除此记录吗?');"><img src="../images/delete.gif" border=0 height=12 width=12>删除</a></td>
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
					Response.Write "前一页"
		        Else
		        %>
		        <a href="javascript:formsubmit('<%=intPage - 1%>');">前一页</a> 
		        <%
		        End If
		        %>
		        &nbsp;&nbsp; 
		        <%
		        If intPage = specialFieldInfoRs.PageCount or specialFieldInfoRs.PageCount=0 Then
					Response.Write "下一页"
		        Else
		        %>
		        <a href="javascript:formsubmit('<%=intPage + 1%>');">下一页</a> 
		        <%
		        End If
		        %>
		        &nbsp; 转向 
		        <select name="selectpage" onChange="changepage();">
		          <%
					If specialFieldInfoRs.PageCount <= 0 Then
						Response.Write "<option value=''>无</option>"
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
		        </select>页
		      </td>
		    </tr>
	    <input type="hidden" name="page" value="">
    </form>
</table>
</BODY>
</HTML>
