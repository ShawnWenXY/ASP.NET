<!--#include file="../DataBase/conn.asp"-->
<!--#include file="../System/config.asp"-->
<!--#include file="../System/function.asp"-->
<%
  '如果教师还没有登陆
  if session("teacherNumber")="" and  session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  '取得查询的关键字信息
  studentNumber = Trim(Request("studentNumber"))
  courseNumber = Trim(Request("courseNumber"))
  sqlString = "select * from [scoreInfo] where 1=1"
  if studentNumber <> "" then
    sqlString = sqlString & " and studentNumber like '%" & studentNumber & "%'"
  end if
  if courseNumber <> "" then
    sqlString = sqlString & " and courseNumber like '%" & courseNumber & "%'"
  end if
  set scoreInfoRs = Server.CreateObject("ADODB.RecordSet")
  scoreInfoRs.Open sqlString,conn,1,3
  If Request("Page") = "" Then
		intPage = 1
	Else
		intPage = Clng(Request("Page"))
	End If
	'设置每页显示的记录数
	scoreInfoRs.PageSize = pageSize 
	If intPage > scoreInfoRs.PageCount Then
		intPage = scoreInfoRs.PageCount
	End If
	If intPage <= 0 Then
		intPage = 1
	End If
	If Not scoreInfoRs.EOF Then
		scoreInfoRs.AbsolutePage = intPage
	End If
%>
<HTML>
<HEAD>
	<Title>成绩信息查询</Title>
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
 <form name="form1" method="post" action="scoreInfoQuery.asp">
 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=6 align="center">
		      <img src="../images/list.gif" width=14px height=14px>成绩信息管理--&gt;成绩信息查询
			 </td>
	   </tr>
 <tr>
	<td  align="left" height="22" colspan="8" bgcolor="#ffffff"> 
	  学号:<input type="text" name=studentNumber size=18 value="<%=studentNumber%>">&nbsp;
	  课程编号:<input type="text" name="courseNumber" size="15"  value="<%=courseNumber%>">&nbsp;
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
	   <div align="center">课程编号</div>
		</td>
		<td>
	   <div align="center">课程名称</div>
		</td>
		<td>
	   <div align="center">课程类型</div>
		</td>
		<td>
		  <div align="center">成绩</div>
		</td>
  </tr>
  <%
    for i=0 to scoreInfoRs.PageSize-1
	    if not scoreInfoRs.EOF then
  %>
  <tr align="center"> 
    <td nowrap>&nbsp;<%=scoreInfoRs("studentNumber")%></td>
    <td nowrap>&nbsp;<%=GetStudentNameByNumber(scoreInfoRs("studentNumber"))%></td>
    <td nowrap>&nbsp;<%=scoreInfoRs("courseNumber")%></td>
	  <td nowrap>&nbsp;
	    <%
		    if CInt(scoreInfoRs("isSelect")) = 0 then
			    Response.Write GetClassCourseNameByNumber(scoreInfoRs("courseNumber"))
				else
				  Response.Write GetPublicCourseNameByNumber(scoreInfoRs("courseNumber"))
				end if
	    %>
	  </td>
	  <td nowrap>&nbsp;
		  <%
		    if CInt(scoreInfoRs("isSelect")) = 0 then
			    Response.Write "必修课"
				else
				  Response.Write "选修课"
				end if
		  %>
	  </td>
	  <td nowrap>&nbsp;<%=scoreInfoRs("score")%></td>
  </tr>
  
  <%
        scoreInfoRs.MoveNext
		  End If
	  Next
	%>
    
    <tr> 
		      <td  align="right" height="22" colspan="8" bgcolor="#ffffff">平均成绩:
              <%
			   sqlString = "SELECT AVG([score]) from [scoreInfo] where  1=1  "
 
 if studentNumber <> "" then
    sqlString = sqlString & " and studentNumber like '%" & studentNumber & "%'"
  end if
  if courseNumber <> "" then
    sqlString = sqlString & " and courseNumber like '%" & courseNumber & "%'"
  end if
  set rs = Server.CreateObject("ADODB.RecordSet")
 rs.Open sqlString,conn,1,1
  response.Write(rs(0))
			  %>
              不及格人数：
             <%
			   sqlString = "SELECT count(scoreId) from [scoreInfo] where score<60    "
  
  set rs1 = Server.CreateObject("ADODB.RecordSet")
 rs1.Open sqlString,conn,1,1
  response.Write(rs1(0))
			  %>
              弃考人数：
             
              <%
			   sqlString = " select COUNT(studentNumber) from studentInfo where studentNumber  not in (SELECT   [studentNumber] FROM [scoreInfo]  ) "
 
  set rs1 = Server.CreateObject("ADODB.RecordSet")
 rs1.Open sqlString,conn,1,1
  response.Write(rs1(0))
			  %>
              
              </td>
    </tr>
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
		        If intPage = scoreInfoRs.PageCount or scoreInfoRs.PageCount=0 Then
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
					If scoreInfoRs.PageCount <= 0 Then
						Response.Write "<option value=''>无</option>"
					Else
						For intLoop = 1 To scoreInfoRs.PageCount
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
