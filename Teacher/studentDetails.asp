<!--#include file="../Database/conn.asp"-->
<!--#include file="../System/md5.asp"--> 
<!--#include file="../System/function.asp"-->
<%
  '如果管理员还没有登陆
  if session("teacherNumber")="" and  session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
  '得到某个学生的详细信息
  set studentDetailRs = Server.CreateObject("ADODB.RecordSet")
  sqlString = "select * from [studentInfo] where studentNumber='" & Request("studentNumber") & "'"
  studentDetailRs.Open sqlString,conn,1,1
%>
<HTML>
<HEAD>
	<Title>学生详细信息</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>学生信息管理--&gt;学生详细信息
			 </td>
	   </tr><br>
		 <%
		   '如果该学生设置了图片则显示该学生的头像
		   if studentDetailRs("studentPhoto") <> "" then
			   Response.Write "<tr><td>学生头像:</td><td><img src='../admin/" & studentDetailRs("studentPhoto") & "' border=0 height=100 width=100></td></tr>"
			 end if 
		 %>
		 <tr>
			 <td>所在班级:</td>
			 <td>
			   <%=GetClassNameByNumber(studentDetailRs("studentClassNumber"))%>
			 </td>
		 </tr>
	   <tr>
	     <td style="height: 26px">
		     学号:</td><td><%=studentDetailRs("studentNumber")%></td>
			 </td>
		 </tr>
		 <tr>
		  <td>学生姓名:</td><td><%=studentDetailRs("studentName")%></td>
		 </tr>
		 <tr>
		   <td>性别:</td>
			 <td>
			   <%=studentDetailRs("studentSex")%>
			 </td>
		 </tr>
		 <tr>
		   <td>学生生日:</td>
			 <td>
			   <%=studentDetailRs("studentBirthday")%>
			</td>
		 </tr>
		 <tr>
		   <td>政治面貌:</td>
			 <td><%=studentDetailRs("studentState")%>
			 </td>
		 </tr>
		  <tr>
		    <td>家庭地址:</td>
			  <td><%=studentDetailRs("studentAddress")%></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">附加信息:</td>
		    <td><textarea cols=40 rows=5 name=studentMemo><%=studentDetailRs("studentMemo")%></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="2" align="center">
		      <input type="button" value=" 返回" onClick="javascript:location.href='studentInfoQuery.asp?studentNumber=<%=Request("studentQueryNumber")%>&studentName=<%=Request("studentQueryName")%>&studentClass=<%=Request("studentQueryClass")%>';">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>
</HTML>
