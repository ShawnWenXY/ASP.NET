<!--#include file="../DataBase/conn.asp"-->
<!--#include file="../System/md5.asp"-->
<%
  '保存错误信息,如果没有错误则为空
  dim errMessage
  errMessage = ""
  '如果学生还没有登陆
  if session("studentNumber")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'如果要求修改为新的密码
	if Request("submit") <> "" then
	  oldPassword = Trim(Request("oldPassword"))
	  newPassword = Trim(Request("newPassword"))
	  newPasswordAgain = Trim(Request("newPasswordAgain"))
	  if oldPassword = "" then
	    errMessage = errMessage & "请输入原来的密码!"
		end if
		if newPassword = "" then
		  errMessage = errMessage & "请输入新的密码"
	  end if
	  if newPassword <> newPasswordAgain then
	    errMessage = errMessage & "两次输入的的新密码不一致"
		end if
		if errMessage = "" then
		  sqlString = "select * from [studentInfo] where studentNumber='" & session("studentNumber") & "'"
		  set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
		  studentInfoRs.Open sqlString,conn,1,3
		  '判断输入的原来密码是否正确
		  if studentInfoRs("studentPassword") <> md5(oldPassword) then
		    Response.Write "<script>alert('你输入的旧密码不正确!');</script>"
			else
			  studentInfoRs("studentPassword") = md5(newPassword)
			  studentInfoRs.Update
			  Response.Write "<script>alert('密码修改成功!');</script>"
			end if
			studentInfoRs.Close
		else
		  Response.Write "<script>alert('" & errMessage & "');</script>"
		end if
	end if
%>
<HTML>
<HEAD>
	<Title>登陆密码设置</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>系统管理--&gt;密码设置
			 </td>
	   </tr>
		 <tr>
		   <td width=100 align="right">原来密码:</td>
			 <td><input type=password name=oldPassword size=20></td>
		</tr>
		<tr>
		  <td width=100 align="right">新密码:</td>
		  <td><input type=password name=newPassword size=20></td>
		</tr>
		<tr>
		  <td width=100 align="right">新密码确认:</td>
		  <td><input type=password name=newPasswordAgain></td>
		</tr>
		<tr>
		  <td colspan=2 align="center">
		    <input type="submit" name=submit value="修改密码">&nbsp;&nbsp;
			  <input type="reset" value="取消">
		  </td>
		</tr>
	</table>
</form>
</body>
</html>
