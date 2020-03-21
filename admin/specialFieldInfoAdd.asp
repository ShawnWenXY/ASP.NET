<!--#include virtual="/Database/conn.asp"-->
<%
  'errMessage保存错误信息
  dim errMessage,sqlString
  errMessage = ""
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'如果管理员添加了新的专业信息并提交
	if Request("submit") <> "" then
	  '如果没有选择所在的学院
	  if Request("specialCollegeNumber") = "" then
	    errMessage = "请选择所在学院"
	  end if
	  '如果没有填写专业编号
	  if Request("specialFieldNumber") = "" then
	    errMessage = "请填写专业编号"
	  end if
	  '如果没有填写专业名称
	  if Request("specialFieldName") = "" then
	    errMessage = "请填写专业名称信息"
	  end if
	  '检查系统中是否已经存在了该专业名称信息
	  sqlString = "select * from [specialFieldInfo] where specialFieldName='" & Request("specialFieldName") & "'"
	  set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
	  specialFieldInfoRs.Open sqlString,conn,1,1
	  if not specialFieldInfoRs.EOF then
	    errMessage = "该专业名称信息已经存在！"
		end if
	  specialFieldInfoRs.Close
	  '检查errMessage的值是否为空决定是否执行新专业信息的添加操作
	  if errMessage = "" then
	    sqlString = "select * from [specialFieldInfo]"
		  specialFieldInfoRs.Open sqlString,conn,1,3
		  specialFieldInfoRs.AddNew
		  specialFieldInfoRs("specialCollegeNumber") = Request("specialCollegeNumber")
		  specialFieldInfoRs("specialFieldNumber") = Request("specialFieldNumber")
		  specialFieldInfoRs("specialFieldName") = Request("specialFieldName")
		  specialFieldInfoRs.Update
		  specialFieldInfoRs.Close
		  Response.Write "<script>alert('专业信息添加成功!');</script>"
	  else
		  Response.Write "<script>alert('" & errMessage & "');</script>"
	  end if
	end if
%>
<HTML>
<HEAD>
	<Title>专业信息添加</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
		<script language=javascript>
	function seltime(inputName)
	{
	  window.open('seltime.asp?InputName='+inputName+'','','width=250,height=220,left=360,top=250,scrollbars=yes');  
	}
	</script>
</HEAD>
<BODY>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr>
	     <td class="th" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>班级信息管理--&gt;专业信息添加
			 </td>
	   </tr><br>
		 <tr>
		  <td width=100 align="right">选择学院:
			 <td>
			   <select name=specialCollegeNumber>
				   <option value="">请选择</option>
				  <%
					  sqlString = "select collegeNumber,collegeName from [collegeInfo]"
					  set collegeInfoRs = Server.CreateObject("ADODB.RecordSet")
					  collegeInfoRs.Open sqlString,conn,1,1
					  while not collegeInfoRs.EOF
					    Response.Write "<option value='" & collegeInfoRs("collegeNumber") & "'>" & collegeInfoRs("collegeName") & "</option>"
						  collegeInfoRs.MoveNext
					  wend
				  %>
				 </select>
			 </td>
		 </tr>
		 <tr>
		    <td width=100 align="right">专业编号:</td>
		    <td><input type=text name=specialFieldNumber size=20></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">专业名称:</td>
		    <td><input type=text name=specialFieldName size=20></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" 确认添加 ">
		      <input type="reset" value=" 重新填写 ">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>

</HTML>
