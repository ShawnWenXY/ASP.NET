<!--#include virtual="/DataBase/conn.asp"-->
<%
  'errMessage保存错误信息
  dim errMessage,sqlString
  errMessage = ""
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'如果管理员修改了班级信息并提交
	if Request("submit") <> "" then
	  '如果没有输入班级名称
	  if Request("className") = "" then
	    errMessage = "请输入班级名称"
	  end if
		
		'根据errMessage判断是否要进入新班级信息的添加流程
		if errMessage = "" then
		  '下面开始进入新班级信息添加程序流程
		  sqlString = "select * from [classInfo] where classNumber='" & Request("classNumber") & "'"
		  set classInfoRs = Server.CreateObject("ADODB.RecordSet")
		  classInfoRs.Open sqlString,conn,1,3
		  classInfoRs("className") = Request("className")
		  classInfoRs("classSpecialFieldNumber") = Request("classSpecialFieldNumber")
		  classInfoRs("classBeginTime") = CDate(Request("classBeginTime"))
		  classInfoRs("classYearsTime") = CInt(Request("classYearsTime"))
		  classInfoRs("classTeacherCharge") = Trim(Request("classTeacherCharge"))
		  classInfoRs("classMemo") = Trim(Request("classMemo"))
		  classInfoRs.Update
		  classInfoRs.Close
		  Response.Write "<script>alert('班级信息修改成功!');</script>"
	  else
	    Response.Write "<script>alert('" & errMessage & "');</script>"
	  end if
	end if
	
	'根据班级编号取得某个班级的信息
	classNumber = Request("classNumber")
	sqlString = "select * from [classInfo] where classNumber='" & classNumber & "'"
	set classInfoRs = Server.CreateObject("ADODB.RecordSet")
	classInfoRs.Open sqlString,conn,1,1
%>

<HTML>
<HEAD>
	<Title>班级信息修改</Title>
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
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>班级信息管理--&gt;班级信息修改
			 </td>
	   </tr>
		 <tr>
		    <td width=100 align="right">班级编号:</td>
		    <td><%=Request("classNumber")%><input type=hidden name=classNumber size=20 value='<%=Request("classNumber")%>'></td>
		  </tr>
		 <tr>
		   <td width=100px align="right">选择专业:
			 <td>
			   <select name=classSpecialFieldNumber>
				   <option value="">请选择</option>
				  <%
					  sqlString = "select specialFieldNumber,specialFieldName from [specialFieldInfo]"
					  set specialFieldInfoRs = Server.CreateObject("ADODB.RecordSet")
					  specialFieldInfoRs.Open sqlString,conn,1,1
					  while not specialFieldInfoRs.EOF
					    selected = ""
						  if specialFieldInfoRs("specialFieldNumber") = classInfoRs("classSpecialFieldNumber") then
						    selected = "selected"
							end if
					    Response.Write "<option value='" & specialFieldInfoRs("specialFieldNumber") & "' " & selected & ">" & specialFieldInfoRs("specialFieldName") & "</option>"
						  specialFieldInfoRs.MoveNext
					  wend
				  %>
				 </select>
			 </td>
		 </tr>
		 
		  <tr>
		    <td width=100 align="right">班级名称:</td>
		    <td><input type=text name=className size=20 value='<%=classInfoRs("className")%>'></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">班级开始时间:</td>
		    <td><input type=text name=classBeginTime width=60px value='<%=classInfoRs("classBeginTime")%>'>
			    <input class="submit" name="Button" onclick="seltime('classBeginTime');" style="width:30px" type="button" value="选择">
				</td>
		  </tr>
		   <tr>
		    <td width=100 align="right">班级时长:</td>
		    <td><select name=classYearsTime>
			       <%
					       if CInt(classInfoRs("classYearsTime")) = 3 then
							     Response.Write "<option value='3'>3</option><option value='4'>4</option>"
								 else
								   Response.Write "<option value='4'>4</option><option value='3'>3</option>"
								 end if
					   %>
					 </select>年制</td>
		  </tr>
			<tr>
		    <td width=100 align="right">班级任姓名:</td>
		    <td><input type=text name=classTeacherCharge size=20 value='<%=classInfoRs("classTeacherCharge")%>'></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">附加信息:</td>
		    <td><textarea cols=40 rows=5 name=classMemo><%=classInfoRs("classMemo")%></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" 确认修改 ">
		      <input type="button" value="返回" onclick="javascript:location.href='classInfoManage.asp'">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>

</HTML>