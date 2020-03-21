<!--#include virtual="/Database/conn.asp"-->
<%
  'errMessage保存错误信息
  dim errMessage,sqlString
  errMessage = ""
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'如果管理员添加了新的班级信息并提交
	if Request("submit") <> "" then
	  '如果没有选择所在的专业
	  if Request("classSpecialFieldNumber") = "" then
	    errMessage = "请选择所在专业"
	  end if 
	  if Request("classBeginTime") = "" then
	    errMessage = "请选择班级开始时间"
	  end if
	  '如果没有输入班级名称
	  if Request("className") = "" then
	    errMessage = "请输入班级名称"
	  end if
	  '检查该班级名称是否存在
	  sqlString = "select * from [classInfo] where className='" & Request("className") & "'"
	  set classInfoRs = Server.CreateObject("ADODB.RecordSet")
	  classInfoRs.Open sqlString,conn,1,1
	  if not classInfoRs.EOF then
	    errMessage = "该班级名称信息已经存在"
		end if
		classInfoRs.Close
		
		'根据errMessage判断是否要进入新班级信息的添加流程
		if errMessage = "" then
		  '下面开始进入新班级信息添加程序流程
		  sqlString = "select * from [classInfo]"
		  classInfoRs.Open sqlString,conn,1,3
		  classInfoRs.AddNew
		  classInfoRs("classNumber") = Request("classNumber")
		  classInfoRs("className") = Request("className")
		  classInfoRs("classSpecialFieldNumber") = Request("classSpecialFieldNumber")
		  classInfoRs("classBeginTime") = CDate(Request("classBeginTime"))
		  classInfoRs("classYearsTime") = CInt(Request("classYearsTime"))
		  classInfoRs("classTeacherCharge") = Trim(Request("classTeacherCharge"))
		  classInfoRs("classMemo") = Trim(Request("classMemo"))
		  classInfoRs.Update
		  classInfoRs.Close
		  Response.Write "<script>alert('班级信息添加成功!');</script>"
	  else
	    Response.Write "<script>alert('" & errMessage & "');</script>"
	  end if
	end if
%>
<HTML>
<HEAD>
	<Title>班级信息添加</Title>
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
	     <td style="height:14px" colspan=2 align="center">
		  <img src="../images/ADD.gif" width=14px height=14px>班级信息管理--&gt;班级信息添加<br>
			 </td>
	   </tr>
		 <tr>
		   <td width=100 align="right">选择专业:
			 <td>
			   <select name=classSpecialFieldNumber>
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
			 </td>
		 </tr>
		 <tr>
		    <td width=100 align="right">班级编号:</td>
		    <td><input type=text name=classNumber size=20></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">班级名称:</td>
		    <td><input type=text name=className size=20></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">班级开始时间:</td>
		    <td><input type=text name=classBeginTime width=25px>
			    <input class="submit" name="Button" onclick="seltime('classBeginTime');" style="width:30px" type="button" value="选择">
				</td>
		  </tr>
		   <tr>
		    <td width=100 align="right">班级时长:</td>
		    <td><select name=classYearsTime>
			        <option value="3">3</option>
					    <option value="4">4</option>
					 </select>年制</td>
		  </tr>
			<tr>
		    <td width=100 align="right">班级任姓名:</td>
		    <td><input type=text name=classTeacherCharge size=20></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">附加信息:</td>
		    <td><textarea cols=40 rows=5 name=classMemo></textarea></td>
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
