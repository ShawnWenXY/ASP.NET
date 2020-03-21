<!--#include file="../Database/conn.asp"-->
<%
  '如果教师还没有登陆
  if session("teacherNumber")="" and  session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'如果更新了教师的相关信息并提交时
	if Request("submit") <> "" then
	  sqlString = "select * from [teacherInfo] where teacherNumber='" & Session("teacherNumber") & "'"
	  set teacherInfoRs = Server.CreateObject("ADODB.RecordSet")
	  teacherInfoRs.Open sqlString,conn,1,3
	  teacherInfoRs("teacherName") = Request("teacherName")
	  teacherInfoRs("teacherSex") = Request("teacherSex")
	  '如果重新选择了教师的个人头像
	  if Request("photoAddress") <> "" then
	    teacherInfoRs("teacherPhoto") = Trim(Request("photoAddress"))
	  end if
	  teacherInfoRs("teacherBirthday") = CDate(Request("teacherBirthday"))
	  teacherInfoRs("teacherArriveTime") = CDate(Request("teacherArriveTime"))
	  teacherInfoRs("teacherCardNumber") = Trim(Request("teacherCardNumber"))
	  teacherInfoRs("teacherAddress") = Trim(Request("teacherAddress"))
	  teacherInfoRs("teacherPhone") = Trim(Request("teacherPhone"))
	  teacherInfoRs("teacherMemo") = Trim(Request("teacherMemo"))
	  teacherInfoRs.Update
	  teacherInfoRs.Close
	  Response.Write "<script>alert('教师信息更新成功!');</script>"
	end if
  '得到某个教师的详细信息
  set teacherDetailRs = Server.CreateObject("ADODB.RecordSet")
  sqlString = "select * from [teacherInfo] where teacherNumber='" & Session("teacherNumber") & "'"
  teacherDetailRs.Open sqlString,conn,1,1
%>
<HTML>
<HEAD>
	<Title>个人信息修改</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript">
	function seltime(inputName)
	{
	  window.open('../admin/seltime.asp?InputName='+inputName+'','','width=250,height=220,left=360,top=250,scrollbars=yes');  
	}
	</script>
</HEAD>
<BODY>
<form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/edit.gif" width=14px height=14px>系统管理--&gt;个人信息修改
			 </td>
	   </tr><br>
		 <%
		   '如果该教师设置了图片则显示该教师的头像
		   if teacherDetailRs("teacherPhoto") <> "" then
			   Response.Write "<tr><td>教师头像:</td><td><img src='../admin/" & teacherDetailRs("teacherPhoto") & "' border=0 height=100 width=100></td></tr>"
			 end if 
		 %>
	   <tr>
	     <td style="height: 26px">
		     教师职工编号:</td><td><%=teacherDetailRs("teacherNumber")%></td>
			 </td>
		 </tr>
		 <tr>
		  <td>教师姓名:</td><td><input type=text name=teacherName size=20 value=<%=teacherDetailRs("teacherName")%>></td>
		 </tr>
		 <tr>
		   <td>性别:</td>
			 <td>
			   <select name=teacherSex>
			   <%
				   if teacherDetailRs("teacherSex") = "男" then
					   Response.Write "<option value='男'>男</option><option value='女'>女</option>"
					 else
					   Response.Write "<option value='女'>女</option><option value='男'>男</option>"
					 end if
				 %>
			 </td>
		 </tr>
		 <tr>
		   <td>教师生日:</td>
			 <td>
			   <input type=text name=teacherBirthday width=77px value=<%=teacherDetailRs("teacherBirthday")%>>
				 <input class="submit" name="Button" onclick="seltime('teacherBirthday');" style="width:30px" type="button" value="选择">
			 </td>
		 </tr>
		 <tr>
		   <td>入校时间:</td>
			 <td>
			   <input type=text name=teacherArriveTime width=77px value=<%=teacherDetailRs("teacherArriveTime")%>>
				 <input class="submit" name="Button" onclick="seltime('teacherArriveTime');" style="width:30px" type="button" value="选择">
			 </td>
		 </tr>
		 <tr>
			  <td>新照片路径:</td>
			  <td><input type="text" name=photoAddress size=20 readonly>*请在下面上传照片,程序会自动生成路径</td>
			</tr>
			<tr> 
       <td>新照片上传：</td>
       <td bgcolor="#F5F5F5" height="30" align="center" width="79%">
		     <iframe marginwidth=0 marginheight=0  frameborder=0 scrolling=no src='../admin/upload.asp' width=450 height=30></iframe> 
       </td>
      </tr>
		  <tr>
		    <td>身份证号:</td>
			  <td><input type=text name=teacherCardNumber size=50 value=<%=teacherDetailRs("teacherCardNumber")%>></td>
		  </tr>
		  <tr>
		    <td>家庭地址:</td>
			  <td><input type=text name=teacherAddress size=50 value=<%=teacherDetailRs("teacherAddress")%>></td>
		  </tr>
		  <tr>
		    <td>电话:</td>
			  <td><input type=text name=teacherPhone size=50 value=<%=teacherDetailRs("teacherPhone")%>></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">附加信息:</td>
		    <td><textarea cols=40 rows=5 name=teacherMemo><%=teacherDetailRs("teacherMemo")%></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" 确认更新 ">
		      <input type="button" value=" 返回" onClick="javascript:location.href='teacherInfoManage.asp';">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>
</HTML>
