<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/md5.asp"--> 
<%
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'如果更新了学生的相关信息并提交时
	if Request("submit") <> "" then
	  sqlString = "select * from [studentInfo] where studentNumber='" & Request("studentNumber") & "'"
	  set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
	  studentInfoRs.Open sqlString,conn,1,3
	  studentInfoRs("studentName") = Request("studentName")
	  studentInfoRs("studentSex") = Request("studentSex")
	  studentInfoRs("studentBirthday") = CDate(Request("studentBirthday"))
	  studentInfoRs("studentState") = Request("studentState")
	  '如果管理给学生设置新的登陆密码
	  if Request("studentPassword") <> "" then
	    studentInfoRs("studentPassword") = md5(Request("studentPassword"))
	  end if
	  '如果管理员给学生上传了新的图片 
	  if Request("photoAddress") <> "" then
	    studentInfoRs("studentPhoto") = Request("photoAddress")
		end if
	  studentInfoRs("studentAddress") = Request("studentAddress")
	  studentInfoRs("studentMemo") = Request("studentMemo")
	  studentInfoRs.Update
	  studentInfoRs.Close
	  Response.Write "<script>alert('学生信息更新成功!');</script>"
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
	<script language="javascript">
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
		      <img src="../images/edit.gif" width=14px height=14px>学生信息管理--&gt;学生详细信息
			 </td>
	   </tr><br>
		 <%
		   '如果该学生设置了图片则显示该学生的头像
		   if studentDetailRs("studentPhoto") <> "" then
			   Response.Write "<tr><td>学生头像:</td><td><img src='" & studentDetailRs("studentPhoto") & "' border=0 height=100 width=100></td></tr>"
			 end if 
		 %>
		 <tr>
			 <td>所在班级:</td>
			 <td>
			   <select name=studentClass id=studentClass>
				  <%
				    '得到所有的班级信息
				    set studentClassRs = Server.CreateObject("ADODB.RecordSet")
					  sqlString = "select classNumber,className from [classInfo]"
					  studentClassRs.Open sqlString,conn,1,1
					  while not studentClassRs.EOF
					    selected = ""
						  if studentClassRs("classNumber") = studentDetailRs("studentClassNumber") then
						    selected = "selected"
							end if
					    Response.Write "<option value='" & studentClassRs("classNumber") &"' " & selected & ">" & studentClassRs("className") & "</option>"
						  studentClassRs.MoveNext
					  wend
				  %>
				</select>
			 </td>
		 </tr>
	   <tr>
	     <td style="height: 26px">
		     学号:</td><td><%=studentDetailRs("studentNumber")%></td>
			   <input type="hidden" name=studentNumber value=<%=studentDetailRs("studentNumber")%>>
			 </td>
		 </tr>
		 <tr>
		  <td>学生姓名:</td><td><input type=text name=studentName size=20 value=<%=studentDetailRs("studentName")%>></td>
		 </tr>
		 <tr>
		   <td>性别:</td>
			 <td>
			   <select name=studentSex>
			   <%
				   if studentDetailRs("studentSex") = "男" then
					   Response.Write "<option value='男'>男</option><option value='女'>女</option>"
					 else
					   Response.Write "<option value='女'>女</option><option value='男'>男</option>"
					 end if
				 %>
			 </td>
		 </tr>
		 <tr>
		   <td>学生生日:</td>
			 <td>
			   <input type=text name=studentBirthday width=77px value=<%=studentDetailRs("studentBirthday")%>>
				 <input class="submit" name="Button" onclick="seltime('studentBirthday');" style="width:30px" type="button" value="选择">
			 </td>
		 </tr>
		 <tr>
		   <td>政治面貌:</td>
			 <td>
			   <select name="studentState">
				   <option value='团员' <% if studentDetailRs("studentState")="团员" then Response.Write "selected" end if%>>团员</option>
					 <option value='党员' <% if studentDetailRs("studentState")="党员" then Response.Write "selected" end if%>>党员</option>
					 <option value='老百姓' <% if studentDetailRs("studentState")="老百姓" then Response.Write "selected" end if%>>老百姓</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td>登陆密码:</td>
			 <td><input type=text name=studentPassword size=20><font color=red>如果要为该学生重新设置密码请在此输入</font></td>
		 </tr>
		 
		 <tr>
			  <td>新照片路径:</td>
			  <td><input type="text" name=photoAddress size=20 readonly="true">*请在下面上传照片,程序会自动生成路径</td>
			</tr>
			<tr> 
       <td>新照片上传：</td>
       <td bgcolor="#F5F5F5" height="30" align="center" width="79%">
		     <iframe marginwidth=0 marginheight=0  frameborder=0 scrolling=no src='upload.asp' width=450 height=30></iframe> 
       </td>
      </tr>
		  <tr>
		    <td>家庭地址:</td>
			  <td><input type=text name=studentAddress size=50 value=<%=studentDetailRs("studentAddress")%>></td>
		  </tr>
		  <tr>
		    <td width=100 align="right">附加信息:</td>
		    <td><textarea cols=40 rows=5 name=studentMemo><%=studentDetailRs("studentMemo")%></textarea></td>
		  </tr>
      <tr bgcolor="#ffffff">
        <td height="30" colspan="4" align="center">
		      <input name="submit"  type="submit" value=" 确认更新 ">
		    </td>
      </tr>
	  </table>
  </form>
</BODY>
</HTML>
