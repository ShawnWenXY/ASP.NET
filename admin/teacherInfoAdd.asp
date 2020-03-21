<!--#include virtual="/Database/conn.asp"-->
<!--#include virtual="/System/md5.asp"--> 
<%
  'errMessage保存错误信息
  dim errMessage
  errMessage = ""
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'如果管理员添加了新的学生信息并提交
	if Request("submit") <> "" then
	  '如果教师职工编号没有输入
	  if Request("teacherNumber") = "" then
	    errMessage = "请输入教师的职工编号!"
	  end if
	  '如果教师的登陆密码没有输入
	  if Request("teacherPassword") = "" then
	    errMessage = "请输入教师的登陆密码!"
		end if
	  if errMessage <> "" then
	    Response.Write "<script>alert('" & errMessage & "');</script>"
		else
	    '将教师个人信息加入的数据库中
		  set teacherInfoRs = Server.CreateObject("ADODB.RecordSet")
	    sqlString = "select * from [teacherInfo]"
	    teacherInfoRs.Open sqlString,conn,1,3
	    teacherInfoRs.AddNew
		  teacherInfoRs("teacherNumber") = Trim(Request("teacherNumber"))
		  teacherInfoRs("teacherName") = Trim(Request("teacherName"))
		  teacherInfoRs("teacherPassword") = md5(Trim(Request("teacherPassword")))
		  teacherInfoRs("teacherSex") = Trim(Request("teacherSex"))
		  '如果上传了教师的图片
		  if Request("photoAddress") <> "" then
		    teacherInfoRs("teacherPhoto") = Trim(Request("teacherPhoto"))
		  end if
		  '如果选择了教师的生日
		  if Request("teacherBirday") <> "" then
		    teacherInfoRs("teacherBirthday") = CDate(Request("teacherBirthday"))
		  else
		    teacherInfoRs("teacherBirthday") = CDate("1900-1-1")
			end if
		  '如果选择了教师的入校时间
		  if Request("teacherArriveTime") <> "" then
		    teacherInfoRs("teacherArriveTime") = CDate(Request("teacherArriveTime"))
		  else
		    teacherInfoRs("teacherArriveTime") = CDate("1900-1-1")
			end if
			teacherInfoRs("teacherCardNumber") = Trim(Request("teacherCardNumber"))
			teacherInfoRs("teacherAddress") = Trim(Request("teacherAddress"))
			teacherInfoRs("teacherPhone") = Trim(Request("teacherPhone"))
			teacherInfoRs("teacherMemo") = Trim(Request("teacherMemo"))
			teacherInfoRs.Update
			teacherInfoRs.Close
			Response.Write "<script>alert('教师信息添加成功!');</script>"
	  end if
	end if
%>

<HTML>
<HEAD>
	<Title>新教师信息添加</Title>
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
		      <img src="../images/ADD.gif" width=14px height=14px>教师信息管理--&gt;教师信息添加
			 </td>
	   </tr><br>
	   <tr>
	     <td style="height: 26px">
		     教师职工编号:</td><td><input type=text name=teacherNumber size=20></td>
			 </td>
		 </tr>
		 <tr>
		  <td>教师姓名:</td><td><input type=text name=teacherName size=20></td>
		 </tr>
		 <tr>
		   <td>性别:</td>
			 <td>
			   <select name=teacherSex>
				   <option value='男'>男</option>
					 <option value='女'>女</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td>教师生日:</td>
			 <td>
			   <input type=text name=teacherBirthday width=60px>
				 <input class="submit" name="Button" onclick="seltime('teacherBirthday');" style="width:30px" type="button" value="选择">
			 </td>
		 </tr>
		 <tr>
		   <td>入校时间:</td>
			 <td>
			   <input type=text name=teacherArriveTime width=60px>
				 <input class="submit" name="Button" onclick="seltime('teacherArriveTime');" style="width:30px" type="button" value="选择">
			 </td>
		 </tr>
		 <tr>
		   <td>登陆密码:</td>
			 <td><input type=text name=teacherPassword size=20></td>
		 </tr>
		 <tr>
		   <td>教师电话:</td>
			 <td><input type=text name=teacherPhone size=20></td>
		 </tr>
		 <tr>
		    <td>身份证号:</td>
			  <td><input type=text name=teacherCardNumber size=40></td>
		  </tr>
		 <tr>
		    <td>家庭地址:</td>
			  <td><input type=text name=teacherAddress size=50></td>
		  </tr>
		 <tr>
			  <td>照片路径:</td>
			  <td><input type="text" name=photoAddress size=20 readonly="true">*请在下面上传照片,程序会自动生成路径</td>
			</tr>
			<tr> 
       <td>照片上传：</td>
       <td bgcolor="#F5F5F5" height="30" align="center" width="79%">
		     <iframe marginwidth=0 marginheight=0  frameborder=0 scrolling=no src='upload.asp' width=450 height=30></iframe> 
       </td>
      </tr>
		  
		  <tr>
		    <td width=100 align="right">附加信息:</td>
		    <td><textarea cols=40 rows=4 name=studentMemo></textarea></td>
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
