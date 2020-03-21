<!--#include file="../database/conn.asp"-->
<!--#include file="../system/function.asp"-->
<%
  '如果学生还没有登陆
  if session("studentnumber")="" then
    response.write "<script>top.location.href='../login.asp';</script>"
	end if
	dim classname(12),classfieldname(12)
	classname(1) = "上午第一节"
	classname(2) = "上午第二节"
	classname(3) = "上午第三节"
	classname(4) = "上午第四节"
	classname(5) = "上午第五节"
	classname(6) = "下午第一节"
	classname(7) = "下午第二节"
	classname(8) = "下午第三节"
	classname(9) = "下午第四节"
	classname(10) = "晚上第一节"
	classname(11) = "晚上第二节"
	classname(12) = "晚上第三节"
	classfieldname(1) = "morningone"
	classfieldname(2) = "morningtwo"
	classfieldname(3) = "morningthree"
	classfieldname(4) = "morningfour"
	classfieldname(5) = "morningfive"
	classfieldname(6) = "afternoonone"
	classfieldname(7) = "afternoontwo"
	classfieldname(8) = "afternoonthree"
	classfieldname(9) = "afternoonfour"
	classfieldname(10) = "eveningone"
	classfieldname(11) = "eveningtwo"
	classfieldname(12) = "eveningthree"
%>
<html>
<head>
	<title>学生选课</title>
	<meta http-equiv="content-type" content="text/html; charset=gb2312">
	<link href="../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
	<table width=700px align="center" border="1" cellspacing="0" cellpadding="2" bordercolorlight='#000000' bordercolordark='#ffffff'>
    <form name="form1" method="post" action="coursetablemake.asp">
      <tr style="color:blue;font-size:14px;">
	      <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/list.gif" width=14px height=14px>选课信息管理--&gt;生成课表
			 </td>
	   </tr>
		 <tr>
		   <td>请选择学期:</td>
			 <td>
			   <select name=termid>
				    <option value="">请选择</option>
					  <%
					    dim sqlstring
						  sqlstring = "select * from [terminfo]"
						  set terminfors = server.createobject("adodb.recordset")
						  terminfors.open sqlstring,conn,1,1
						  while not terminfors.eof
						    response.write "<option value='" & terminfors("termid") & "'>" & terminfors("termbeginyear") & "-" & terminfors("termendyear") & "年" & terminfors("termupordown") & "</option>"
							  terminfors.movenext
						  wend
					  %>
				  </select>&nbsp;<input type="submit" name="submit" value="生成课表">
				</td>
		 </tr>
	  <tr>
	    <td height="30">课表结果:</td>
			<td>
				<table width=100% border=1 cellspacing=0 bordercolor="green">
				  <%
						if request("termid") <> "" then
						   response.write "<tr><td colspan=8 align=center>" & gettermnamebyid(request("termid")) & " 课表</td></tr>"
						end if
				  %>
				  <tr><td width="12.5%">课表</td><td width="12.5%">星期一</td><td width="12.5%">星期二</td><td width="12.5%">星期三</td><td width="12.5%">星期四</td><td width="12.5%">星期五</td><td width="12.5%">星期六</td><td width="12.5%">星期日</td></tr>
			  <%
			      for i = 1 to 12 
						   response.write "<tr><td>" & classname(i) & "</td>"
					     for j = 1 to 7
						      response.write "<td align=center id='class" & i & j & "'>&nbsp;</td>"
						   next
							 response.write "</tr>"
						next
			  %>
			  </table>
			</td>
	 </tr>
	</form>
 </table>
 <script language="javascript">
 <%
   '如果选择了要生成某个学期的课表
   if request("termid") <> "" then
	   '得到该学生的班级编号
	   classnumber = getclassnumberbystudentnumber(session("studentnumber"))
		 '得到该学生的专业编号
		 specialfieldnumber = getspecialfieldnumberbystudentnumber(session("studentnumber"))
		 '查询该学期该班级的必修课程上课信息
		 sqlstring = "select * from [classcourseteachview] where classnumber='" & classnumber & "' and termid=" & request("termid")
		 set classcourseteachrs = server.createobject("adodb.recordset")
		 classcourseteachrs.open sqlstring,conn,1,1
		 while not classcourseteachrs.eof
		   for i = 1 to 12
			   if classcourseteachrs(classfieldname(i)) = true then
				   '如果某节课存在上课信息就输出到对应的位置
				   response.write "document.all.class" & i & classcourseteachrs("teachday") & ".innerhtml=" & """" & classcourseteachrs("coursename") & "(" & classcourseteachrs("teachclassroom") & ")" & """" & ";" & vbcrlf
				 end if
			 next
			 classcourseteachrs.movenext
		 wend
		 classcourseteachrs.close
		 '查询该学期该学生选修课程上课信息
		 sqlstring = "select * from [publiccourseteachview] where studentnumber='" & session("studentnumber") & "' and termid=" & request("termid")
		 set publiccourseteachrs = server.createobject("adodb.recordset")
		 publiccourseteachrs.open sqlstring,conn,1,1
		 while not publiccourseteachrs.eof
			 for i = 1 to 12
			   if publiccourseteachrs(classfieldname(i)) = true then
				   '如果某节课存在上课信息就输出到对应的位置
				   response.write "document.all.class" & i & publiccourseteachrs("teachday") & ".innerhtml+=" & """" & publiccourseteachrs("coursename") & "(" & publiccourseteachrs("teachclassroom") & ")" & """" & ";" & vbcrlf
				 end if
			 next
			 publiccourseteachrs.movenext
		 wend
	end if
 %>
 </script>
 </body>
 </html>