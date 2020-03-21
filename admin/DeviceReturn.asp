<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'如果要求归还某个设备
	if Request("action") = "return" then
	  useInfoId = Request("useInfoId")
	  '首先更新该条借用信息(修改归还时间和归还标志变量)
	  sqlString = "select * from [deviceUseInfo] where useInfoId=" & useInfoId
	  set rs = server.CreateObject("ADODB.RecordSet")
	  rs.Open sqlString,conn,1,3
	  rs("useEndTime") = Now
	  rs("isReturn") = 1
	  rs.Update
	  '然后增加对应设备的库存
	  sqlString = "update [deviceStockInfo] set deviceStock = deviceStock + 1 where deviceName='" & rs("deviceName") & "'"
    conn.Execute(sqlString)
	  Response.Write "<script>alert('设备归还成功!');</script>"
  end if
  
	'取得各个查询的参数信息
	studentNumber = Request("studentNumber")
	startTime = Request("startTime")
	endTime = Request("endTime")
	sqlString = "select * from [deviceUseInfo] where isReturn=0"
	'根据各个参数的信息构造查询sql
	if studentNumber <> "" then
	  sqlString = sqlString & " and studentNumber like '%" & studentNumber & "%'"
	end if
	if startTime <> "" then
	  sqlString = sqlString & " and useBeginTime > '" & CDate(startTime) & "'"
	end if
	if endTime <> "" then
	  sqlString = sqlString & " and useBeginTime < '" & CDate(endTime) & "'"
	end if
	set deviceUseInfoRs = Server.CreateObject("ADODB.RecordSet")
	deviceUseInfoRs.Open sqlString,conn,1,1
	
%>
<html>
<HEAD>
	<Title>设备归还</Title>
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
<br>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=650 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=6 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>设备信息管理--&gt;设备归还
			 </td>
	   </tr>
		 <tr>
		   <td colspan=6 height=10>
			   学号:<input type="text" name=studentNumber size=15 value=<%=Request("studentNumber")%>>
				 开始时间:<input type=text name=startTime size=10 value='<%=Request("startTime")%>'>
				 <input class="submit" name="Button" onclick="seltime('startTime');" style="width:30px" type="button" value="选择">
			   结束时间:
				 <input type=text name=endTime size=10 value='<%=Request("endTime")%>'>
				 <input class="submit" name="Button" onclick="seltime('endTime');" style="width:30px" type="button" value="选择">
			   &nbsp;&nbsp
				<input type="submit" value=" 检索 " class="button1">
		   </td>
		 </tr>
		 <tr>
		   <td>设备名称</td>
			 <td>设备类型</td>
			 <td>学号</td>
			 <td>姓名</td>
			 <td>借用时间</td>
			 <td>归还</td>
		 </tr>
		 <%
			 while not deviceUseInfoRs.EOF
			   Response.Write "<tr><td>" & deviceUseInfoRs("deviceName") & "</td><td>" & GetDeviceTypeNameById(deviceUseInfoRs("deviceTypeId")) & "</td><td>" & deviceUseInfoRs("studentNumber") & "</td><td>" & GetStudentNameByNumber(deviceUseInfoRs("studentNumber")) & "</td><td>" & deviceUseInfoRs("useBeginTime") & "</td><td><a href='DeviceReturn.asp?action=return&useInfoId=" & deviceUseInfoRs("useInfoId") & "' onclick='return confirm(" & """" & "确认归还吗?" & """" & ");'>归还</a></td></tr>"
				 deviceUseInfoRs.MoveNext
			wend
		 %>
	 </table>
 </form>
</body>
</html>