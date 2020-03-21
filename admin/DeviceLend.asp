<!--#include virtual="/DataBase/conn.asp"-->
<!--#include virtual="/System/function.asp"-->
<%
  '保存错误信息
  dim errMessage
  errMessage = ""
  '如果管理员还没有登陆
  if session("adminUsername")="" then
    Response.Write "<script>top.location.href='../login.asp';</script>"
	end if
	'如果要进行设备的外借
	if Request("submit") <> "" then
	  if Request("deviceName") = "" then
	    errMessage = "你没有选择设备对象!"
		end if
		if Request("studentNumber") = "" then
		  errMessage = "你没有填写学号"
		end if
		sqlString = "select * from [studentInfo] where studentNumber='" & Request("studentNumber") & "'"
		set studentInfoRs = Server.CreateObject("ADODB.RecordSet")
		studentInfoRs.Open sqlString,conn,1,1
		if studentInfoRs.EOF then
		  errMessage = "对不起,你输入的学号不存在"
		end if
		studentInfoRs.Close
		if Request("deviceName") <> "" then
		  '查询该设备库存量是否大于0
		  sqlString = "select * from [deviceStockInfo] where deviceName='" & Request("deviceName") & "'"
		  set deviceStockInfoRs = Server.CreateObject("ADODB.RecordSet")
		  deviceStockInfoRs.Open sqlString,conn,1,1
		  if CInt(deviceStockInfoRs("deviceStock")) <= 0 then
		    errMessage = "该设备没有库存了"
		  end if
		  deviceStockInfoRs.Close
		end if
		'如果验证没有错误则进行新的设备使用信息的加入
		if errMessage = "" then
		  sqlString = "select * from [deviceUseInfo]"
		  set deviceUseInfoRs = Server.CreateObject("ADODB.RecordSet")
		  deviceUseInfoRs.Open sqlString,conn,1,3
		  deviceUseInfoRs.AddNew
		  deviceUseInfoRs("deviceName") = Request("deviceName")
		  deviceUseInfoRs("deviceTypeId") = Request("deviceTypeId")
		  deviceUseInfoRs("studentNumber") = Request("studentNumber")
		  deviceUseInfoRs("useBeginTime") = Now
		  deviceUseInfoRs("isReturn") = 0
		  deviceUseInfoRs.Update
		  deviceUseInfoRs.Close
		  '将该设备的库存量减1
		  sqlString = "update [deviceStockInfo] set deviceStock = deviceStock - 1 where deviceName='" & Request("deviceName") & "'"
		  conn.Execute(sqlString)
		  Response.Write "<script>alert('设备使用登记完成!');</script>"
		else
		  Response.Write "<script>alert('" & errMessage & "');</script>"
		end if
	end if
%>
<HEAD>
	<Title>设备外借</Title>
	<META http-equiv="Content-Type" content="text/html; charset=gb2312">
	<LINK href="../css/style.css" type="text/css" rel="stylesheet">
	<script language=javascript>
	function seltime(inputName)
	{
	  window.open('seltime.asp?InputName='+inputName+'','','width=250,height=220,left=360,top=250,scrollbars=yes');  
	}
	var deviceType_deviceName = new Array();
	//初始化所有(设备类型－设备名称信息)信息记录数组
  function initArray() {
  <%
    sql = "select * from deviceStockInfo" 
	  set deviceStockInfoRs = conn.Execute(sql)
	  i = 0
	  while not deviceStockInfoRs.eof
	    Response.Write "deviceType_deviceName[" & i & "]='" & deviceStockInfoRs("deviceTypeId") & ":" & deviceStockInfoRs("deviceName") & "';" & vbCrLf
		  i = i + 1
		  deviceStockInfoRs.MoveNext
	  wend
  %>
 }
 //当选择不同的设备类别时显示该类别下的设备信息
function changeDeviceType() {
  var searchDeviceTypeId; //要搜索的设备类型
  var eachDeviceTypeId; //每个记录的设备类型
  var eachDeviceName; //记录每个设备的名称
  var indexOfSplit; // :号分割符号的位置
  var innerHtmlText;
  var oOption; 
  var index;
  innerHtmlText = "";
  searchDeviceTypeId = document.all.deviceTypeId.value;
  initArray(); //初始化信息数组
  index = document.all.deviceName.length
  for(;index>0;index--) {
    document.all.deviceName.remove(index);
  }
  for(var i=0;i<deviceType_deviceName.length;i++) {
    indexOfSplit = deviceType_deviceName[i].indexOf(":"); //得到:号分割符号的位置
	  eachDeviceTypeId = deviceType_deviceName[i].substr(0,indexOfSplit); //取得当前记录的设备类型编号
	  if(searchDeviceTypeId == eachDeviceTypeId) { //如果设备类型一样就把设备信息取出来加入设备信息下拉框中
	    eachDeviceName = deviceType_deviceName[i].substr(indexOfSplit+1);
		   oOption = document.createElement("OPTION");
		   document.all.deviceName.options.add(oOption);
			  oOption.innerText = eachDeviceName;
       oOption.value = eachDeviceName;
	  }
  }
}
	</script>
</HEAD>
<BODY>
<br>
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=500 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>设备信息管理--&gt;设备借用
			 </td>
	   </tr>
		 <tr>
		   <td width=100px align="right">设备类型:</td>
			 <td>
			   <select name=deviceTypeId onchange="changeDeviceType();">
				  <option value=''>请选择设备类型</option>
				   <%
					   sqlString = "select * from [deviceTypeInfo]"
						 set deviceInfoRs = Server.CreateObject("ADODB.RecordSet")
						 deviceInfoRs.Open sqlString,conn,1,1
						 while not deviceInfoRs.EOF
						   Response.Write "<option value='" & deviceInfoRs("deviceTypeId") & "'>" & deviceInfoRs("deviceTypeName") & "</option>"
							deviceInfoRs.MoveNext
						 wend
						 deviceInfoRs.Close
					 %>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td width=100 align="right">设备名称:</td>
			 <td>
			   <select name=deviceName>
				   <option value="">请选择设备</option>
				 </select>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">学号:</td>
			 <td>
			   <input type=text name=studentNumber size=20>
			 </td>
		 </tr>
		 <tr bgcolor="#ffffff">
        <td height="30" colspan="2" align="center">
		      <input name="submit"  type="submit" value="借用登记"> &nbsp;
				  <input type="reset" value="取消"></td>
      </tr>
	 </table>
 </form>
</body>
</html>