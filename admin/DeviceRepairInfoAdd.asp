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
	'如果要进行设备维修信息的登记
	if Request("submit") <> "" then
	  if Request("deviceName") = "" then
	    errMessage = "请选择设备对象"
	  end if
	  sqlString = "select * from [deviceRepairInfo]"
	  set deviceRepairInfoRs = Server.CreateObject("ADODB.RecordSet")
	  deviceRepairInfoRs.Open sqlString,conn,1,3
	  deviceRepairInfoRs.AddNew
	  deviceRepairInfoRs("deviceName") = Request("deviceName")
	  deviceRepairInfoRs("deviceTypeId") = Request("deviceTypeId")
	  deviceRepairInfoRs("repairPlace") = Request("repairPlace")
	  deviceRepairInfoRs("repairMan") = Request("repairMan")
	  if Request("repairMoney") <> "" then
	    deviceRepairInfoRs("repairMoney") = CSng(Request("repairMoney"))
		else
		  deviceRepairInfoRs("repairMoney") = 0.0
		end if
		deviceRepairInfoRs("errorReason") = Request("errorReason")
		if Request("repairDate") <> "" then
		  deviceRepairInfoRs("repairDate") = CDate(Request("repairDate"))
		else
		  deviceRepairInfoRs("repairDate") = CDate("1900-1-1")
		end if
		deviceRepairInfoRs("repairMemo") = Request("repairMemo")
		deviceRepairInfoRs.Update
		Response.Write "<script>alert('设备维修信息登记成功!');location.href='DeviceRepairInfoList.asp';</script>"
	end if
%>
<HEAD>
	<Title>设备维修信息添加</Title>
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
 <form method="post" name="frmAnnounce" runat="server">
	 <table width=700 border=0 cellpadding=0 cellspacing=0 align="center">
		 <tr style="color:blue;font-size:14px;">
	     <td style="height: 14px" colspan=2 align="center">
		      <img src="../images/ADD.gif" width=14px height=14px>设备信息管理--&gt;设备维修信息添加
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
		   <td width=100px align="right">维修人:</td>
			 <td>
			   <input type=text name=repairMan size=20>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">维修地点:</td>
			 <td>
			   <input type=text name=repairPlace size=40>
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">维修金额:</td>
			 <td>
			   <input type=text name=repairMoney size=5>元
			 </td>
		 </tr>
		 <tr>
		   <td width=100px align="right">错误原因:</td>
			 <td><input type=text name=errReason size=40></td>
		 </tr>
		 <tr>
		   <td width=100px align="right">维修日期:</td>
			 <td>
			   <input type=text name=repairDate width=20px>
				 <input class="submit" name="Button" onclick="seltime('repairDate');" style="width:30px" type="button" value="选择">
			 </td>
		 </tr>
		 <tr>
		    <td width=100 align="right">附加信息:</td>
		    <td><textarea cols=40 rows=5 name=repairMemo></textarea></td>
		  </tr>
		 <tr bgcolor="#ffffff">
        <td height="30" colspan="2" align="center">
		      <input name="submit"  type="submit" value="维修登记"> &nbsp;
				  <input type="reset" value="取消"></td>
      </tr>
	 </table>
 </form>
</body>
</html>