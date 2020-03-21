<!-- #include virtual="DataBase/conn.asp" -->
<!--
/*系统配置信息表*/
CREATE TABLE [dbo].[config] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,														/*信息记录编号*/
	[systemName] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,		/*系统名称*/
	[author] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,				/*作者姓名*/
	[email] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,					/*作者Email*/
	[qq] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,						/*作者qq*/
	[phone] [varchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,					/*作者电话*/
	[pageSize] [int] NULL ,																					/*每页显示记录条数*/
	[canSelect] [int] NULL ,																				/*是否开放选课*/
	[termId] [int] NULL ,																						/*选课所在年度和学期*/
	[selectStartTime] [datetime] NULL ,															/*选课开始时间*/
	[selectEndTime] [datetime] NULL 																/*选课结束时间*/
) ON [PRIMARY]
-->
<%
  '从系统数据库中读取系统的设置
  dim systemName,author,email,qq,phone,pageSize,canSelect,termId,selectStartTime,selectEndTime
  set rsConfig=server.CreateObject("adodb.recordset")
	sqlString="select * from config where id = 1"
	rsConfig.open sqlString,conn,1,1
	'将读取的系统信息保存在全局变量中
	if not rsConfig.EOF then
	  systemName = rsConfig.Fields("systemName")
	  author = rsConfig.Fields("author")
	  email = rsConfig.Fields("email")
	  qq = rsConfig.Fields("qq")
	  phone = rsConfig.Fields("phone")
	  pageSize = CInt(rsConfig.Fields("pageSize"))
	  canSelect = CInt(rsConfig.Fields("canSelect"))
	  termId = rsConfig.Fields("termId")
	  selectStartTime = CDate(rsConfig.Fields("selectStartTime"))
	  selectEndTime = CDate(rsConfig.Fields("selectEndTime"))
	end if
	rsConfig.Close
%>
