<!--#include file="../DataBase/conn.asp"-->
<!--#include file="../System/config.asp"-->
<%
  '判断学生是否已经登陆系统
  if session("studentNumber") = "" then
    Response.Redirect "../login.asp"
	end if
%>
<HTML>
	<HEAD>
		<title><%=systemName%></title>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<LINK href="../css/style.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body onmousemove="hiddenMenu();">
		<form id="Form1" method="post" runat="server">
			<table align="center" cellSpacing="0" cellPadding="0" width="760" border="0" ID="Table5">
				<TBODY>
					<tr>
						<td>
							<!--   /菜单层   -->
							<table class="tbl" id="Table2" cellSpacing="0" cellPadding="0" width="760" border="0">
								<tr>
									<td colspan=3>
										<!--   导航超链接   -->
											<!--#include file="head.inc"-->
										<!--   /导航主菜单   --></td>
								</tr>
								<tr width=100%>
									<td width=45%>&nbsp;学生<font color=blue><%=session("studentNumber")%></font>,你好，你的登陆时间是:<%=Now%></td><td width=35%></td><td width=20%>&nbsp;</td>
								</tr>
							</table>
							<!-- END PAGE HEADER MODULE -->
							<!--   内容层   -->
							<table class="lrb" align="center" cellSpacing="0" cellPadding="0" width="760" border="0"
								ID="Table3">
								<tr>
									<td bgcolor="#d6ebff" style="height: 400px">
									<iframe style="height: 500px; width: 760px;" frameborder="0"  name="ContentFrame" scrolling="auto" src="../System/systemInfo.asp" width="760"></IFRAME>
									</td>
								</tr>
							</table>
							<!--   /内容层   -->
						</td>
					</tr>
					<tr><td>
					  <table align=center width=100% cellspacing=0 cellpadding=0>
	            <tr>
					     <td align=center bgcolor="#00ff33" width=100%>程序设计 by <%=author%>;   QQ: <%=qq%>  手机:<%=phone%> Email:<a href="mailto:<%=email%>"><%=email%></a></td>
	            </tr>
            </table>
				  </td></tr>
				</TBODY>
			</table>
		</form>
	</body>
</HTML>



