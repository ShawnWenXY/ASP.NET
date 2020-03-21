<!--BEGIN HEADER MODULE-->
<script>
	var linksrcTop;
	var linksrcLeft;
	var linksrcH;
	var linksrcW;
	var isShow = false;
	var mArray;
	var mArrayURL;

	function showMenu(menu,tableW)
	{
		switch (menu)
		{
			case '1' :mArray = new Array('学生信息添加','学生信息维护');mArrayURL = new Array('studentInfoAdd.asp','studentInfoManage.asp');break;
			case '2' :mArray = new Array('教师信息添加','教师信息维护');mArrayURL = new Array('teacherInfoAdd.asp','teacherInfoManage.asp');break;
			case '3' :mArray = new Array('添加专业信息','专业信息维护','添加班级信息','班级信息维护');mArrayURL = new Array('specialFieldInfoAdd.asp','specialFieldInfoManage.asp','classInfoAdd.asp','classInfoManage.asp');break;
			case '4' :mArray = new Array('添加必修课程','必修课程管理','添加公选课程','公选课程管理');mArrayURL = new Array('classCourseInfoAdd.asp','classCourseInfoManage.asp','publicCourseInfoAdd.asp','publicCourseInfoManage.asp');break;
			case '5' :mArray = new Array('班级课程排课','选修课程排课');mArrayURL = new Array('classCourseTeachMakeFirst.asp','publicCourseTeachMakeFirst.asp');break;
			
			case '6' :mArray = new Array('学期信息管理','修改登陆密码','系统参数设置','关于系统');mArrayURL = new Array('termInfoManage.asp','passwordSet.asp','systemConfig.asp','../System/systemInfo.asp');break;
			case '7' :mArray = new Array('添加学生成绩','成绩信息查询');mArrayURL = new Array('../Teacher/scoreInfoAdd.asp','../Teacher/scoreInfoQuery.asp');break;
		
		}
		
		if( mArray !=null )
		{
			var linksrc =window.event.srcElement;
			linksrcTop = linksrc.offsetTop;
			linksrcLeft = linksrc.offsetLeft;
			linksrcH = linksrc.offsetHeight;
			linksrcW = linksrc.offsetWidth;
			linksrcParent = linksrc.offsetParent;
			while(linksrcParent.tagName.toUpperCase()!="BODY")
			{
				linksrcTop += parseInt(linksrcParent.offsetTop);
				linksrcLeft += parseInt(linksrcParent.offsetLeft);
				linksrcParent = linksrcParent.offsetParent;
			}

			//linksrcTop += document.body.scrollTop;
			//linksrcLeft += document.body.scrollLeft;

			mdiv = document.getElementById("menuDiv");
			mdiv.style.left = linksrcLeft + 8;
			mdiv.style.top = linksrcTop + linksrcH; 

			text ="<table width='96' border='0' cellpadding='3' cellspacing='0' bgcolor='e8e8e8' class='tbl'>";
			for(var i=0;i<mArray.length;i++)
			{
				text += "<tr>";
				if(i==mArray.length-1)
				{
					text += "<td align='center'>";
				}
				else
				{
					text += "<td align='center' class='bottom'>";
				}
				text+="<a href='" + mArrayURL[i] + "' target='ContentFrame'>" + mArray[i];
				text+="</a></td></tr>";
			}
			text +="</table>"; 
			mdiv.innerHTML = text;
			mdiv.style.visibility="visible";
			isShow = true;
		}
	}

	function hiddenMenu()
	{
		if (isShow == true)
		{
			var mx = document.body.scrollLeft + window.event.clientX;
			var my = document.body.scrollTop + window.event.clientY;
			mdiv=document.all.menuDiv;
			mdivT = parseInt(mdiv.style.top);
			mdivL = parseInt(mdiv.style.left);
			
			mdivH = parseInt(mdiv.offsetHeight);
			mdivW = parseInt(mdiv.offsetWidth);
			
			if(mx < mdivL || mx > mdivL + mdivW || my < linksrcTop || my > linksrcTop + linksrcH + mdivH)
			{
				mdiv.style.visibility = "hidden";isShow = false;
			}
										
		}
	}	    
</script>
<!--   菜单层   -->
<div id="menuDiv" style="VISIBILITY: hidden; POSITION: absolute">
</div>
<FONT face="宋体"></FONT><FONT face="宋体"></FONT><FONT face="宋体"></FONT><FONT face="宋体">
</FONT><FONT face="宋体"></FONT><FONT face="宋体"></FONT><FONT face="宋体"></FONT><FONT face="宋体">
</FONT><FONT face="宋体"></FONT><FONT face="宋体"></FONT><FONT face="宋体"></FONT><FONT face="宋体">
</FONT><FONT face="宋体"></FONT><FONT face="宋体"></FONT><FONT face="宋体"></FONT><FONT face="宋体">
</FONT>
<!--   /菜单层   -->
<table class="tbl" id="Table2" cellSpacing="0" cellPadding="0" width="760" border="0">
	<tr>
		<td background="../images/titl_bg.jpg" height="20" align="right">
			<!--   导航超链接   --> &nbsp;&nbsp;
			<span onmouseover="showMenu('1',100)" style="COLOR: black;"><font color="#ffffff">学生信息管理</font></span>&nbsp;<font color="#ffffff">┊</font>&nbsp; 
			<span onmouseover="showMenu('2',100)" style="COLOR: black;"><font color="#ffffff">教师信息管理</font></span>&nbsp;<font color="#ffffff">┊</font>&nbsp;
			<span onmouseover="showMenu('3',100)" style="COLOR: black;"><font color="#ffffff">班级信息管理</font></span>&nbsp;<font color="#ffffff">┊</font>&nbsp;
			<span onmouseover="showMenu('4',100)" style="COLOR: black;"><font color="#ffffff">课程信息管理</font></span>&nbsp;<font color="#ffffff">┊</font>&nbsp;
			<span onmouseover="showMenu('5',100)" style="COLOR: black;"><font color="#ffffff">排课信息管理</font></span>&nbsp;<font color="#ffffff">┊</font>&nbsp; 
			
			<span onmouseover="showMenu('6',100)" style="COLOR: black;"><font color="#ffffff">系统管理</font></span>&nbsp;<font color="#ffffff">┊</font>&nbsp; 
			<span onmouseover="showMenu('7',100)" style="COLOR: black;"><font color="#ffffff">成绩信息管理</font></span>&nbsp;<font color="#ffffff">┊</font>&nbsp;

<A style="COLOR: black" href='../Logout.asp'><font color="#ffffff">退 出</font></A> &nbsp;&nbsp;
<!--   /导航主菜单   --></td>
	</tr>
	
</table>