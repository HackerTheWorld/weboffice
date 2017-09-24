<%@ page contentType="text/html; charset=gb2312" %>
<%@ page import="java.io.*,java.text.*,java.util.*,java.sql.*,javax.servlet.*,javax.servlet.http.*" %>
<html>
<head>
<%
request.setCharacterEncoding("GB2312");
String strHeight = "100%";
String strTitlebar = "1";
String strToolbar = "1";
String strHost = request.getServerName();
String strURL = request.getRequestURI();
strURL = strURL.substring(0,strURL.lastIndexOf("/")+1);
String strDefaultRoot = "http://" + strHost + strURL;
String strOpenUrl = "http://" + strHost  + strURL + "writefile.jsp";
String strPdfSaveUrl ="http://" + strHost + strURL + "pdf.jsp";
String strHTMLSaveUrl ="http://" + strHost + strURL + "html.jsp";
String strppFileSaveUrl ="http://" + strHost + strURL + "ppfile.jsp";
strURL ="http://" + strHost + strURL + "uploadedit.jsp"; 
String office =  request.getParameter("office");
String strflsid =  request.getParameter("flsid");
String strnum =  request.getParameter("num");
String strfname =  request.getParameter("fname");
String strfcreator =  request.getParameter("fcreator");
String tselect =  request.getParameter("tselect");
String strflag =  request.getParameter("flag");
String stroper =  request.getParameter("oper");

String strJsScript="";

String loadfun="";
if (office=="wps"){
	loadfun="WebOpen('1');";
}
else{
	loadfun="WebOpen('');";
}

String strid = request.getParameter("id");
String strSQL = null;

String DBDriver = "sun.jdbc.odbc.JdbcOdbcDriver"; 		   
String rootpath=application.getRealPath("/");
String ConnStr="jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ="+rootpath+"\\database\\weboffice.mdb";			
//String ConnStr = "jdbc:odbc:webofficedsn";
Connection conn = null; 
ResultSet rs = null; 
Statement stmt = null;



if (stroper.compareTo("read")==0)
{
	  strSQL = "select * from dt_document where  o_pkid="+strid;
	  try 
		{ 
			
			Class.forName(DBDriver).newInstance(); 
		} 
		catch(java.lang.ClassNotFoundException e)
		{ 
			System.err.println("Conn(): " + e.getMessage()); 
		}

	try
		{ 
		conn = DriverManager.getConnection(ConnStr); 
		stmt = conn.createStatement();
		rs = stmt.executeQuery(strSQL);
			if(rs.next())
			{
			 
			  try
				  {
					strfname = rs.getString("o_name");
					strfcreator = rs.getString("o_creator");
					strnum = rs.getString("o_number");
					strflsid = rs.getString("o_flsid");
					strflag = rs.getString("o_flag");
					if (strfname == null)strfname="";
					if (strfcreator == null)strfcreator="";
					if (strnum == null)strnum="";
					if (strflsid == null)strflsid="";
					if (strflag == null)strflag="";
				  }
				  catch(Throwable e)
				  {
					System.out.println(e.toString());
					throw new ServletException(e.toString());
				  }


			}
			stmt.close();
			rs.close();
			conn.close();
		} 
	catch(SQLException e)
		{ 
		System.err.println("QueryErr: " + e.getMessage()); 
		 
		}
	out.println("<"+"script language=javascript>flag='1"+ strflag +"';<"+"/script>");
	strOpenUrl = strOpenUrl+ "?id="+strid +"&operfile=office&file="+strflsid;
	strURL = strURL + "?oper=edit&flsid="+ strflsid + "&flag="+strflag;
	strHeight="500";
	strTitlebar = "0";
	strToolbar = "0";
}

if (stroper.compareTo("doedit")==0) 
{
		 
			 
	 strSQL = "update dt_document set o_number='"+strnum+"',o_name='"+strfname+"', o_flag="+strflag+" where o_flsid='"+strflsid+"'";
	
	 try 
	{ 
		
		Class.forName(DBDriver).newInstance();  
	} 
	catch(java.lang.ClassNotFoundException e)
	{ 
		System.err.println("Conn(): " + e.getMessage()); 
		 
	}

	try
	{ 
		conn = DriverManager.getConnection(ConnStr); 
		stmt = conn.createStatement();
		stmt.executeUpdate(strSQL);
		stmt.close();
			conn.close();

		out.println("<script language=javascript>window.opener.location.reload();window.close();</script>"); 
	} 
	catch(SQLException e)
	{ 
		System.err.println("QueryErr: " + e.getMessage()); 
		 
	} 
	
	 return; 


	
}



if (stroper.compareTo("addnew")==0)
{
	strfcreator ="admin";
	 
	strSQL = "update dt_document set o_number='"+strnum+"',o_name='"+strfname+"',o_creator='"+strfcreator+"', o_flag="+strflag+" where o_flsid='"+strflsid+"'";
		
		 try 
	{ 
		
		Class.forName(DBDriver).newInstance(); 
	} 
	catch(java.lang.ClassNotFoundException e)
	{ 
		System.err.println("Conn(): " + e.getMessage()); 
	}

	try
	{ 
		conn = DriverManager.getConnection(ConnStr); 
		stmt = conn.createStatement();
		stmt.executeUpdate(strSQL);
			stmt.close();
			
			conn.close();

		out.println( "<script language=javascript>window.opener.location.reload();window.close();</script>");
	} 
	catch(SQLException e)
	{ 
		System.err.println("QueryErr: " + e.getMessage()); 
		
	} 
 	return;
}

if (stroper.compareTo("edit")==0)
{
	  strSQL = "select * from dt_document where  o_pkid="+strid;
	  try 
		{ 
			
			Class.forName(DBDriver).newInstance(); 
		} 
		catch(java.lang.ClassNotFoundException e)
		{ 
			System.err.println("Conn(): " + e.getMessage()); 
		}

	try
		{ 
		conn = DriverManager.getConnection(ConnStr); 
		stmt = conn.createStatement();
		rs = stmt.executeQuery(strSQL);
			if(rs.next())
			{
		 
			  try
				  {
					strfname = rs.getString("o_name");
					strfcreator = rs.getString("o_creator");
					strnum = rs.getString("o_number");
					strflsid = rs.getString("o_flsid");
					strflag = rs.getString("o_flag");
					if (strfname == null)strfname="";
					if (strfcreator == null)strfcreator="";
					if (strnum == null)strnum="";
					if (strflsid == null)strflsid="";
					if (strflag == null)strflag="";
				  }
				  catch(Throwable e)
				  {
					System.out.println(e.toString());
					throw new ServletException(e.toString());
				  }


			}
			stmt.close();
			rs.close();
			conn.close();
		} 
	catch(SQLException e)
		{ 
		System.err.println("QueryErr: " + e.getMessage()); 
		 
		}
		strJsScript = strJsScript + "<"+"script language=javascript>flag='1"+ strflag +"';<"+"/script>";
	//out.println("<"+"script language=javascript>flag='1"+ strflag +"';<"+"/script>");
	strOpenUrl = strOpenUrl+ "?id="+strid +"&operfile=office&file="+strflsid;
	strPdfSaveUrl = strPdfSaveUrl+ "?flsid="+ strflsid;
	strHTMLSaveUrl = strHTMLSaveUrl+ "?flsid="+ strflsid;
	strppFileSaveUrl = strppFileSaveUrl+ "?flsid="+ strflsid;
	strURL = strURL + "?oper=edit&flsid="+ strflsid + "&flag="+strflag+"&id="+strid;
	stroper = "doedit";   
}



SimpleDateFormat formatter=new java.text.SimpleDateFormat("yyyyMMddHHmmss"); 
java.util.Date currentTime=new java.util.Date(); 
String crrtime=formatter.format(currentTime);

if (stroper.compareTo("new")==0)
{ 
		strflsid = crrtime;
		/*
		switch(Integer.parseInt(strflag))	 
		{ 
			case 2:
			strflsid = strflsid+".xls";
			break;
			case 3:
			strflsid = strflsid+".ppt";
			break;
			default:
			strflsid = strflsid+".doc";
				
		}
		*/		 
		strURL = strURL+"?oper=new&flsid="+strflsid ; 
		strnum = crrtime;		
		stroper = "addnew";
		if( tselect!=null && tselect.compareTo("0")==0 && tselect!=""){
			strJsScript = strJsScript + "<script language=javascript>flag='0';</script>";
			//	out.println( "<script language=javascript>flag='0';</script>");
				strOpenUrl = strOpenUrl + "?oper=new&id=" + tselect;
		}else
		{
			if(strflag!=""){
		 
				//out.println(  "<script language=javascript>flag='"+strflag+"';</script>");
				strJsScript = strJsScript + "<script language=javascript>flag='"+strflag+"';</script>";

				
				}
		}
		
if (strfname==null || strfname==""){ strfname="测试文档"+ crrtime;}
 if(strnum=="" || strnum ==null)strnum=crrtime;
			 
}
 

%>


<title>WebOffice办公文档控件演示版：<%=strfname%></title>
<link rel=stylesheet type=text/css href="cssjs/style.css">
<script language=javascript>
var strRoot;
var strOpenUrl;
var strURL;
var autoSave=1;
var pfile='<%=strflsid%>'; 
var flag='<%=strflag%>';
strOpenUrl = '<%=strOpenUrl%>';
strURL='<%=strURL%>';
strRoot = '<%=strDefaultRoot%>';
var strPdfSaveUrl='<%=strPdfSaveUrl%>';
var strHTMLSaveUrl='<%=strHTMLSaveUrl%>';
var strppFileSaveUrl='<%=strppFileSaveUrl%>';
</script><%=strJsScript%>
<script language=javascript src="cssjs/weboffice.js"></script>
</head>
<body topmargin=0 leftmargin=0 onload="javascript:<%=loadfun%>"><form action="WebDocEdit.jsp?oper=<%=stroper%>"  name=frm method="post" onsubmit="return WebSave()">
<table width=100% bgcolor=#cccccc cellpadding=1 cellspacing=1><input type=hidden value=<%=strflag%> name=flag><tr><td width=80 valign=top bgcolor=white align=center><input type=hidden name=oper value=<%=stroper%> id=oper>
[<b style="color:green">保存操作</b>]<br>
<input type=button  class="button" value="同步保存" style="width:80" onclick="WebHttpSave();"><br>
<input type=submit class="button"   value="异步保存" style="width:80"><br>
<input type=button class="button"   value="重调文档"  onclick="WebDocReload()"><br>[<b style="color:green">复制</b>]<br><input type=button  class="button" value="不充许复制" onclick="frm.WebOffice.IsNotCopy=1;"><br>
<input type=button class="button" value="充许复制" onclick="frm.WebOffice.IsNotCopy=0;"><br>
[<b style="color:green">自动保存</b>]<br>
<input type=button class="button" id=savbtn onclick="alert('当前设置使控件每隔30秒自动保存文档！请稍等,控件正在处理...');MyTimer();this.value='保存已开启';clsbtn.style.display='';"  value="开启自动保存" style="width:80"><input type=button id=clsbtn style="display:none;" class="button" onclick="autoSave=0;this.value='自动保存已关';savbtn.value='开启自动保存';info.innerHTML='';this.style.display='none'" value="关闭自动保存" style="width:80"><font id=info style="display:none"></font>[<b style="color:green">打开保存</b>]<br>
<input type=button class="button" value="打开方式一" onclick="var a=document.all.WebOffice.WebLoadFile('http://www.officectrl.com/weboffice/temp/file1.doc',null);"><br>
<input type=button class="button" value="打开方式二" onclick="var a=document.all.WebOffice.WebLoadFile('http://www.officectrl.com/weboffice/temp/file1.doc','Word.Document');"><br>
<input type=button class="button" value="打开方式三" onclick="document.all.WebOffice.WebOpen('http://www.officectrl.com/weboffice/temp/file1.doc');"><br>
<input type=button class="button" value="打开方式四" onclick="document.all.WebOffice.WebUrl('http://www.officectrl.com/weboffice/temp/file1.doc','Word.Document');"><br>
<input type=button class="button" value="新建文件EXCEL" onclick="var a=document.all.WebOffice.WebLoadFile('','xls');frm.flag.value=2;"><br>
<input type=button class="button" value="新建文件WORD" onclick="document.all.WebOffice.WebLoadFile('','doc');frm.flag.value=1;"><br>
<input type=button class="button" value="新建文件PPT" onclick="var a=document.all.WebOffice.WebLoadFile('','ppt');frm.flag.value=3;"><br>
<input type=button class="button"   value="保存到C:\a.doc"  onclick="WebSaveC()">
<input type=button class="button"   value="打开本地文件"  onclick="WebOpenLocal()"><br>
<input type=button class="button"   value="存为本地文件"  onclick="WebSaveLocal()"><br>
<input type=button class="button"   value="指定页保存"  onclick="document.all.WebOffice.SetPageAs('c:\\b.doc',1,0);alert('已将文档第1页的内容另存为c盘根目录下的b.doc文件');"><br>
<!--<input type=button class="button"   value="下载指定文件"  onclick="document.all.WebOffice.DownloadFile('http://www.officectrl.com/weboffice/temp/file1.doc','c:\\a.doc');alert('已将http://www.officectrl.com/weboffice/temp/file1.doc\n下载到c盘根目录下的a.doc');"><br>
--><input type=button class="button"   value="删除文件"  onclick="document.all.WebOffice.DeleteLocalFile('c:\\a.doc');alert('已将c盘根目a.doc删除');"><br>

[<b style="color:green">文档转换</b>]<br>
<%if(strflag.equals("1")){%>
<input type=button class="button"   value="保存远程PDF"  onclick="WebSaveRemotePdf();"><br>
<input type=button class="button"   value="保存远程HTML"  onclick="WebSaveRemoteHTML()" id="idWordTable"><br>
<input type=button class="button"   value="保存本地PDF"  onclick="WebSaveLocalPdf()"><br>
<input type=button class="button"   value="保存本地HTML"  onclick="WebSaveLocalHTML()"><br>
<%}%><%if(strflag.equals("2")){%><input type=button class="button"   value="保存本地PDF"   onclick="WebSaveXLSLocalPDF();"><br>
<input type=button class="button"   value="保存本地HTML"   onclick="WebSaveXLSLocalHTML();"><br>
<input type=button class="button"   value="保存远程PDF"   onclick="WebSaveXLSAsPDF();"><br>
<input type=button class="button"   value="保存远程HTML"   onclick="WebSaveXLSAsHTML();"><br>
<input type=button class="button"   value="保存远程MHT"   onclick="WebSaveXLSAsMHT();"> <br>
<%}%><%if(strflag.equals("3")){%><input type=button class="button"   value="保存本地PDF"   onclick="WebSavePPTLocalPDF();"><br>
<input type=button class="button"   value="保存本地图片"   onclick="WebSavePPTLocalJPG();"><br>
<input type=button class="button"   value="保存远程PDF"   onclick="WebSavePPTAsPDF();"><br>
<input type=button class="button"   value="保存远程HTML"   onclick="WebSavePPTAsHTML();"><br>
<%}%><input type=hidden  value="" id=field1 name=field1>
<input type=hidden  value="" id=field2 name=field2>
<input type=hidden  value="" id=field3 name=field3>
<input type=hidden  value="" id=field4 name=field4>
<input type=hidden  value="<%="http://" +  request.getServerName() + request.getRequestURI() + "?" + request.getQueryString()%>" id=field5 name=field5>
<input type=hidden  value="<%=strDefaultRoot%>" id=field6 name=field6>
[<b style="color:green">打印控制</b>]
<input type=button class="button"   value="打印预览"  onclick="WebDocPrintPreView()">
<input type=button class="button"   value="设置后打印一"  onclick="WebDocPrint()">
<input type=button class="button"   value="设置后打印二"  onclick="WebDocPrint2()">
<input type=button class="button"   value="直接打印"  onclick="WebPrintDirc()">
[<b style="color:green">窗口设置</b>]<br>
<input type=button class="button"   value="关闭标题栏"   onclick="WebTitlebar(false)"><br>
<input type=button class="button"   value="打开标题栏"   onclick="WebTitlebar(true)"><br>
<input type=button class="button"   value="关闭菜单栏"   onclick="WebMenubar(false)"><br>
<input type=button class="button"   value="打开菜单栏"   onclick="WebMenubar(true)"><br>
<input type=button class="button"   value="关闭工具栏"   onclick="WebToolbar(false)"><br>
<input type=button class="button" value="打开工具栏"   onclick="WebToolbar(true)"><br>
<!--<input type=button class="button" value="禁止文件菜单(不能新建、打开、保存等)" onclick="alert('此功能可以将控件文件菜单里的菜单项无效！');document.all.WebOffice.SetMenuDisplay(0);"><br>
<input type=button class="button" value="设置系统时间" onclick="alert('按确定后，将操作系统时间设置为：2016-09-02 12:10:10');document.all.WebOffice.SetCurrTime('2016-09-02 12:10:10');"><br>

<input type=button class="button" value="Office版本" onclick="var a=document.all.WebOffice.GetOfficeVersion('Word.Document');alert(a);"><br>-->
[<b style="color:green">文档视图</b>]<br>
<input type=button class="button"   value="普通视图"  onclick="document.all.WebOffice.ShowView(1);"><br>
<input type=button class="button"   value="大纲视图"  onclick="document.all.WebOffice.ShowView(2);"><br>
<input type=button class="button"   value="页面视图"  onclick="document.all.WebOffice.ShowView(3);"><br>
<input type=button class="button"   value="打印预览"  onclick="document.all.WebOffice.ShowView(4);"><br>
<input type=button class="button"   value="主控视图"  onclick="document.all.WebOffice.ShowView(5);"><br>
<input type=button class="button"   value="WEB视图"  onclick="document.all.WebOffice.ShowView(6);"><br>
<input type=button class="button"   value="阅读视图"  onclick="document.all.WebOffice.ShowView(7);"><br>
[<b style="color:green">文档属性</b>]<br>
<input type=button class="button" value="Application示例" onclick="var app=document.all.WebOffice.GetApplication;app.Dialogs(163);"><br>
<input type=button class="button" value="ActiveDocument示例" onclick="form.WebOffice.ActiveDocument.Application.Dialogs(129);"><br>
<input type=button class="button"   value="取得缓冲路径"  onclick="var a=document.all.WebOffice.TempFilePath;alert(a);"><br>
</td><td  nowrap bgcolor=white  valign=top width=80 align=center>[<b style="color:green">全屏显示</b>]<br>
<input type=button class="button" value="桌面全屏" class="button"   onclick="document.all.WebOffice.MenuBars=1;document.all.WebOffice.FullScreenType=1;document.all.WebOffice.WebFullScreen();"><br>
<!--<input type=button class="button" value="IE内全屏" class="button"   onclick="document.all.WebOffice.MenuBars=1;document.all.WebOffice.FullScreenType=2;document.all.WebOffice.WebFullScreen();"><br>[<b style="color:green">鉴名盖章</b>]<br>
--><input type=button class="button"   value="手写签名"  onclick="WebDocSignature()"><br>
<input type=button class="button"   value="添加财务章"  onclick="WebSignature('1')"><br>
<input type=button class="button"   value="添加合同章"  onclick="WebSignature('2')"><br>
<input type=button class="button"   value="添加行政章"  onclick="WebSignature('3')"><br> 
[<b style="color:green">书签/替换</b>]<br>
<input type=button class="button" value="替换文本" onclick="document.all.WebOffice.ReplaceText('a','b',0);"><br>
<input type=button class="button" value="添加书签" onclick="document.all.WebOffice.SetFieldValue('mark_1','test','::ADDMARK::');"><br>
<input type=button class="button" value="书签红色" onclick="document.all.WebOffice.SetFieldValue('mark_1','255','::SETCOLOR::');"><br>
<!--<input type=button class="button" value="获得书签色" onclick="var vcolor= document.all.WebOffice.SetFieldValue('mark_1','','::GETCOLOR::');alert(vcolor);"><br>-->
<input type=button class="button" value="删除书签" onclick="document.all.WebOffice.SetFieldValue('mark_1','','::DELMARK::');"><br>
<input type=button class="button" value="获得书签" onclick="document.all.WebOffice.SetFieldValue('mark_1','','::GETMARK::');"><br>
<input type=button class="button" value="获得所有书签" onclick="WebGetAllMark()"><br>
<input type=button class="button" value="设置套红SetFieldValue" onclick="document.all.WebOffice.SetFieldValue('mark_1','http://www.officectrl.com/weboffice/temp/file1.doc','::FILE::');"><br>
<input type=button class="button" value="插入图片" onclick="WebAddPic()"><br>
<input type=button class="button" value="插入浮动图片" onclick="WebAddFloatPic()"><br>
<input type=button class="button" value="EXCEL加入图片SetFieldValue" onclick="document.all.WebOffice.SetFieldValue('mark_1','test','::ADDMARK::');document.all.WebOffice.SetFieldValue('mark_1','http://www.officectrl.com/weboffice/images/weboffice.jpg','::JPG::');"><br>
<input type=button class="button"   name="insertword"  value="插入WORD文档"  onclick="document.all.WebOffice.InSertFile('http://www.officectrl.com/weboffice/temp/file1.doc',0);"><br>

[<b style="color:green">文档保护</b>]<br>
<input type=button class="button" value="文档完全保护" onclick="document.all.WebOffice.ProtectDoc(1,2,'123');"><br>
<input type=button class="button" value="保护后接受修订" onclick="document.all.WebOffice.ProtectDoc(1,0,'123');"><br>
<input type=button class="button" value="保护后充许批注" onclick="document.all.WebOffice.ProtectDoc(1,1,'123');"><br>
<input type=button class="button" value="文档解除保护" onclick="document.all.WebOffice.ProtectDoc(0,0,'123');"><br>
<input type=button class="button"   value="Excel保护部份单元格" onclick="WebSheetsLock()"><br>
 [<b style="color:green">页面控制</b>]<br>
<input type=button class="button"   value="文档页数"  onclick="WebDocumentPageCount()"><br>
<input type=button class="button"   value="插入页眉"  onclick="frm.WebOffice.ActiveDocument.ActiveWindow.ActivePane.View.SeekView=9;"><br>
<input type=button class="button"   value="插入页码"  onclick="frm.WebOffice.WebDialogs(294);"><br>
<input type=button class="button"   value="插入字符"  onclick="frm.WebOffice.WebDialogs(162);"><br>
<input type=button class="button"   value="插入目录"  onclick="frm.WebOffice.WebDialogs(171);"><br>
<input type=button class="button"   value="插入表格"  onclick="frm.WebOffice.WebDialogs(129);"><br>
[<b style="color:green">Word对话框</b>]<br>
<input type=button class="button"   value="项目符号与编号ToolsBulletsNumbers"  onclick="frm.WebOffice.WebDialogs(196);"><br>
<input type=button class="button"   value="排序文字TableSort"  onclick="frm.WebOffice.WebDialogs(199);"><br>
<input type=button class="button"   value="字数统计ToolsWordCount"  onclick="frm.WebOffice.WebDialogs(228);"><br>
<input type=button class="button"   value="取消选择ShrinkSelection"  onclick="frm.WebOffice.WebDialogs(236);"><br>
<input type=button class="button"   value="选择全部EditSelectAll"  onclick="frm.WebOffice.WebDialogs(237);"><br>
<input type=button class="button"   value="插入页码InsertPageField"  onclick="frm.WebOffice.WebDialogs(239);"><br>
<input type=button class="button"   value="插入日期InsertDateField"  onclick="frm.WebOffice.WebDialogs(240);"><br>
<input type=button class="button"   value="插入时间InsertTimeField"  onclick="frm.WebOffice.WebDialogs(241);"><br>
<input type=button class="button"   value="页面设置FilePageSetup"  onclick="frm.WebOffice.WebDialogs(178);"><br>
<input type=button class="button"   value="制表位FormatTabs"  onclick="frm.WebOffice.WebDialogs(179);"><br>
<input type=button class="button"   value="样式FormatStyle"  onclick="frm.WebOffice.WebDialogs(180);"><br>
<input type=button class="button"   value="字体格式对话框FormatDefineStyleFont"  onclick="frm.WebOffice.WebDialogs(181);"><br>
<input type=button class="button"   value="制表位FormatDefineStyleTabs"  onclick="frm.WebOffice.WebDialogs(183);"><br>
<input type=button class="button"   value="样式FormatDefineStyleFrame"  onclick="frm.WebOffice.WebDialogs(184);"><br>
<input type=button class="button"   value="段落格式对话框FormatParagraph"  onclick="frm.WebOffice.WebDialogs(175);"><br>
<input type=button class="button"   value="节的样式FormatSectionLayout"  onclick="frm.WebOffice.WebDialogs(176);"><br>
<input type=button class="button"   value="分栏FormatColumns"  onclick="frm.WebOffice.WebDialogs(177);"><br>

</td>
<td  nowrap bgcolor=white  valign=top width=80 align=center>
[<b style="color:green">模板套红</b>]<br>
<input type=button class="button"   name="Revision"  value="文档顶部套红"  onclick="document.all.WebOffice.InSertFile('http://www.officectrl.com/weboffice/temp/file1.doc',1);"><br>
<input type=button class="button"   name="Revision"  value="当前位置套红"  onclick="document.all.WebOffice.InSertFile('http://www.officectrl.com/weboffice/temp/file1.doc',0);"><br>
<input type=button class="button"   name="Revision"  value="文档底部套红"  onclick="document.all.WebOffice.InSertFile('http://www.officectrl.com/weboffice/temp/file1.doc',2);"><br>
[<b style="color:green">痕迹示例</b>]<br>
<input type=button class="button" value="设置当前用户" onclick="document.all.WebOffice.SetCurrUserName('jenny');alert('已设置当前用户为jenny');"><br>
<input type=button class="button" value="开始修订SetTrackRevisions" onclick="document.all.WebOffice.SetTrackRevisions(1);"><br>
<input type=button class="button" value="接受修订SetTrackRevisions" onclick="document.all.WebOffice.SetTrackRevisions(4);"><br>
<input type=button class="button" value="隐藏修订ShowRevisions" onclick="document.all.WebOffice.ShowRevisions(0);"><br>
<input type=button class="button" value="修订统计" onclick="document.all.WebOffice.showdialog(6);"><br>
<input type=button class="button"   value="突出显示修订ToolsRevisions"  onclick="frm.WebOffice.WebDialogs(197);"><br>
<input type=button class="button" name="Revision" value="隐藏痕迹"  onclick="ShowRevision(false)"><br>
<input type=button class="button" value="获取痕迹" onclick="ShowRevision(true)"><br>
<input type=button class="button" value="接受所有修订" onclick="WebAcceptAllRevisions()"><br>
[<b style="color:green">VBA示例</b>]<br>
<select id="template"><option value="1">发文套红<option value="2">公文套红2<option value="3">公文套红<option value="4">收文套红</select><br>
<input type=button class="button"   value="VBA选择套红"  onclick="WebAddTemplate(template.value)"><br>
<input type=button class="button"   value="VBA套红定稿"  onclick="WebTempFile();"><br>
<input type=button class="button"   value="插入本地图片"  onclick="WebOpenPicture()"><br>
<input type=button class="button"   value="插入URL图片"  onclick="WebInsertImage()"><br>
<input type=button class="button" value="替换Word内容" onclick="alert('本功能将演示控件里WORD内容中的“文”替换成“档”,每执行一次替换一处');document.all.WebOffice.ReplaceText('文','档',1);"><br>
<input type=button class="button"   value="取Word内容"   onclick="WebGetWordContent()"><br>
<input type=button class="button"   value="写Word内容"  onclick="WebSetWordContent()"><br>
<select id="img" style="width:80px"><option value="1">图像1<option value="2">图像2<option value="3">图像3</select><br>

<input type=button class="button"   value="选URL图片插入"  onclick="WebInsertURLImage(img.value)"><br>
<input type=button class="button"   value="用Excel求和"  onclick="WebGetExcelContent()"><br>
<input type=button class="button"   value="光标后插入文字"  onclick="WebInsertAfter()"><br>[<b style="color:green">Word对话框</b>]<br>
<input type=button class="button"   value="光标到行开始处"  onclick="frm.WebOffice.WebDialogs(4012);"><br>
<input type=button class="button"   value="光标到行尾"  onclick="frm.WebOffice.WebDialogs(4013);"><br>
<input type=button class="button"   value="光标到窗口开始处"  onclick="frm.WebOffice.WebDialogs(4014);"><br>
<input type=button class="button"   value="光标到窗口后面"  onclick="frm.WebOffice.WebDialogs(4015);"><br>
<input type=button class="button"   value="光标到文档开始处"  onclick="frm.WebOffice.WebDialogs(4016);"><br>
<input type=button class="button"   value="光标到文档后面"  onclick="frm.WebOffice.WebDialogs(4017);"><br>
<input type=button class="button"   value="查找EditFind"  onclick="frm.WebOffice.WebDialogs(112);"><br>
<input type=button class="button"   value="查找替换EditReplace"  onclick="frm.WebOffice.WebDialogs(117);"><br>
<input type=button class="button"   value="页眉NormalViewHeaderArea"  onclick="frm.WebOffice.WebDialogs(155);"><br>
<input type=button class="button"   value="插入文字框InsertFrame"  onclick="frm.WebOffice.WebDialogs(158);"><br>
<input type=button class="button"   value="插入分隔符InsertBreak"  onclick="frm.WebOffice.WebDialogs(159);"><br>

<input type=button class="button"   value="插入批注InsertAnnotation"  onclick="frm.WebOffice.WebDialogs(161);"><br>
<input type=button class="button"   value="插入特殊符号InsertSymbol"  onclick="frm.WebOffice.WebDialogs(162);"><br>
<input type=button class="button"   value="插入图片InsertPicture"  onclick="frm.WebOffice.WebDialogs(163);"><br>
<input type=button class="button"   value="插入文件InsertFile"  onclick="frm.WebOffice.WebDialogs(164);"><br>
<input type=button class="button"   value="插入日期时间InsertDateTime"  onclick="frm.WebOffice.WebDialogs(165);"><br>
<input type=button class="button"   value="插入域InsertField"  onclick="frm.WebOffice.WebDialogs(166);"><br>
<input type=button class="button"   value="编辑书签EditBookmark"  onclick="frm.WebOffice.WebDialogs(168);"><br>
<input type=button class="button"   value="标记索引贡MarkIndexEntry"  onclick="frm.WebOffice.WebDialogs(169);"><br>
<input type=button class="button"   value="插入索引InsertIndex"  onclick="frm.WebOffice.WebDialogs(170);"><br>
<input type=button class="button"   value="插入目录InsertTableOfContents"  onclick="frm.WebOffice.WebDialogs(171);"><br>
<input type=button class="button"   value="插入对象InsertObject"  onclick="frm.WebOffice.WebDialogs(172);"><br>

</td>
<td nowrap bgcolor=white valign=top style="padding-top:10px;">
<center><b style="font-family:黑体;font-size:25px;color:red;"><%=strfname%></b><br>Office控件在线演示文件编号：<input type=text class="text" name=num value="<%=strnum%>">&nbsp;&nbsp;文件名：<input type=text class="text" style="width:280" name=fname value="<%=strfname%>">附件：<input type="file" name="file1" id="file1" />&nbsp;<input type=hidden name=flsid value="<%=strflsid%>"><div align=left>&nbsp;无法看到控件?&nbsp;<a href="http://www.officectrl.com/weboffice/weboffice.rar"><img src="images/weboffice-install.jpg" border=0></a>完成后重新打开本页面。</div><script language=javascript src="cssjs/webofficeocx.js" charset="utf-8"></script><div id=divinfo style="display:none;">
<iframe style="position:absolute;left:300px;top:250px;background-color:transparent;" id=webifrm src="about:blank" width="580" height="19" scrolling="no" frameborder="0"></iframe>
<div style="position:absolute;left:300px;top:250px;background-color:red;color:white;display:none;" id=webinfo>正在处理请求和上传数据，为确保成功，请先不要执行其它操作， 请您稍等 。。。</div></div>

</td></tr></table>&nbsp;&nbsp;<font color=red>&copy;</font>2001-2016 All Rights Reserved!&nbsp;&nbsp;本站网址：<a href="http://www.officectrl.com/">http://www.officectrl.com</a><br></div><br><br></form>
</body></html>

