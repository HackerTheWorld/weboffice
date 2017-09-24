<%@ page contentType="text/html;charset=gb2312" language="java" %>
<%@ page import="java.io.*" %>
<%
String field1 = request.getParameter("field1"); 
String field2 = request.getParameter("field2"); 
String field3 = request.getParameter("field3"); 
String field4 = request.getParameter("field4"); 
String field5 = request.getParameter("field5"); 
String field6 = request.getParameter("field6"); 
System.out.println(field1);
%><!DOCTYPE html><html><head><meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1"><meta http-equiv="content-type" content="text/html;charset=utf-8">
</head><body onload="savehtml();">
<script language=javascript>
function savehtml()
{
			var htmlpath='<%=field1%>';
			var htmlname='<%=field2%>';
			var htmlExtend ='<%=field3%>';
			var strHTMLSaveUrl='<%=field4%>';
			var strRoot='<%=field6%>';
			
			frm.WebOffice.WebSaveAsHTML(htmlpath,htmlname,htmlExtend,strHTMLSaveUrl);
			if(confirm('已将当前打开的EXCEL文档转成HTML文件并远程保存至服务器成功！是否现在打开查看？'))
			{
				window.open(strRoot+'html/' +htmlname+htmlExtend,'_self');				
			}else{
			  window.open('<%=field5%>','_self');}
}
</script>
<form   name=frm method="post">
<script language=javascript src="cssjs/webofficeocx.js" charset="utf-8"></script> </form>
</body></html>