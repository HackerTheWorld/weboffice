<%@ page contentType="text/html;charset=gb2312" %>
<%@ page import="java.io.*" %>
<%@ page import="java.text.*" %>
<%@ page import="java.util.*" %>
<%@ page import="java.sql.*" %>
<%
String flsid = request.getParameter("flsid");
String strfile = request.getParameter("file");
String strfileextend = request.getParameter("fileextend");
String strfilenew = request.getParameter("filenew");
String rootpath=application.getRealPath("/");
if (strfile==null)strfile="";
if (strfilenew==null)strfilenew="";
if (strfileextend==null)strfileextend="";

String strSavePath="";
String strImagefloder = rootpath + "\\jsp\\html\\" + strfile + ".files";
//创建图片存储文件夹 
if (!(new java.io.File(strImagefloder).isDirectory())) //如果文件夹不存在
{
	new java.io.File(strImagefloder).mkdir();  
}
System.out.println(strImagefloder); 

//判断是HTML文件还是图片夹里的文件
if (strfile.equals(strfilenew))
{
	strSavePath=rootpath + "\\jsp\\html\\" + strfile + strfileextend;
	System.out.println(strSavePath);
}else
{
	strSavePath=rootpath + "\\jsp\\html\\" + strfile + ".files\\" + strfilenew;
} 
 
FileOutputStream fileOut=new FileOutputStream(strSavePath); 
DataInputStream din=new DataInputStream(request.getInputStream()); 
int formDataLength=request.getContentLength(); 
byte dataBytes[]=new byte[formDataLength];
 //int num=din.skipBytes(10);
din.readFully(dataBytes,0,formDataLength);
fileOut.write(dataBytes); 
fileOut.close(); 
din.close(); 



String DBDriver = "sun.jdbc.odbc.JdbcOdbcDriver"; 
String ConnStr="jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ="+rootpath+"\\database\\weboffice.mdb";			
//String ConnStr = "jdbc:odbc:webofficedsn";
Connection conn = null;
try 
{
	Class.forName(DBDriver).newInstance();  
} 
catch(java.lang.ClassNotFoundException e)
{ 
	System.err.println("Conn(): " + e.getMessage()); 
}

conn = DriverManager.getConnection(ConnStr); 

String sql=null;
PreparedStatement ps = null;

sql="Update dt_document set o_html=? where o_flsid='"+flsid+"'";
ps = conn.prepareStatement(sql);
ps.setInt(1,1);		 
ps.executeUpdate();

System.out.println(strfilenew); 
%>
