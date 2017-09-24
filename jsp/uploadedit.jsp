<%@ page contentType="text/html;charset=gb2312" %>
<%@ page import="java.io.*" %>
<%@ page import="java.text.*" %>
<%@ page import="java.util.*" %>
<%@ page import="java.sql.*" %>
<%

String oper = request.getParameter("oper");
String flsid = request.getParameter("flsid");
String flag = request.getParameter("flag");
String id = request.getParameter("id");
if (flsid==null)flsid="nothing"; 
String DBDriver = "sun.jdbc.odbc.JdbcOdbcDriver"; 
String rootpath=application.getRealPath("/");
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
try
{ 

		String filepath="";
		if   (flag.equals("1") || flag.equals("11")) 
		{
			filepath = "\\upload\\"+ flsid +".doc";
		}
		if  (flag.equals("2") || flag.equals("12")) 
		{
			filepath = "\\upload\\"+ flsid +".xls";
		}
		if  (flag.equals("3") || flag.equals("13")) 
		{
			filepath = "\\upload\\"+ flsid +".ppt";
		}
 
		filepath = rootpath +filepath;

		conn = DriverManager.getConnection(ConnStr); 
		FileOutputStream fileOut=new FileOutputStream(filepath); 
		DataInputStream din=new DataInputStream(request.getInputStream()); 
		int formDataLength=request.getContentLength(); 
		byte dataBytes[]=new byte[formDataLength];
		 //int num=din.skipBytes(10);
		din.readFully(dataBytes,0,formDataLength);
		fileOut.write(dataBytes); 
		fileOut.close(); 
		din.close(); 
		
  
		String sql=null;
		PreparedStatement ps = null;
		if (oper.compareTo("edit")==0)
		{ 
			sql="Update dt_document set o_size=? where o_pkid="+id;
			ps = conn.prepareStatement(sql);
			ps.setInt(1,formDataLength);		 
			ps.executeUpdate();
		}
		else
		{ 
			sql="INSERT INTO dt_document(o_flsid,o_size) VALUES(?,?)";
			ps = conn.prepareStatement(sql);
			ps.setString(1,flsid);
			ps.setInt(2,formDataLength);	 
			ps.executeUpdate();
		} 
		/*
		java.io.File file = new java.io.File("c:\\jsptemp.txt"); 
		java.io.InputStream fin = new java.io.FileInputStream(file); 	
		String sql=null;
		PreparedStatement ps = null;
		if (oper.compareTo("edit")==0)
		{ 
			sql="Update dt_document set o_size=?,o_file=? where o_pkid="+id;
			ps = conn.prepareStatement(sql);
			ps.setInt(1,formDataLength);
			ps.setBinaryStream(2,fin,fin.available());	 
			ps.executeUpdate();
		}
		else
		{ 
			sql="INSERT INTO dt_document(o_flsid,o_size,o_file) VALUES(?,?,?)";
			ps = conn.prepareStatement(sql);
			ps.setString(1,flsid);
			ps.setInt(2,formDataLength);
			ps.setBinaryStream(3,fin,fin.available());
			ps.executeUpdate();
		} 
		*/
} 
catch(SQLException e)
{ 
	System.err.println("Upload QueryErr: " + e.getMessage()); 
}
conn.close();
%>
