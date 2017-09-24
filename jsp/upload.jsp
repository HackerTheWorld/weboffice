<%@ page contentType="text/html;charset=gb2312" %>
<%@ page language="java" import="java.io.*" %>
<%@ page language="java" import="java.text.*" %>
<%@ page language="java" import="java.util.*" %>
<%@ page language="java" import="java.sql.*" %>
<%@ page language="java" import="com.jspsmart.upload.*" %>
<%

/*
		FileOutputStream fileOut=new FileOutputStream("c:\\aab.txt"); 
		DataInputStream din=new DataInputStream(request.getInputStream()); 
		int formDataLength=request.getContentLength(); 
		byte dataBytes[]=new byte[formDataLength];
		//int num=din.skipBytes(10);
		din.readFully(dataBytes,0,formDataLength);
		fileOut.write(dataBytes); 
		fileOut.close(); 
		din.close(); 


*/


        SmartUpload uploader = new SmartUpload();
		String msg="";
		String flag ="";
		String flsid="";
		String num = "";
		String fname ="";
		String oper ="";
		long filesize=0;
        try {
             uploader.initialize(config, request, response);// ��ʼ��������
             uploader.upload(); // ���ر�����
			// ��ʱ���ܶ�ȡ������
             Enumeration<?> e = uploader.getRequest().getParameterNames();
             while (e.hasMoreElements()) { //�������б�����(�������ļ�)
                 String key = (String) e.nextElement();
				 
                 if ("num".equals(key)) { //�ҵ���Ҫ�Ĳ���
					//������request.getParameter()��ֻ������������ȡ����ֵ
                    num = uploader.getRequest().getParameterValues(key)[0];
					System.out.println(num+"<br>");
                 }
				if ("fname".equals(key)) { //�ҵ���Ҫ�Ĳ���
					//������request.getParameter()��ֻ������������ȡ����ֵ
                    fname = uploader.getRequest().getParameterValues(key)[0];
					System.out.println(fname+"<br>");
                 }
				if ("oper".equals(key)) { //�ҵ���Ҫ�Ĳ���
					//������request.getParameter()��ֻ������������ȡ����ֵ
                    oper = uploader.getRequest().getParameterValues(key)[0];
					System.out.println(oper+"<br>");
                 }
				if ("flsid".equals(key)) { //�ҵ���Ҫ�Ĳ���
					//������request.getParameter()��ֻ������������ȡ����ֵ
                    flsid = uploader.getRequest().getParameterValues(key)[0];
					System.out.println(flsid+"<br>");
                 }
				if ("flag".equals(key)) { //�ҵ���Ҫ�Ĳ���
					//������request.getParameter()��ֻ������������ȡ����ֵ
                    flag = uploader.getRequest().getParameterValues(key)[0];
					System.out.println(flag+"<br>");
                 }
             } 
			for (int i = 0; i < uploader.getFiles().getCount(); i++) {
                 com.jspsmart.upload.File myFile = uploader.getFiles().getFile(i);
                 if (!myFile.isMissing()) { //�ļ��ϴ��ɹ�
					   
                      String fileName = "/upload/"+myFile.getFileName();//new SimpleDateFormat("yyyyMMdd").format(new Date())+ (int) (Math.random() * 90+10)+"." + myFile.getFileExt();
						String fileExt=myFile.getFieldName();
					   if (fileExt.equals("docfile") && (flag.equals("1") || flag.equals("11")))
					  {
						fileName = "/upload/"+ flsid +".doc";
						filesize =myFile.getSize();
					   }
					    if (fileExt.equals("docfile") && (flag.equals("2") || flag.equals("12")))
					  {
						fileName = "/upload/"+ flsid +".xls";
						filesize =myFile.getSize();
					   }
					    if (fileExt.equals("docfile") && (flag.equals("3") || flag.equals("13")))
					  {
						fileName = "/upload/"+ flsid +".ppt";
						filesize =myFile.getSize();
					   }

					//  out.println(fileName+"<br>");
                     myFile.saveAs(fileName, uploader.SAVE_VIRTUAL);    

                 } //��һ��Ϊ��ʾ��Ϣ
            } 
		// msg="�ϴ��ɹ�,���ϴ�"+uploader.getFiles().getCount()+"���ļ�.";
        } catch (SmartUploadException e) {
            msg=e.getMessage(); //��������Ϣ����ʾ��Ϣ��ʽ��ʾ
            e.printStackTrace();
        }
			

			String DBDriver = "sun.jdbc.odbc.JdbcOdbcDriver"; 
			String rootpath=application.getRealPath("/");
			String ConnStr="jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ="+rootpath+"\\database\\weboffice.mdb";			
	  		
			//String ConnStr = "jdbc:odbc:webofficedsn";
			Connection conn = DriverManager.getConnection(ConnStr); 
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
				String sql="";
				String id="";
				PreparedStatement ps = null;
				if (oper.equals("doedit") || oper.equals("edit"))
				{ 
					sql="Update dt_document set o_size=?,o_name=? where o_flsid='"+flsid+"'";
				 System.out.println(sql);
					ps = conn.prepareStatement(sql);
				 	ps.setString(1,Long.toString(filesize));	
					ps.setString(2,fname);	
					ps.executeUpdate();
					
				}	else
				{
			 
					sql="INSERT INTO dt_document(o_flsid,o_size,o_name,o_flag) VALUES(?,?,?,?)";
					ps = conn.prepareStatement(sql);
					ps.setString(1,flsid);
					ps.setString(2,Long.toString(filesize));	
					ps.setString(3,fname);
					ps.setString(4,flag);
					ps.executeUpdate();
					 
				}
	

			

			}catch(SQLException e)
			{ 
				System.err.println("Upload Query Errinfo: " + flsid + filesize +fname + e.getMessage()); 
			}
 
		 	conn.close();



%>