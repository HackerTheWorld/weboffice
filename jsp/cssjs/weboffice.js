<!--
function WebOpen()
{
 obj = document.all.item("WebOffice");
 if (obj !='')
 	{
 		//�ȴ��ؼ���ʼ����ϣ�ʱ�䳤�̿��Ը��������ٶ��趨��
 		setTimeout('openfile()',100);
	}
}
function openfile()
{ 
	try{
		switch(flag)
			{
			case '11':
					frm.WebOffice.Open(strOpenUrl,true,"Word.Document","","");break;
			case '12':
					frm.WebOffice.Open(strOpenUrl,true,"Excel.Sheet","","");break;
			case '13': 
					frm.WebOffice.Open(strOpenUrl,true,"PowerPoint.Show","","");break;
			case '2':
					frm.WebOffice.CreateNew("Excel.Sheet.8");				break;
			case '3':
					frm.WebOffice.CreateNew("PowerPoint.Show");break;
			default: 
					frm.WebOffice.CreateNew("Word.Document");	}
	}
	catch(e)
	{
		//�����������һ���ǿͻ���û�а�װWebOffice�ؼ�����ģ���ʾ�û����ذ�װ
		//	if(confirm('�������⣬��������״�ʹ�ã��������ذ�װWebOffice�ؼ���\n������Ѿ���װ��WebOffice�ؼ�����ע��Windows���µ�½�����þͿ����ˡ�')) 
		//		{
				//window.location.href='http://www.officectrl.com/weboffice/weboffice.rar';
		//		}
	}

}
function WebSave()
{
	alert('�����ܽ���ʾWeb�����ݡ�������WebOffice�ĵ��������첽(����)��ʽ���б��档\n���ȱ���WebOffice�ĵ����ݣ�����ٱ���Web�����ݡ���ɺ���ҳ����ʾ�ر�');
	try
	{
		frm.WebOffice.Save(strURL);
		 
	}
	catch(e)
	{
		alert('Զ�̱���Office�ĵ���������');
	}
	return true;
}
function MyTimer()
{ 
	if (autoSave==1)
	{
		info.innerHTML='<br><b style=color:red>�����ĵ���...</b><br>';
		info.style.display='';
		//frm.WebOffice.Save(strURL);
		frm.WebOffice.HttpInit();
		frm.WebOffice.HttpAddPostString('num',frm.num.value);
		frm.WebOffice.HttpAddPostString('fname',frm.fname.value);
		frm.WebOffice.HttpAddPostString('oper',frm.oper.value);
		frm.WebOffice.HttpAddPostString('flsid',frm.flsid.value);
		frm.WebOffice.HttpAddPostString('flag',frm.flag.value);	
		frm.WebOffice.HttpAddPostFile('file1',frm.file1.value);
		frm.WebOffice.HttpAddPostCurrFile("docfile","");
		frm.WebOffice.HttpPost(strRoot + '/upload.jsp');
		info.innerHTML='<br><b style=color:red>�������...</b><br>';
		var timer=setTimeout("MyTimer()",30000);
	}
}
function WebHttpSave()
{
	try
	{
		alert('�����ܽ���ʾWeb�����ݡ�������WebOffice�ĵ�����ͬʱ�ύ��ָ����URL·�����ղ��洢��');
		frm.WebOffice.HttpInit();
		frm.WebOffice.HttpAddPostString('num',frm.num.value);
		frm.WebOffice.HttpAddPostString('fname',frm.fname.value);
		frm.WebOffice.HttpAddPostString('oper',frm.oper.value);
		frm.WebOffice.HttpAddPostString('flsid',frm.flsid.value);
		frm.WebOffice.HttpAddPostString('flag',frm.flag.value);	 		 
		frm.WebOffice.HttpAddPostFile('file1',frm.file1.value);
		frm.WebOffice.HttpAddPostCurrFile("docfile","");	
		frm.WebOffice.HttpPost(strRoot + '/upload.jsp');
		
		alert('�����ݺ�office�ĵ�����ɹ�!');
	}
	catch(e)
	{
		alert('Զ�̱���Office�ĵ���������');
	}
 
	return true;

}
function WebSaveRemotePdf()
{
	try
	{
		if(pfile!='' && strPdfSaveUrl!=''){
			if(confirm('ע�⣺��������Ҫ��װOFFICE2010�����ϰ汾֧�֡�\n����Ѱ�װ����԰�ȷ����ʼ����ת��������\n�����밴ȡ����ťֹͣ���β���......')) 
			{
				alert('���ڿ�ʼִ��ת��PDF�ļ��Ĳ������˹���ʱ����ܱȽϳ���\n �����ĵȴ�ת���ɹ�����ʾ����ִ������������');
				//�����ڱ��ػ���������PDF�ļ�
				tempPath=frm.WebOffice.TempFilePath;
				var pdfile = tempPath +pfile+'.pdf';
				//�ؼ�����ִ���������û���ͣ������������ʾ
				webinfo.style.display='';
				divinfo.style.display='';
				//PDF���ɷ�ʽһ��
				frm.WebOffice.ActiveDocument.saveas(pdfile,17);
				//PDF���ɷ�ʽ����
				//frm.WebOffice.ActiveDocument.ExportAsFixedFormat(pdfile,17);
				//����ϴ� pdf�ļ���������������ջ������е��ĵ� 
				frm.WebOffice.WebSaveAsPDF(pdfile,strPdfSaveUrl);
				//�ر��û���ͣ������������ʾ
				webinfo.style.display='none';
				divinfo.style.display='none';

				if(confirm('�ѽ���ǰ�򿪵�WORD�ĵ�ת��PDF�ļ���Զ�̱������������ɹ���\n�����㱾�ص����Ѱ�װPDF�Ķ��������ھͿ��Դ򿪲鿴���Ƿ����ڴ򿪲鿴��'))
				{
					window.open(strRoot+'pdf/' +pfile+'.pdf','_blank');
				}
				
			}	
		}
	}
	catch(e)
	{
		webinfo.style.display='none';
		divinfo.style.display='none';
		alert('����ԭ���У�1�������ص�Office�汾���Ͳ�֧�ֽ�WORD�ļ�תΪPDF�ļ�,�밲װOFFICE2010���ϰ汾�� \n            2��Զ�̱���ʱ�������˷�������  ');
	}
}
function WebSaveRemoteHTML()
{
	try
	{
		if(pfile!='' && strHTMLSaveUrl!=''){
			alert('ע�⣺���ڿ�ʼִ��ת��HTML�ļ��Ĳ������˹���ʱ����ܱȽϳ���\n �����ĵȴ�ת���ɹ�����ʾ����ִ������������\n�����밴ȷ����ʼ����ת������......');
			var htmlpath=frm.WebOffice.TempFilePath;
			var htmlname=pfile;
			var htmlExtend =".html";
			var htmlfullpath= htmlpath+htmlname+htmlExtend;
			//�ؼ�����ִ���������û���ͣ������������ʾ
			divinfo.style.display='';
			webinfo.style.display='';
			frm.WebOffice.ActiveDocument.saveas(htmlfullpath,8) ;
			openfile(); 	 
			frm.WebOffice.WebSaveAsHTML(htmlpath,htmlname,htmlExtend,strHTMLSaveUrl);
			//�ر��û���ͣ������������ʾ
			webinfo.style.display='none';
			divinfo.style.display='none';
			if(confirm('�ѽ���ǰ�򿪵�WORD�ĵ�ת��HTML�ļ���Զ�̱������������ɹ����Ƿ����ڴ򿪲鿴��'))
			{
				window.open(strRoot+'html/' +pfile+htmlExtend,'_blank');
			} 
		}
	}
	catch(e)
	{
		alert('��ע�������ص��Ժ����ԣ�����������⣬��ϵ����Ա��');
	}	
}
function WebSaveLocalPdf()
{
	try
	{
		if(pfile!=''){
			var pdfile = 'C:\\'+pfile+'.pdf';
			//�ؼ�����ִ���������û���ͣ������������ʾ
			divinfo.style.display='';
			webinfo.style.display='';
			//PDF���ɷ�ʽһ��
			//frm.WebOffice.ActiveDocument.saveas(pdfile,17);
			//PDF���ɷ�ʽ����
			frm.WebOffice.ActiveDocument.ExportAsFixedFormat(pdfile,17);
			//�ر��û���ͣ������������ʾ
			webinfo.style.display='none';
			divinfo.style.display='none';
			alert('����C�̸�Ŀ¼������'+pdfile+'���뵽�㱾�ص��Ե�C��Ŀ¼�鿴��');
		}
	}
	catch(e)
	{
		alert('�����ص�Office�汾���Ͳ�֧�ֽ�WORDתΪPDF,�밲װOFFICE2010���ϰ汾�� ');
	}
}
function WebSaveLocalHTML()
{ 
	try
	{
		if(pfile!=''){
			var htmlpath="C:\\"
			var htmlname=pfile;
			var htmlExtend =".html";
			var htmlfullpath= htmlpath+htmlname+htmlExtend;
			//�ؼ�����ִ���������û���ͣ������������ʾ
			divinfo.style.display='';
			webinfo.style.display='';
			frm.WebOffice.ActiveDocument.saveas(htmlfullpath,8);
			//�ر��û���ͣ������������ʾ
			webinfo.style.display='none';
			divinfo.style.display='none';
			//openfile();
			alert('����C�̸�Ŀ¼������'+htmlfullpath+'���뵽�㱾�ص��Ե�C��Ŀ¼�鿴��');
		}
	}
	catch(e)
	{
		alert('��ע�������ص��Ժ����ԣ�����������⣬��ϵ����Ա��');
	}
}
function WebSaveXLSAsPDF()
{ 
	try{ 
	
		if(pfile!='' && strPdfSaveUrl!=''){
				if(confirm('ע�⣺��������Ҫ��װOFFICE2010�����ϰ汾֧�֡�\n����Ѱ�װ����԰�ȷ����ʼ����ת��������\n�����밴ȡ����ťֹͣ���β���......')) 
				{
					alert('���ڿ�ʼִ��ת��PDF�ļ��Ĳ������˹���ʱ����ܱȽϳ���\n �����ĵȴ�ת���ɹ�����ʾ����ִ������������');
					//�����ڱ��ػ���������PDF�ļ�
					tempPath=frm.WebOffice.TempFilePath;
					var pdfile = tempPath +pfile+'.pdf';
					//�ؼ�����ִ���������û���ͣ������������ʾ
					divinfo.style.display='';
					webinfo.style.display='';					 
					//PDF���ɷ�ʽһ�� 
					frm.WebOffice.ActiveDocument.Application.ActiveWindow.ActiveSheet.ExportAsFixedFormat(0,pdfile);
					//����ϴ� pdf�ļ���������������ջ������е��ĵ� 
					frm.WebOffice.WebSaveAsPDF(pdfile,strPdfSaveUrl);
					//�ر��û���ͣ������������ʾ
					webinfo.style.display='none';
					divinfo.style.display='none';
					if(confirm('�ѽ���ǰ�򿪵�EXCEL�ĵ�ת��PDF�ļ���Զ�̱������������ɹ���\n�����㱾�ص����Ѱ�װPDF�Ķ��������ھͿ��Դ򿪲鿴���Ƿ����ڴ򿪲鿴��'))
					{
						window.open(strRoot+'pdf/' +pfile+'.pdf','_blank');
					}
				}
			} 
	
	}
	catch(e)
	{
		alert('�����ص�Office�汾���Ͳ�֧�ֽ�EXCELתΪPDF,�밲װOFFICE2010���ϰ汾�� ');
	}
}
function WebSaveXLSLocalPDF(){
	try
	{
		var pdfile = 'C:\\'+pfile+'.pdf';
		//�ؼ�����ִ���������û���ͣ������������ʾ
		divinfo.style.display='';
		webinfo.style.display='';
		frm.WebOffice.ActiveDocument.Application.ActiveWindow.ActiveSheet.ExportAsFixedFormat(0,pdfile);
		//�ر��û���ͣ������������ʾ
		webinfo.style.display='none';
		divinfo.style.display='none'; 
		alert('����C�̸�Ŀ¼������'+pdfile+'���뵽�㱾�ص��Ե�C��Ŀ¼�鿴��');
	}
	catch(e)
	{
		 
	}
}

function WebSaveXLSLocalHTML(){
	try
	{
		var pdfile = 'C:\\'+pfile+'.html';
		//�ؼ�����ִ���������û���ͣ������������ʾ
		divinfo.style.display='';
		webinfo.style.display='';
		frm.WebOffice.ActiveDocument.Application.ActiveWorkbook.SaveAs(pdfile,44); 
		//�ر��û���ͣ������������ʾ
		webinfo.style.display='none';
		divinfo.style.display='none';
		alert('����C�̸�Ŀ¼������'+pdfile+'���뵽�㱾�ص��Ե�C��Ŀ¼�鿴��');
	}
	catch(e)
	{
		
	}
}
function WebSaveXLSAsHTML(){	
		try
	{
		if(pfile!='' && strHTMLSaveUrl!=''){
			alert('ע�⣺���ڿ�ʼִ��ת��HTML�ļ��Ĳ������˹���ʱ����ܱȽϳ���\n �����ĵȴ�ת���ɹ�����ʾ����ִ������������\n�����밴ȷ����ʼ����ת������......');
			var htmlpath=frm.WebOffice.TempFilePath;
			var htmlname=pfile;
			var htmlExtend =".html";
			var htmlfullpath= htmlpath+htmlname+htmlExtend;
			//�ؼ�����ִ���������û���ͣ������������ʾ
			//divinfo.style.display='';
			//webinfo.style.display='';			
			frm.WebOffice.ActiveDocument.Application.ActiveWorkbook.SaveAs(htmlfullpath,44);	
			
			frm.field1.value=htmlpath;
			frm.field2.value=htmlname;
			frm.field3.value=htmlExtend;
			frm.field4.value=strHTMLSaveUrl;
			 
			frm.action="excelhtml.jsp";
			frm.submit();
			//�ر��û���ͣ������������ʾ
		//	webinfo.style.display='none';
		//	divinfo.style.display='none';
	 
		
		}
	}
	catch(e)
	{
		alert(e);
		alert('��ע�������ص��Ժ����ԣ�����������⣬��ϵ����Ա��');
	}	

}
function WebSaveXLSAsMHT(){
	try
	{
		if(pfile!='' && strHTMLSaveUrl!=''){
			alert('ע�⣺���ڿ�ʼִ��ת��HTML�ļ��Ĳ������˹���ʱ����ܱȽϳ���\n �����ĵȴ�ת���ɹ�����ʾ����ִ������������\n�����밴ȷ����ʼ����ת������......');
			var htmlpath=frm.WebOffice.TempFilePath;
			var htmlname=pfile;
			var htmlExtend =".mht";
			var htmlfullpath= htmlpath+htmlname+htmlExtend;
			//�ؼ�����ִ���������û���ͣ������������ʾ
			divinfo.style.display='';
			webinfo.style.display='';
			frm.WebOffice.ActiveDocument.Application.ActiveWorkbook.SaveAs(htmlfullpath,45);			 
			frm.action="excelhtml.jsp";
			frm.field1.value=htmlpath;
			frm.field2.value=htmlname;
			frm.field3.value=htmlExtend;
			frm.field4.value=strHTMLSaveUrl;
			frm.submit(); 
			//�ر��û���ͣ������������ʾ
			webinfo.style.display='none';
			divinfo.style.display='none';
		}
	}
	catch(e)
	{
		alert('��ע�������ص��Ժ����ԣ�����������⣬��ϵ����Ա��');
	}	
}


function WebSavePPTLocalPDF(){
	try
	{	
		//�ؼ�����ִ���������û���ͣ������������ʾ
		divinfo.style.display='';
		webinfo.style.display='';
		var pdfile = 'C:\\'+pfile+'.pdf';
		frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs (pdfile,32);
		//�ر��û���ͣ������������ʾ
		webinfo.style.display='none';
		divinfo.style.display='none';
		alert('����C�̸�Ŀ¼������'+pdfile+'���뵽�㱾�ص��Ե�C��Ŀ¼�鿴��');
	}
	catch(e)
	{
		 alert(e);
	}
}

function WebSavePPTLocalJPG(){
	try
	{
		//�ؼ�����ִ���������û���ͣ������������ʾ
		divinfo.style.display='';
		webinfo.style.display='';
		var pdfile = 'C:\\'+pfile;
		frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs (pdfile,17);//jpg
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",6);//rtf
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",7);//pps
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",16);//gif
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",17);//jpg
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",18);//png
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",19);//bmp
		//frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs ("c:\\a15",21);//tif word
		//�ر��û���ͣ������������ʾ
		webinfo.style.display='none';
		divinfo.style.display='none';
		alert('�Ѱ�PPTת��ͼƬ����'+pdfile+'Ŀ¼�£��뵽�㱾�ص��Ե�C��Ŀ¼�鿴��');
	}
	catch(e)
	{
		alert(e);
	}
}
function WebSavePPTAsPDF()
{ 
	try{ 
	
		if(pfile!='' && strPdfSaveUrl!=''){
				if(confirm('ע�⣺��������Ҫ��װOFFICE2010�����ϰ汾֧�֡�\n����Ѱ�װ����԰�ȷ����ʼ����ת��������\n�����밴ȡ����ťֹͣ���β���......')) 
				{
					alert('���ڿ�ʼִ��ת��PDF�ļ��Ĳ������˹���ʱ����ܱȽϳ���\n �����ĵȴ�ת���ɹ�����ʾ����ִ������������');
					//�ؼ�����ִ���������û���ͣ������������ʾ
					divinfo.style.display='';
					webinfo.style.display='';
					//�����ڱ��ػ���������PDF�ļ�
					tempPath=frm.WebOffice.TempFilePath;
					var pdfile = tempPath +pfile+'.pdf';
					//PDF���ɷ�ʽһ�� 
					frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs (pdfile,32);
					//����ϴ� pdf�ļ���������������ջ������е��ĵ� 
					frm.WebOffice.WebSaveAsPDF(pdfile,strPdfSaveUrl);
					//�ر��û���ͣ������������ʾ
					webinfo.style.display='none';
					divinfo.style.display='none';
					if(confirm('�ѽ���ǰ�򿪵�PPT�ĵ�ת��PDF�ļ���Զ�̱������������ɹ���\n�����㱾�ص����Ѱ�װPDF�Ķ��������ھͿ��Դ򿪲鿴���Ƿ����ڴ򿪲鿴��'))
					{
						window.open(strRoot+'pdf/' +pfile+'.pdf','_blank');
					}
				}
			} 
	
	}
	catch(e)
	{
		alert('�����ص�Office�汾���Ͳ�֧�ֽ�PPTתΪPDF,�밲װOFFICE2010���ϰ汾�� ');
	}
}
function WebSavePPTAsHTML(){	
	try
	{
		if(pfile!='' && strppFileSaveUrl!=''){
			alert('ע�⣺���ڿ�ʼִ��ת��HTML�ļ��Ĳ������˹���ʱ����ܱȽϳ���\n �����ĵȴ�ת���ɹ�����ʾ����ִ������������\n�����밴ȷ����ʼ����ת������......');
			var htmlpath=frm.WebOffice.TempFilePath;
			var htmlname=pfile;
			var htmlExtend ='';
			var htmlfullpath= htmlpath+pfile;			
			//�ؼ�����ִ���������û���ͣ������������ʾ
			divinfo.style.display='';
			webinfo.style.display='';
			frm.WebOffice.ActiveDocument.Application.ActivePresentation.SaveAs (htmlfullpath,17);//jpg 
	 
			frm.WebOffice.WebSaveFormFolder(htmlfullpath+'\\',strppFileSaveUrl+'&file='+pfile);	
		 
			//�ر��û���ͣ������������ʾ
			webinfo.style.display='none';
			divinfo.style.display='none';
			if(confirm('�ѽ���ǰ�򿪵�PPT�ĵ�ת��HTML�ļ���Զ�̱������������ɹ���\n���ھͿ��Դ򿪲鿴���Ƿ����ڴ򿪲鿴��'))
					{
						window.open(strRoot+'html/' +pfile+'.html','_blank');
					}
		}
	}
	catch(e)
	{
		alert('��ע�������ص��Ժ����ԣ�����������⣬��ϵ����Ա��');
	}	

}
function WebSaveLocal()
{
	//��������Ի���
	frm.WebOffice.showdialog(3);
}
function WebOpenLocal()
{
	//�����򿪶Ի���
	frm.WebOffice.showdialog(1);
}

function WebDocReload()
{
	//alert('�����ܽ�����װ�ر���ҳ���������ݣ���ˢ�£�');
	location.reload();	
}
function WebOpenPicture()
{
	//��������Ի���
	frm.WebOffice.ActiveDocument.Application.Dialogs(163).Show();
}
 
function WebDocPageSetup()
{
	frm.WebOffice.showdialog(5);
}
function ShowRevision(boolvalue)
{	
	frm.WebOffice.ActiveDocument.ShowRevisions = boolvalue;
}
function WebAcceptAllRevisions()
{
	frm.WebOffice.ActiveDocument.AcceptAllRevisions();
}
function WebSignature(str)
{
	 
	var strPic ='';
	switch(str)
	{
		//�˴�������������URL
	case '1':
	strPic = strRoot + "/images/001.gif";
	break;
	case '2':
	strPic = strRoot + "/images/002.gif";
	break;
	case '3':
	strPic = strRoot + "/images/003.gif";
	break;		
	}  

	document.all.WebOffice.SetFieldValue('mark_1','','::ADDMARK::');
	document.all.WebOffice.SetFieldValue('mark_1',strPic,'::FLOATJPG::');
	var doc = frm.WebOffice.ActiveDocument;
	//doc.Shapes.AddPicture(strPic, false, true,100,0,207,209,doc.Application.Selection.Range);
	doc.Shapes(doc.Shapes.Count).Select(); 
	var range = doc.Application.Selection.ShapeRange;
	range.WrapFormat.Type = 3;
	range.PictureFormat.TransparentBackground = true;
	range.PictureFormat.TransparencyColor = 0xFFFFFF;
	range.Fill.Visible = false;

}


function WebAddFloatPic()
{
 
	document.all.WebOffice.SetFieldValue('mark_1','','::ADDMARK::');
	document.all.WebOffice.SetFieldValue('mark_1','http://www.officectrl.com/weboffice/images/weboffice.jpg','::FLOATJPG::');
	 
} 
function WebAddPic()
{
	
	document.all.WebOffice.SetFieldValue('mark_1','','::ADDMARK::');
	document.all.WebOffice.SetFieldValue('mark_1','http://www.officectrl.com/weboffice/images/weboffice.jpg','::JPG::');
	 
}
function WebDocSignature()
{
	try{
	
	frm.WebOffice.WebSign();	 
	var doc = frm.WebOffice.ActiveDocument;	
	document.all.WebOffice.SetFieldValue('mark_1','','::ADDMARK::');
	document.all.WebOffice.SetFieldValue('mark_1','c:\\Sign.bmp','::FLOATJPG::');
	//doc.Shapes.AddPicture('c:\\Sign.bmp', false, true,100,0,219,112,doc.Application.Selection.Range);
	doc.Shapes(doc.Shapes.Count).Select(); 
	var range = doc.Application.Selection.ShapeRange;
	range.WrapFormat.Type = 3;
	range.PictureFormat.TransparentBackground = true;
	range.PictureFormat.TransparencyColor = 0xFFFFFF;
	range.Fill.Visible = false;
	//frm.WebOffice.WebSignTempFileDel(); 
	//var strFile = frm.WebOffice.WebSignTempFile;	
	  
	}
	catch(E)
	{
		
	}
}
function WebTempFile(str)
{
	var strValue='';
	switch(str)
	{
		case '1':
		strValue='OfficeCTRL�����������ķ���';
			break;
		case '2':
		strValue='OfficeCTRL�����������Ĺ���';
		var doc = frm.WebOffice.ActiveDocument;	
		doc.Shapes.AddPicture(strRoot + "/images/weboffice.jpg",false, true,0,-60);
			break;
		case '3':
		strValue='OfficeCTRL�����������Ĺ���';		
			break;
		case '4':
		strValue='OfficeCTRL����������������';
			break;
		default:
		strValue='���������ļ�';
	}
	//����
	var object=frm.WebOffice.ActiveDocument;
	//var myl=object.Shapes.AddLine(100,60,305,60)
	//myl.Line.ForeColor=255;
	//myl.Line.Weight=2;
	//var myl1=object.Shapes.AddLine(326,60,520,60)
	//myl1.Line.ForeColor=255;
	//myl1.Line.Weight=2;

	//object.Shapes.AddLine(200,200,450,200).Line.ForeColor=6;
   	var myRange=frm.WebOffice.ActiveDocument.Range(0,0);
	myRange.Select();
	var mtext="��";
	frm.WebOffice.ActiveDocument.Application.Selection.Range.InsertAfter (mtext+"\n");
   	var myRange=frm.WebOffice.ActiveDocument.Paragraphs(1).Range;
   	myRange.ParagraphFormat.LineSpacingRule =1.5;
   	myRange.font.ColorIndex=6;
   	myRange.ParagraphFormat.Alignment=1;
   	myRange=frm.WebOffice.ActiveDocument.Range(0,0);
	myRange.Select();
	mtext="[����16]��72��";
	frm.WebOffice.ActiveDocument.Application.Selection.Range.InsertAfter (mtext+"\n");
	myRange=frm.WebOffice.ActiveDocument.Paragraphs(1).Range;
	myRange.ParagraphFormat.LineSpacingRule =1.5;
	myRange.ParagraphFormat.Alignment=1;
	myRange.font.ColorIndex=1;	
	mtext=strValue;
	frm.WebOffice.ActiveDocument.Application.Selection.Range.InsertAfter (mtext+"\n");
	myRange=frm.WebOffice.ActiveDocument.Paragraphs(1).Range;
	myRange.ParagraphFormat.LineSpacingRule =1.5;	
	//myRange.Select();
	myRange.Font.ColorIndex=6;
	myRange.Font.Name="����_GB2312";
	myRange.font.Bold=true;
	myRange.Font.Size=28;
	myRange.ParagraphFormat.Alignment=1;	
	//myRange=myRange=frm.WebOffice.ActiveDocument.Paragraphs(1).Range;
	frm.WebOffice.ActiveDocument.PageSetup.LeftMargin=70;
	frm.WebOffice.ActiveDocument.PageSetup.RightMargin=70;
	frm.WebOffice.ActiveDocument.PageSetup.TopMargin=70;
	frm.WebOffice.ActiveDocument.PageSetup.BottomMargin=70;
}
function WebSetWordTable()
{
	try{
		var mText="",mTmp="",iColumns=10,iCells=10,iPost,iold=-1;
		var myRange=frm.WebOffice.ActiveDocument.Range(0,0);     //���λ��
		frm.WebOffice.ActiveDocument.Tables.Add(myRange,10,10);   //���ɱ�� 
/*
		for (var n=0; n<iColumns; n++)
		{
			for (var i=0; i<iCells; i++)
			{
				iPos  = mText.indexOf(";",1+iold);
				mTmp = mText.substring(iold+1,iPos);
				frm.WebOffice.ActiveDocument.Tables(1).Columns(n+1).Cells(i+1).Range.Text=mTmp;   //��䵥Ԫֵ
				iold = iPos; 
			}
		}   
	*/
	} 
	catch(e)
	{
		alert(e);
	}

}
function WebGetWordContent()
{
  try{
    alert(frm.WebOffice.ActiveDocument.Content.Text);
  }catch(e){}
}
function WebSetWordContent()
{
	var mText=window.prompt("����������:","��������");
	if (mText==null){
		return (false);
	}
	else
	{
		 //����Ϊ��ʾѡ�е��ı�
		 //alert(frm.WebOffice.ActiveDocument.Application.Selection.Range.Text);
		 //����Ϊ�ڵ�ǰ���������ı�
		 frm.WebOffice.ActiveDocument.Application.Selection.Range.InsertAfter (mText+"\n");
		 //����Ϊ�ڵ�һ�κ�����ı�
		 //frm.WebOffice.ActiveDocument.Application.ActiveDocument.Range(1).InsertAfter(mText);
	}
}
function WebGetExcelContent()
{	try{
	frm.WebOffice.ActiveDocument.Application.Sheets(1).Select;
    frm.WebOffice.ActiveDocument.Application.Range("C5").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "126";
    frm.WebOffice.ActiveDocument.Application.Range("C6").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "446";
    frm.WebOffice.ActiveDocument.Application.Range("C7").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "556";
    frm.WebOffice.ActiveDocument.Application.Range("C5:C8").Select;
    frm.WebOffice.ActiveDocument.Application.Range("C8").Activate;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "=SUM(R[-3]C:R[-1]C)";
    frm.WebOffice.ActiveDocument.Application.Range("D8").Select;
    alert(frm.WebOffice.ActiveDocument.Application.Range("C8").Text);
	 }catch(e){
		alert('�˹��ܶ�Excel�ĵ���Ч!');
	 }
}
//���ã�����������Ԫ
function WebSheetsLock(){
	try{
    frm.WebOffice.ActiveDocument.Application.Sheets(1).Select;
    frm.WebOffice.ActiveDocument.Application.Range("A1").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "��Ʒ";
    frm.WebOffice.ActiveDocument.Application.Range("B1").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "�۸�";
    frm.WebOffice.ActiveDocument.Application.Range("C1").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "��ϸ˵��";
    frm.WebOffice.ActiveDocument.Application.Range("D1").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "���";
    frm.WebOffice.ActiveDocument.Application.Range("A2").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "��ǩ";
    frm.WebOffice.ActiveDocument.Application.Range("A3").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "ë��";
    frm.WebOffice.ActiveDocument.Application.Range("A4").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "�ֱ�";
    frm.WebOffice.ActiveDocument.Application.Range("A5").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "����";
    frm.WebOffice.ActiveDocument.Application.Range("B2").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "0.5";
    frm.WebOffice.ActiveDocument.Application.Range("C2").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "ӣ��";
    frm.WebOffice.ActiveDocument.Application.Range("D2").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "300";
    frm.WebOffice.ActiveDocument.Application.Range("B3").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "2";
    frm.WebOffice.ActiveDocument.Application.Range("C3").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "�Ǻ�";
    frm.WebOffice.ActiveDocument.Application.Range("D3").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "50";
    frm.WebOffice.ActiveDocument.Application.Range("B4").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "3";
    frm.WebOffice.ActiveDocument.Application.Range("C4").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "��ɫ";
    frm.WebOffice.ActiveDocument.Application.Range("D4").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "90";
    frm.WebOffice.ActiveDocument.Application.Range("B5").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "1";
    frm.WebOffice.ActiveDocument.Application.Range("C5").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "20cm";
    frm.WebOffice.ActiveDocument.Application.Range("D5").Select;
    frm.WebOffice.ActiveDocument.Application.ActiveCell.FormulaR1C1 = "40";
    //����������
    frm.WebOffice.ActiveDocument.Application.Range("B2:D5").Select;
    frm.WebOffice.ActiveDocument.Application.Selection.Locked = false;
    frm.WebOffice.ActiveDocument.Application.Selection.FormulaHidden = false;
    frm.WebOffice.ActiveDocument.Application.ActiveSheet.Protect(true,true,true);
    alert("�Ѿ�����������ֻ��B2-D5��Ԫ������޸ġ�");
	 }catch(e){
		alert('�˹��ܶ�Excel�ĵ���Ч!');
	 }
}
//���ã���ȡ�ĵ�ҳ��
function WebDocumentPageCount(){
	var intPageTotal;
	intPageTotal = frm.WebOffice.ActiveDocument.Application.ActiveDocument.BuiltInDocumentProperties(14);
	alert("�ĵ�ҳ������"+intPageTotal);
}

function WebTitlebar(boolvalue)
{
	frm.WebOffice.Titlebar = boolvalue;
}
function WebToolbar(boolvalue)
{
	frm.WebOffice.Toolbars = boolvalue;
}
function WebMenubar(boolvalue)
{
	   frm.WebOffice.MenuBars =boolvalue;
	   
}
function WebInsertImage()
{

	frm.WebOffice.ActiveDocument.Application.Selection.InlineShapes.AddPicture(strRoot+"/images/login.gif",false,true);

}
function WebInsertURLImage(str)
{
	var fileName='';

	switch(str)
	{
	case '2':
		fileName = strRoot + "/images/sec.jpg";		
		break;
	case '3':
		fileName = strRoot + "/images/buy.gif";
		break;
	default:
		fileName = strRoot + "/images/180.jpg";
	}
	frm.WebOffice.ActiveDocument.Application.Selection.InlineShapes.AddPicture(fileName,false,true);
}
function WebAddTemplate(str)
{
	WebTempFile(str);	
}
function WebPrintDirc()
{
frm.WebOffice.ActiveDocument.PrintOut();
}
function WebDocPrint()
{
 frm.WebOffice.printout(true);
}
function WebDocPrint2()
{
	 frm.WebOffice.ActiveDocument.Application.Dialogs(88).Show();
}
function WebDocPrintPreView()
{
frm.WebOffice.ActiveDocument.Application.PrintPreview=1;
}
function WebSaveC()
{
	try{
	frm.WebOffice.ActiveDocument.SaveAs("c:\\a.doc");
	alert('�ѱ��浽c��a.doc');
	}
	 catch(e){
		alert('��������!');
	 }

}function WebInsertAfter()
{
	frm.WebOffice.ActiveDocument.Application.Selection.Range.InsertAfter('��˾����');
	//������Ƶ��������ֺ���
	frm.WebOffice.ActiveDocument.Application.Dialogs(4013).Show();
}
function WebGetAllMark()
{
	var iCount=frm.WebOffice.ActiveDocument.BookMarks.count;
	for (i=1;i<=iCount ; i++ )
	{
		alert(frm.WebOffice.ActiveDocument.BookMarks.item(i).Name);
		 
	}
}
//-->