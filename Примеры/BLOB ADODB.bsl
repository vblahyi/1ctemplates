


Функция ЗаписатьОтчетВБазуТелеграм(ИмяОтчета,ДатаОтчета,ВариантОтчета,РасширениеФайла,ПолноеИмяФайла) Экспорт
		
	Соединение = ПолучитьСоединениеСБазой();    	
	НаборЗаписей = Новый COMОбъект("ADODB.Recordset");
	КомандаSQL = Новый COMОбъект("ADODB.Command");
	КомандаSQL.ActiveConnection = Соединение;
	КомандаSQL.CommandText=    "INSERT INTO TABLE_NAME(ReportName, DateTime, ReportType,Data, FileExtension) VALUES(?,?,?,?,?)";    
	КомандаSQL.CommandType=1;//adCmdText
		
	//ReportName
	//name, type, direction=1, size
	ПараметрReportName = КомандаSQL.CreateParameter("ReportName", 200, 1, 100);//adVarChar, adParamInput	
	ПараметрReportName.Value = Строка(ИмяОтчета); 
	КомандаSQL.Parameters.Append(ПараметрReportName);
	
	//DateTime
	ПараметрDateTime = КомандаSQL.CreateParameter("DateTime", 135);//adDBTimeStamp 	
	ПараметрDateTime.Value = Строка(ДатаОтчета); 
	КомандаSQL.Parameters.Append(ПараметрDateTime);
	
	//ReportName
	//name, type, direction=1, size
	ПараметрReportType = КомандаSQL.CreateParameter("ReportType", 200, 1, 100);//adVarChar, adParamInput	
	ПараметрReportType.Value = Строка(ВариантОтчета); 
	КомандаSQL.Parameters.Append(ПараметрReportType);
	
	
	//DATA	
	Поток = Новый COMОбъект("ADODB.Stream");
	Поток.Type = 1;
	Поток.Open();
	//ДанныеФайла = Новый Файл(ПолноеИмяФайла);
	Поток.LoadFromFile(ПолноеИмяФайла);
	Если Поток.Size > 0 тогда
		ПараметрDATA = КомандаSQL.CreateParameter("Data", 128, 1, Поток.Size);//adLongVarBinary, adParamInput
		
		Пока НЕ Поток.EOS Цикл
			ПараметрDATA.AppendChunk(Поток.Read(10240));
		КонецЦикла;     
		КомандаSQL.Parameters.Append(ПараметрDATA);
	КонецЕсли;	
	//ReportName
	//name, type, direction=1, size
	ПараметрFileExtension = КомандаSQL.CreateParameter("FileExtension", 200, 1, 100);//adVarChar, adParamInput	
	ПараметрFileExtension.Value = Строка(РасширениеФайла); 
	КомандаSQL.Parameters.Append(ПараметрFileExtension);

	
	//PEW
	НаборЗаписей=КомандаSQL.Execute();

		
КонецФункции // ()


Функция ПолучитьСоединениеСБазой()
    extConnSQL = Новый ComОбъект("ADODB.Connection");
    СтрокаСоединения =  "Provider=SQLOLEDB.1;
    |User ID=login;
    |Pwd=pass;
    |Data Source=server;
    |Initial Catalog=base";
    extConnSQL.ConnectionString = СтрокаСоединения;
    extConnSQL.Open();
    Возврат extConnSQL;
КонецФункции

