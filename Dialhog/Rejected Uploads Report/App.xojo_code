#tag Class
Protected Class App
Inherits ConsoleApplication
	#tag Event
		Function Run(args() as String) As Integer
		  RegisterMBS()
		  
		  Log = New LoggingClass("C:\Log\" + ExecutableFile.Name + ".sqlite")
		  Log.Write("Starting", LoggingClass.LogLevels.Information)
		  
		  Dim i As Integer
		  
		  For i = 0 To args.Ubound
		    Select Case NthField(args(i), "=", 1)
		    Case "TO"
		      Recipients = Split(NthField(args(i), "=", 2), ",")
		    Case "CC"
		      CCRecipients = Split(NthField(args(i), "=", 2), ",")
		    Case "BCC"
		      BCCRecipients = Split(NthField(args(i), "=", 2), ",")
		    End Select
		  Next
		  
		  If Recipients.Ubound < 0 Then
		    Log.Write("No recipients specified.", LoggingClass.LogLevels.Error)
		    Quit(2)
		  End If
		  
		  // Initialise MailSocket Object
		  MailSocket = New MailSocket
		  MailSocket.Address = "smtp.voyager.co.nz"
		  
		  InitialiseEmailMessage()
		  
		  // Connect to database
		  DialhogConnection.GetSettings
		  // Dialhog.DataConfig()
		  db = New SQLDatabaseMBS
		  db.Option("UseAPI") = "OLEDB"
		  db.UserName = DialhogConnection.UserName
		  db.Password = DialhogConnection.Password
		  db.DatabaseName = "SQLServer:DHSQL@Dialhog"
		  
		  Try
		    db.Connect
		  Catch err As DatabaseException
		    Break
		  End Try
		  
		  Log.Write("Connected to database", LoggingClass.LogLevels.Information)
		  // Create Report here
		  
		  GenerateReport
		  
		  // Create Event loop to support asynchronous smtp socket
		  
		  Do
		    DoEvents
		  Loop Until MailSent
		  
		End Function
	#tag EndEvent

	#tag Event
		Function UnhandledException(error As RuntimeException) As Boolean
		  Log.Write(error.Message + EndOfLine + Join(error.Stack, EndOfLine), LoggingClass.LogLevels.Error)
		  Quit(1)
		End Function
	#tag EndEvent


	#tag Method, Flags = &h21
		Private Sub GenerateReport()
		  Var rs As RowSet
		  
		  EndTime = DateTime.Now
		  StartTime = EndTime.SubtractInterval(0, 0, 1)
		  
		  rs = App.db.SelectSQL(kSelectEvents, StartTime.SQLDateTime, EndTime.SQLDateTime)
		  
		  If Not rs.AfterLastRow Then
		    For Each Row As DatabaseRow In rs
		      REJEvents.Add(Row.Column("eventid").StringValue)
		      LERREvents.Add(Row.Column("eventid").StringValue)
		    Next
		  End If
		  
		  rs.Close
		  
		  stdout.WriteLine(App.REJEvents.Count.ToString + " events found")
		  
		  ProcessREJEvents()
		  
		  ProcessLERREvents()
		  
		  ProcessFTPErrorFiles()
		  
		  ProcessXMLErrorFiles()
		  
		  Var eml As New EmailMessage
		  eml.FromAddress = "reportservice@dialhog.com"
		  
		  Var Body As String = "<HTML><BODY>"
		  
		  // Create Summary Table
		  Body = Body + "<TABLE border=""1""><TR><TH align=""left"">Description</TH><TH align=""right"">Count</TH></TR>"
		  
		  // REJ Summary
		  Body = Body + "<TR><Td>Files with 10 or more sequential REJ statuses</Td><Td align=""right"">" + App.REJEvents.Count.ToString + "</Td></TR>"
		  
		  // LERR Summary
		  Body = Body + "<TR><Td>Files with 10 or more LERR statuses</Td><Td align=""right"">" + App.LERREvents.Count.ToString + "</Td></TR>"
		  
		  // XML Summary
		  Body = Body + "<TR><Td>Rejected CSV Uploaded Files</Td><Td align=""right"">" + App.XMLFiles.Count.ToString + "</Td></TR>"
		  
		  // FTP Summary
		  Body = Body + "<TR><Td>Rejected FTP Uploaded Files</Td><Td align=""right"">" + App.FTPFiles.Count.ToString + "</Td></TR>"
		  
		  Body = Body + "</TABLE><p>" + EndOfLine
		  
		  If App.REJEvents.Count > 0 _
		    Or App.LERREvents.Count > 0 _
		    Or App.FTPFiles.Count > 0 _
		    Or App.XMLFiles.Count > 0 Then
		    eml.Headers.AddHeader("Importance", "High")
		    
		    Body = Body + "<TABLE border=""1""><TR><TH>Reject Report Details</TH></TR>"
		    If App.REJEvents.Count > 0 Then
		      // REJ Detail
		      Body = Body + "<TR><TH>Batches with &gt 10 sequential REJ Statuses</th></tr><tr>"
		      Body = Body + "<TABLE border=""0"" width=""100%""><TR><TH align=""left"">Client</TH><th align=""left"">Description</th><th align=""right"">Created</th><th align=""right"">Delivery At</th></tr>"
		      Var s As String = String.FromArray(App.REJEvents, ",")
		      rs = App.db.SelectSQL("SELECT s.batchID, s.firststartdatetimelocal, m.description, c.businessname, m.createddate FROM schedule AS s " + _
		      "INNER JOIN messageevent AS m On m.Eventid = s.eventid INNER JOIN customer AS c ON c.customerid = m.customerid " + _
		      "WHERE s.eventid IN (" + s + ") ORDER BY c.businessname;")
		      If Not rs.AfterLastRow Then
		        For Each Row As DatabaseRow In rs
		          Var created As DateTime = DateTime.FromString(row.Column("createddate").StringValue, Nil, New TimeZone(0))
		          created = New DateTime(created.SecondsFrom1970, TimeZone.Current)
		          Body = Body + "<tr><td>" + Row.Column("businessname").StringValue + "</td>" + _
		          "<td>" + row.Column("description").StringValue + "</td>" + _
		          "<td align=""right"">" + created.SQLDateTime + "</td>" + _
		          "<td align=""right"">" + row.Column("firststartdatetimelocal").StringValue + "</td></tr>"
		        Next Row
		      End If
		      Body = Body + "</table></tr>" + EndOfLine
		      App.Log.Log("REJ detail written")
		    End If
		    
		    If App.LERREvents.Count > 0 Then
		      // LERR Detail
		      Body = Body + "<TR><th>Batches with &gt 10 LERR statuses</th></tr><tr>"
		      Body = Body + "<TABLE border=""0"" width=""100%""><TR><TH align=""left"">Client</TH><th align=""left"">Description</th><th align=""right"">Created</th><th align=""right"">Delivery At</th></tr>"
		      Var s As String = String.FromArray(App.LERREvents, ",")
		      rs = App.db.SelectSQL("SELECT s.batchID, s.firststartdatetimelocal, m.description, c.businessname, m.createddate FROM schedule AS s " + _
		      "INNER JOIN messageevent AS m On m.Eventid = s.eventid INNER JOIN customer AS c ON c.customerid = m.customerid " + _
		      "WHERE s.eventid IN (" + s + ") ORDER BY c.businessname;")
		      If Not rs.AfterLastRow Then
		        For Each Row As DatabaseRow In rs
		          Var created As DateTime = DateTime.FromString(row.Column("createddate").StringValue, Nil, New TimeZone(0))
		          created = New DateTime(created.SecondsFrom1970, TimeZone.Current)
		          Body = Body + "<tr><td>" + Row.Column("businessname").StringValue + "</td>" + _
		          "<td>" + row.Column("description").StringValue + "</td>" + _
		          "<td align=""right"">" + created.SQLDateTime + "</td>" + _
		          "<td align=""right"">" + row.Column("firststartdatetimelocal").StringValue + "</td></tr>"
		        Next Row
		      End If
		      Body = Body + "</table></tr>" + EndOfLine
		      App.Log.Log("LERR detail written")
		    End If
		    
		    If App.XMLFiles.Count > 0 Then
		      // Rejected CSV Files
		      Body = Body + "<tr><th>Rejected CSV files</th></tr><tr>"
		      Body = Body + "<table border = ""0"" width=""100%""><tr><th align=""left"">File Name</th><th align=""left"">Client</th><th align=""left"">User Name</th><th align=""right"">Uploaded At</th></tr>"
		      For Each Item As FolderItem In App.XMLFiles
		        Var xDoc As New XMLDocument
		        xDoc.LoadXML(Item)
		        Var FileName As String = xDoc.FirstChild.Child(5).FirstChild.Value
		        Var UserID As String = xDoc.FirstChild.Child(6).FirstChild.Value
		        rs = App.db.SelectSQL("SELECT username, businessname FROM [user] INNER JOIN customer ON customer.customerid = [user].customerid WHERE userid = ?", UserID)
		        Body = Body + "<tr><td>" + FileName + "</td><td>" + rs.Column("businessname").StringValue + "</td><td>" + rs.Column("username").StringValue + "</td><td align=""right"">" + item.ModificationDateTime.SQLDateTime + "</td></tr>"
		      Next Item
		      Body = Body + "</table></tr>" + EndOfLine
		      App.Log.Log("XML detail written")
		      
		    End If
		    
		    If App.FTPFiles.Count > 0 Then
		      // Rejected FTP Files
		      Body = Body + "<tr><th>Rejected FTP Uploaded Files</th></tr><tr>"
		      Body = Body + "<table border=""0"" width=""100%""><tr><th align=""left"">File Name</th><th align=""left"">Client</th><th align=""right"">Uploaded At</th></tr>"
		      For Each Item As FolderItem In App.FTPFiles
		        Body = Body + "<tr><td>" + Item.Name + "</td><td>" + Item.Parent.Parent.Name + "</td><td align=""right"">" + Item.ModificationDateTime.SQLDateTime + "</td></tr>"
		      Next Item
		      Body = Body + "</table></tr>" + EndOfLine
		    End If
		    Body = Body + "</table>"
		    App.Log.Log("FTP detail written")
		  End If
		  
		  Body = Body + "</body></html>"
		  
		  #If DebugBuild
		    eml.AddRecipient("wayne@dialhog.com")
		  #Else
		    For Each Recipient As String In Recipients
		      eml.AddRecipient(Recipient)
		    Next Recipient
		  #Endif
		  
		  eml.Subject = "File Upload status report"
		  
		  eml.BodyHTML = Body
		  
		  App.MailSocket.Messages.Add(eml)
		  App.SendMail()
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub InitialiseEmailMessage()
		  // Initialise Email Message
		  Var i As Integer
		  msg = New EmailMessage
		  For i = 0 to Recipients.Ubound
		    msg.AddRecipient(Recipients(i))
		  Next
		  For i = 0 to CCRecipients.Ubound
		    msg.AddCCRecipient(CCRecipients(i))
		  Next
		  For i = 0 to BCCRecipients.Ubound
		    msg.AddBCCRecipient(BCCRecipients(i))
		  Next
		  
		  msg.FromAddress = Sender
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ProcessFTPErrorFiles()
		  Var f As New FolderItem("C:\FTPUploader\Work", FolderItem.PathModes.Native)
		  
		  For Each ClientFolder As FolderItem In f.Children
		    Var ErrorFolder As FolderItem = ClientFolder.Child("Error")
		    If ErrorFolder.Exists Then
		      For Each Item As FolderItem In ErrorFolder.Children
		        If Item.ModificationDateTime > StartTime Then
		          Var t As TextInputStream = TextInputStream.Open(Item)
		          Var s As String = t.ReadAll.Trim
		          t.Close
		          If (ErrorFolder.Parent.Name = "waikato dhb auto ftp" _
		            Or ErrorFolder.Parent.Name = "northland dhb" _
		            Or ErrorFolder.Parent.Name = "cmdhb - main ftp upload" _
		            Or ErrorFolder.Parent.Name = "HealthAlliance") _
		            And s.CountFields(EndOfLine) > 1 Then
		            FTPFiles.Add(Item)
		          Else
		            If s.CountFields(EndOfLine) > 2 Then
		              FTPFiles.Add(Item)
		            End If
		          End If
		        End If
		      Next Item
		    End If
		  Next ClientFolder
		  
		  App.Log.Log("Found " + App.FTPFiles.Count.ToString + " FTP Files in Error Folders")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ProcessLERREvents()
		  Var s As String = String.FromArray(App.LERREvents, ",")
		  
		  App.LERREvents.RemoveAll
		  
		  Var rs As RowSet
		  rs = App.db.SelectSQL(kLERREvents.Replace("?", s))
		  
		  If Not rs.AfterLastRow Then
		    For Each Row As DatabaseRow In rs
		      App.LERREvents.Add(rs.Column("eventid").StringValue)
		    Next
		  End If
		  
		  If App.LERREvents.Count = 0 Then
		    Return
		  End If
		  
		  stdout.WriteLine("Found " + App.LERREvents.Count.ToString + " Events with 10 or more sequential LERR statuses.")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ProcessREJEvents()
		  Var s As String = String.FromArray(App.REJEvents, ",")
		  
		  App.REJEvents.RemoveAll
		  
		  Var rs As RowSet
		  rs = App.db.SelectSQL(kREJEvents.Replace("?", s))
		  
		  If Not rs.AfterLastRow Then
		    For Each Row As DatabaseRow In rs
		      App.REJEvents.Add(rs.Column("eventid").StringValue)
		    Next
		  End If
		  
		  If App.REJEvents.Count = 0 Then
		    Return
		  End If
		  
		  // Remove events where the sequential count of REJ statuses < 10
		  
		  For i As Integer = App.REJEvents.LastIndex DownTo 0
		    rs = App.db.SelectSQL("SELECT eventlogid FROM eventlog WHERE processingstatusid = 5 AND eventid = ?;", REJEvents(i).ToInteger)
		    Var count As Integer
		    Var previousid As Integer
		    If Not rs.AfterLastRow Then
		      For Each Row As DatabaseRow In rs
		        If rs.Column("eventlogid").IntegerValue > previousid + 1 Then
		          count = 0
		        Else
		          count = count + 1
		          If count > 9 Then
		            Exit For
		          End If
		        End If
		        previousid = rs.Column("eventlogid").IntegerValue
		      Next
		    End If
		    If count < 10 Then
		      App.REJEvents.RemoveAt(i)
		    End If
		  Next
		  
		  stdout.WriteLine("Found " + App.REJEvents.Count.ToString + " Events with 10 or more sequential REJ statuses.")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ProcessXMLErrorFiles()
		  Var f As New FolderItem("C:\inetpub\wwwroot\dialhog.com\temp\xml\error", FolderItem.PathModes.Native)
		  
		  For Each Item As FolderItem In f.Children
		    If Item.ModificationDateTime > StartTime Then
		      App.XMLFiles.Add(Item)
		    End If
		  Next Item
		  
		  App.Log.Log("Found " + App.XMLFiles.Count.ToString + " CSV Files in Error Folders")
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub RegisterMBS()
		  Var serial1 as integer = 13847453 
		  Var serial2 as integer = 32535266
		  Var year as integer = 2023
		  Var month as integer = 7
		  Var x100 as integer = 100
		  Var name as string  = DecodeBase64("V2F5bmUgR29sZGluZw==", encodings.UTF8)
		  SQLGlobalsMBS.setLicenseCode name, year*x100+month, serial1*x100+02, serial2*x100+13
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SendMail()
		  REM MailSocket.Messages.Append(msg)
		  MailSocket.SendMail
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		BCCRecipients() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		CCRecipients() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		db As SQLDatabaseMBS
	#tag EndProperty

	#tag Property, Flags = &h0
		EndTime As DateTime
	#tag EndProperty

	#tag Property, Flags = &h0
		FTPFiles() As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h0
		LERREvents() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		Log As LoggingClass
	#tag EndProperty

	#tag Property, Flags = &h0
		MailSent As Boolean = False
	#tag EndProperty

	#tag Property, Flags = &h21
		Private MailServerAddress As String = "smtp.voyager.co.nz"
	#tag EndProperty

	#tag Property, Flags = &h0
		MailSocket As MailSocket
	#tag EndProperty

	#tag Property, Flags = &h0
		msg As EmailMessage
	#tag EndProperty

	#tag Property, Flags = &h0
		Recipients() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		REJEvents() As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private Sender As String = "ReportService@dialhog.com"
	#tag EndProperty

	#tag Property, Flags = &h0
		StartTime As DateTime
	#tag EndProperty

	#tag Property, Flags = &h0
		XMLFiles() As FolderItem
	#tag EndProperty


	#tag Constant, Name = kLERREvents, Type = String, Dynamic = False, Default = \"SELECT DISTINCT eventid\r\nFROM eventlog\r\nWHERE eventid IN (\?)\r\nAND ProcessingStatusID \x3D 10\r\nGROUP BY eventid\r\nHAVING COUNT(eventid) > 9", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kREJEvents, Type = String, Dynamic = False, Default = \"SELECT DISTINCT eventid\rFROM eventlog\rWHERE eventid IN (\?)\rAND ProcessingStatusID \x3D 5\rGROUP BY eventid\rHAVING COUNT(eventid) > 9", Scope = Public
	#tag EndConstant

	#tag Constant, Name = kSelectEvents, Type = , Dynamic = False, Default = \"SELECT s.EventID\r\nFROM Schedule AS s\r\nINNER JOIN eventlog AS e ON e.EventID \x3D s.EventID\r\nWHERE s.FirstStartDateTimeLocal BETWEEN \? AND \?\r\nGROUP BY s.EventID\r\nHAVING count(e.eventid) > 9;\r\n", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="MailSent"
			Visible=false
			Group="Behavior"
			InitialValue="False"
			Type="Boolean"
			EditorType=""
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
