Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop
Imports webfocusmodule.dll

Namespace My

    ' The following events are available for MyApplication: 
    ' 
    ' Startup: Raised when the application starts, before the startup form is created. 
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally. 
    ' UnhandledException: Raised if the application encounters an unhandled exception. 
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected. 

    Partial Friend Class MyApplication

        Private Sub MyApplication_Startup(sender As Object, e As ApplicationServices.StartupEventArgs) Handles Me.Startup

            Dim UserDocs As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\BCCopy"
            Dim LotsCSV As String = UserDocs & "\LotQtys.txt"
            Dim QtyPerCSV As String = UserDocs & "\QtyPerMold.txt"
            If Not FileIO.FileSystem.DirectoryExists(UserDocs) Then FileIO.FileSystem.CreateDirectory(UserDocs)
            If Not FileIO.FileSystem.FileExists(LotsCSV) Or Not FileIO.FileSystem.FileExists(QtyPerCSV) Then UpdateData()
            AddHandler Microsoft.Win32.SystemEvents.SessionEnding, AddressOf Handler_SessionEnding
        End Sub

        Public Sub Handler_SessionEnding(ByVal sender As Object, _
               ByVal e As Microsoft.Win32.SessionEndingEventArgs)
            If e.Reason = Microsoft.Win32.SessionEndReasons.Logoff Then
                Form1.Close()
            ElseIf e.Reason = Microsoft.Win32.SessionEndReasons.SystemShutdown Then
                Form1.Close()
            End If
        End Sub


        Sub UpdateData()
            Dim logininfo() As String

            Dim refLots As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23wipandshopco&IBIMR_fex=pprasino/lots_by_location_fast_for_barcodeprinting.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/shop_control_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&&IBIMR_random=69890&"
            Dim refMoldQtys As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23routingandpa&IBIMR_fex=pprasino/pieces_per_mold.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/part_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&&IBIMR_random=24365&"

            Dim wflist As New List(Of String()())
            Dim wf As New WebfocusDLL.WebfocusModule
            Do Until wf.IsLoggedIn
                logininfo = Form1.GetUserPasswordandFex()
                wf.LogIn(logininfo(0), logininfo(1))
            Loop

            refLots = Replace(refLots, "&IBIMR_sub_action=MR_MY_REPORT", logininfo(2))
            refMoldQtys = Replace(refMoldQtys, "&IBIMR_sub_action=MR_MY_REPORT", logininfo(2))

            wflist.Add(wf.GetReporth(refLots))
            wflist.Add(wf.GetReporth(refMoldQtys))
            SaveFlatFiles(wflist)
            wf = Nothing

        End Sub

        Private Function SaveFlatFiles(WfResponses As List(Of String()())) As Boolean
            Dim UserDocs As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\BCCopy"
            Dim LotsCSV As String = UserDocs & "\LotQtys.txt"
            Dim QtyPerCSV As String = UserDocs & "\QtyPerMold.txt"
            Dim SavePath As String

            Dim j()() As String = WfResponses(0)

            SavePath = LotsCSV
            If FileIO.FileSystem.FileExists(LotsCSV) Then FileIO.FileSystem.DeleteFile(LotsCSV)
            If FileIO.FileSystem.FileExists(QtyPerCSV) Then FileIO.FileSystem.DeleteFile(QtyPerCSV)
            Dim fs As New System.IO.StreamWriter(LotsCSV, True)

            For x = 1 To j.Length - 1
                Dim sr As String = ""
                For y = 0 To j(x).Length - 1
                    sr = sr & j(x)(y) & ";"
                Next y
                fs.WriteLine(sr)
            Next x
            fs.Flush()
            fs.Close()

            Dim w()() As String = WfResponses(1)
            Dim fs1 As New System.IO.StreamWriter(QtyPerCSV, True)

            For x = 1 To w.Length - 1
                Dim sr1 As String = ""
                For y = 0 To w(x).Length - 1
                    sr1 = sr1 & w(x)(y) & ";"
                Next y
                fs1.WriteLine(sr1)
            Next x
            fs1.Flush()
            fs1.Close()
            Return True
        End Function

        '=Microsoft.ACE.OLEDB.12.0;Data Source=C:\users\cdarosa\documents\SalesTrackData\SalesInfo.accdb

        Private Sub MyApplication_UnhandledException(sender As Object, e As ApplicationServices.UnhandledExceptionEventArgs) Handles Me.UnhandledException
            EmailFile(e)

        End Sub



        Private Function EmailFile(ex As ApplicationServices.UnhandledExceptionEventArgs)

            Dim OutLookApp As New Outlook.Application
            Dim Mail As Outlook.MailItem = OutLookApp.CreateItem(OlItemType.olMailItem)
            Dim mailRecipient As Outlook.Recipient
            Dim address As String = "pprasinos" & "@pccstructurals.com"
            mailRecipient = Mail.Recipients.Add(address)
            mailRecipient.Resolve()
            If Not mailRecipient.Resolved Then MsgBox(address)

            Mail.Recipients.ResolveAll()
            Mail.HTMLBody = ex.Exception.InnerException.ToString & Chr(34) & vbCr & vbCrLf & ex.Exception.GetBaseException.ToString & vbCr & vbCr & vbCr & ex.Exception.Message & vbCr & vbCr & vbCrLf & ex.Exception.Source
            Mail.Subject = "EXCEPTION MAILER: " & ex.Exception.TargetSite.Module.ToString
            Mail.Save()
            Mail.Send()

        End Function


    End Class


    

End Namespace
