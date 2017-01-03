
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports WebfocusDLL
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.ComponentModel

Public Class Form1
    Friend WithEvents timer1 As New Timer
    Friend WithEvents timer2 As New Timer
    Friend WithEvents timer3 As New Timer

    Friend WithEvents kbHook As New KeyboardHook
    Dim CatchKeys As Boolean = True
    Dim Full_Stop As String = ""
    Dim TelnetClient As System.Net.Sockets.TcpClient
    Dim UserDocs As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\BCCopy"
    Dim LotsCSV As String = UserDocs & "\LotQtys.txt"
    Dim QtyPerCSV As String = UserDocs & "\QtyPerMold.txt"
    Dim refLots As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23wipandshopco&IBIMR_fex=pprasino/lots_by_location_fast_for_barcodeprinting.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/shop_control_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&&IBIMR_random=69890&"
    Dim refMoldQtys As String = "http://opsfocus01:8080/ibi_apps/Controller?WORP_REQUEST_TYPE=WORP_LAUNCH_CGI&IBIMR_action=MR_RUN_FEX&IBIMR_domain=qavistes/qavistes.htm&IBIMR_folder=qavistes/qavistes.htm%23routingandpa&IBIMR_fex=pprasino/pieces_per_mold.fex&IBIMR_flags=myreport%2CinfoAssist%2Creport%2Croname%3Dqavistes/mrv/part_data.fex%2CisFex%3Dtrue%2CrunPowerPoint%3Dtrue&IBIMR_sub_action=MR_MY_REPORT&WORP_MRU=true&&WORP_MPV=ab_gbv&&IBIMR_random=24365&"
    Dim LotQtys As String()
    Dim MoldQtys As String()
    Dim CurrentEvalString As String
    Dim wf As WebfocusModule

    Sub timer1_tick() Handles timer1.Tick
        CurrentEvalString = ""
    End Sub

    'Sub timer3_tick() Handles timer3.Tick
    '    TelnetClient = Nothing
    'End Sub

    Sub timer2_tick() Handles timer2.Tick

        If FileIO.FileSystem.GetFileInfo(LotsCSV).LastWriteTime.Hour <> Now.Hour Then
            timer2.Interval = 1000 * 60 * 3
            Try
                Dim t As String = UpdateDataFromSQL()
            Catch ex As Exception
                UpdateData()
            End Try
        Else
            timer2.Interval = 1000 * 60
        End If
    End Sub



    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        UpdateDataFromSQL()

        Me.Visible = False
        Me.Hide()
        NotifyIcon1.Text = "BCCopier"
        If Not FileIO.FileSystem.DirectoryExists(UserDocs) Then FileIO.FileSystem.CreateDirectory(UserDocs)
        If Not FileIO.FileSystem.FileExists(UserDocs & "\Update.bat") Then FileIO.FileSystem.CopyFile("\\slfs01\shared\prasinos\mods\windowsforms\BCPrintwithinterceptUpdater\Update.bat", UserDocs & "\Update.bat")
        'If FileIO.FileSystem.FileExists("C:\users\DataCollSL\documents") Then UserDocs = "C:\users\DataCollSL\documents"
        'LotsCSV = UserDocs & "\LotQtys.txt"
        'QtyPerCSV = UserDocs & "\QtyPerMold.txt"


        If Not FileIO.FileSystem.FileExists(LotsCSV) Or Not FileIO.FileSystem.FileExists(QtyPerCSV) Then
            Try
                Dim t As String = UpdateDataFromSQL()
            Catch ex As Exception
                UpdateData()
            End Try
        Else
            LotQtys = Split(File.ReadAllText(LotsCSV), vbCr)
            MoldQtys = Split(File.ReadAllText(QtyPerCSV), vbCr)
        End If
        timer1.Interval = 500
        timer2.Interval = 1000 * 60
        timer2.Start()
        ' PrintBarcodes(Environment.UserName & Hour(Now) & Minute(Now), True)
    End Sub

    Private Function UpdateDataFromSQL() As String

        Dim cn As New SqlClient.SqlConnection("Server=SLREPORT01; Database=WFLocal; User Id=PrasinosApps; Password=Wyman123-;")
        Dim cmd As New SqlClient.SqlCommand
        cmd.Connection = cn
        cmd.CommandText = "
        Select a.PARTNO, WORKORDERNO, QTY, ISNULL(b.PIECES_PER_MOLD, 0) AS PIECES_PER_MOLD
        From wflocal..CERT_ERRORS a 
        Left Join wflocal..ALLOYS b
        On a.PARTNO=b.Partno
        WHERE ACTIVE<>0 AND OPERATION<>9999 AND OPERATION<>10000 AND MILESTONE < 4
        ORDER BY WORKORDERNO"
        Dim lotlist As String = ""
        Dim moldlist As String = ""
        Try
            cn.Open()
            Using dr As SqlClient.SqlDataReader = cmd.ExecuteReader
                While dr.Read

                    lotlist = lotlist & Replace(dr("WORKORDERNO").ToString & ";" & dr("PARTNO").ToString & ";" & dr("QTY").ToString & ";", " ", "") & vbCrLf
                    If InStr(moldlist, Replace(dr("PARTNO").ToString, " ", "")) = 0 Then
                        moldlist = moldlist & Replace(dr("PARTNO").ToString & ";" & dr("PIECES_PER_MOLD").ToString & ";", " ", "") & vbCrLf
                    End If
                End While
            End Using
        Catch ex As Exception : Finally
            cn.Close()

        End Try
        FileIO.FileSystem.WriteAllText(QtyPerCSV, moldlist, False, System.Text.Encoding.UTF8)
        FileIO.FileSystem.WriteAllText(LotsCSV, lotlist, False, System.Text.Encoding.UTF8)
    End Function

    Sub UpdateData()
        Dim LoginInfo() As String

        Try
            NotifyIcon1.BalloonTipText = "Updating lot Lists, Printing disabled temporarily"
            NotifyIcon1.ShowBalloonTip(20)

            If IsNothing(wf) Then
                wf = New WebfocusModule
                wf.LogIn("PPRASINOS", "Wyman123-")
                LoginInfo = GetUserPasswordandFex()
                Do Until wf.IsLoggedIn
                    LogInInfo = GetUserPasswordandFex()
                    wf.LogIn(LogInInfo(0), LogInInfo(1))
                Loop

                refLots = Replace(refLots, "&IBIMR_sub_action=MR_MY_REPORT", LoginInfo(2))
                refMoldQtys = Replace(refMoldQtys, "&IBIMR_sub_action=MR_MY_REPORT", LoginInfo(2))
                wf.GetReporthAsync(refLots, "1")
                wf.GetReporthAsync(refMoldQtys, "2")
            ElseIf wf.IsDone("1") And wf.IsDone("2") Then
                Dim re As Object = wf.GetAllResponses
                SaveFlatFiles(re)
                wf = Nothing
                NotifyIcon1.BalloonTipText = "lot Lists created, Function() resumed"
                NotifyIcon1.ShowBalloonTip(10)
                LotQtys = Split(File.ReadAllText(LotsCSV), vbCr)
                MoldQtys = Split(File.ReadAllText(QtyPerCSV), vbCr)
                NotifyIcon1.BalloonTipText = "lot Lists created, Function() resumed"
                NotifyIcon1.ShowBalloonTip(10)
            Else
                NotifyIcon1.BalloonTipText = "Updating lot Lists"
                NotifyIcon1.ShowBalloonTip(20)
            End If
        Catch ex As Exception
            NotifyIcon1.BalloonTipText = "Failed To update database. Default values being used"
            NotifyIcon1.ShowBalloonTip(20)
            '  MsgBox(ex.InnerException)
            ' MsgBox(ex.Message)
            '  MsgBox(ex.ToString)

            wf = Nothing
        End Try
    End Sub

    Public Shared Function GetUserPasswordandFex() As String()
        Dim h As New Random

        Dim Usernames() As String = {"hfaizi", "mreyes", "MALMARAZ", "MARJMAND", "HYANG", "GWONG", "VDELACRUZ", "JTIBAYAN", "JSOLIS", "ASINGH", "GREYES", "JPIMENTEL", "TOSULLIVAN", "MMARTIN", "VLOPEZ", "SLI", "JIMPERIAL", "JHERNANDEZ", "FHARO", "CGOUTAMA", "HGOMEZ", "EGONZALEZ", "CDAROSA"}

        Dim y As Integer = h.Next(0, Usernames.Length)
        Dim ps As String

        Dim FexAdd As String = "&IBIMR_sub_action=MR_MY_REPORT"
        If Usernames(y) <> "pprasinos" Then
            FexAdd = "&IBIMR_sub_action=MR_MY_REPORT&IBIMR_proxy_id=pprasino.htm&"
            ps = ChrW(112) & ChrW(97) & ChrW(115) & ChrW(115) & ChrW(50) & ChrW(48) & ChrW(49) & ChrW(53)
        Else
            ps = ChrW(87) & ChrW(121) & ChrW(109) & ChrW(97) & ChrW(110) & ChrW(49) & ChrW(50) & ChrW(51) & ChrW(45)
        End If
        Debug.Print(Usernames(y))
        Return {Usernames(y), ps, FexAdd}

    End Function


    Private Sub kbHook_KeyDown(ByVal Key As System.Windows.Forms.Keys) Handles kbHook.KeyDown
        If CatchKeys = False Then Exit Sub
        timer1.Stop()
        ' """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""


        ' """"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
        CurrentEvalString = CurrentEvalString & Key.ToString
        Debug.Print(CurrentEvalString)
        Dim lotno As String
        If InStr(CurrentEvalString, "LShiftKeyD4") > 0 Then
            If InStr(CurrentEvalString, "MD") > 0 Then
                If Mid(CurrentEvalString, 1, 3) <> "LSh" Then CurrentEvalString = "LShiftKey" & CurrentEvalString
                lotno = CurrentEvalString
                CurrentEvalString = ""

                lotno = Mid(lotno, InStr(lotno, "LShiftKeyM") + 9, InStr(lotno, "LShiftKeyD4") - 10)
                lotno = Replace(lotno, "LShiftKeyQ", "Q")
                lotno = Replace(lotno, "LShiftKeyT", "T")

                If InStr(lotno, "LShiftKeyD4") > 0 Then lotno = Mid(lotno, 1, InStr(lotno, "LS") - 1)
                If InStr(lotno, "LS") > 0 Then lotno = Mid(lotno, 1, InStr(lotno, "LS") - 1)

                lotno = Replace(lotno, "OemMinus", "-")
                lotno = Replace(lotno, "OemPeriod", ".")

                lotno = Replace(lotno, "LShiftKey", "")
                lotno = Replace(lotno, "ReturnDown", "")
                lotno = Replace(lotno, "MM", "M")
                lotno = Replace(lotno, "D", "")


                CurrentEvalString = lotno

                If InStr(lotno, "-") > 0 And Mid(lotno, 2, 2) = "10" Then
                    Debug.Print(lotno)
                    PrintBarcodes(lotno)
                End If
            End If
        End If
        timer1.Start()
    End Sub

    Public Sub ConnectCallBack()

    End Sub

    Public Sub PrintBarcodes(lotno As String, Optional ScrollOverride As Boolean = False)
        CatchKeys = False

        If My.Computer.Keyboard.ScrollLock Or ScrollOverride Then
            ' If IsNothing(TelnetClient) Then 
            Try
                If IsNothing(TelnetClient) Then
                    TelnetClient = New TcpClient
                    TelnetClient.Client = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                    TelnetClient.Client.ReceiveTimeout = 100
                    TelnetClient.Client.SendTimeout = 100

                    Dim IPaddr As System.Net.IPAddress = Dns.GetHostAddresses("10.60.3.149")(0)

                    TelnetClient.Client.Connect("10.60.3.149", 9100) ', New AsyncCallback(AddressOf ConnectCallBack), TelnetClient)
                    ' Threading.Thread.Sleep(500)
                    If Not TelnetClient.Client.Connected Then
                        MsgBox("Not Connected")
                        GoTo SKIPTHIS
                    End If
                End If

                Dim NumPrints As Integer
                Dim QPM As Integer = GetQtyPerMold(lotno)
                Dim QPL As Integer = GetQtyInLot(lotno)
                Console.WriteLine("Read LotNo: " & lotno)
                Console.WriteLine(QPM & "<" & QPL)
                If QPM * QPL <> 0 Then
                    NumPrints = QPL / QPM
                    If QPL Mod QPM <> 0 Then NumPrints = NumPrints + 1
                    'Console.WriteLine("Printing " & NumPrints & " labels")
                Else
                    'Console.WriteLine("Mold data not found. Defaulting to 3 labels.")
                    NumPrints = 3
                End If
                If NumPrints > 9 Then NumPrints = 9
                If lotno = My.Computer.Name Then NumPrints = 1
                For X = 1 To NumPrints

                    Dim StringToSend As String = "^LPI;8" & Chr(10) & Chr(13) &
                "^CREATE;LOTIDLBL;72" & Chr(10) & Chr(13) &
                "VERT" & Chr(10) & Chr(13) &
                "1;38;1;6" & Chr(10) & Chr(13) &
                "STOP" & Chr(10) & Chr(13) &
                "ALPHA" & Chr(10) & Chr(13) &
                "AF1;18;5;9;1;1" & Chr(10) & Chr(13) &
                "AF2;18;5;43;1;1" & Chr(10) & Chr(13) &
                "STOP" & Chr(10) & Chr(13) &
                "BARCODE" & Chr(10) & Chr(13) &
                "C3/9;X1D;H7;BF1;18;DARK;1.5;9" & Chr(10) & Chr(13) &
                "STOP" & Chr(10) & Chr(13) &
                "BARCODE" & Chr(10) & Chr(13) &
                "C3/9;X1D;H7;BF2;18;DARK;1.5;43" & Chr(10) & Chr(13) &
                "STOP" & Chr(10) & Chr(13) &
                "END" & Chr(10) & Chr(13) &
                "^EXECUTE;LOTIDLBL" & Chr(10) & Chr(13) &
                "^BF1;" & Chr(34) & "" & lotno & "$" & Chr(34) & Chr(10) & Chr(13) &
                "^BF2;" & Chr(34) & "" & lotno & "$" & Chr(34) & Chr(10) & Chr(13) &
                "^AF1;" & Chr(34) & Replace(Replace(lotno, "$", ""), "M", "", , 1) & Chr(34) & Chr(10) & Chr(13) &
                "^AF2;" & Chr(34) & Replace(Replace(lotno, "$", ""), "M", "", , 1) & Chr(34) & Chr(10) & Chr(13) &
                "^NORMAL" & Chr(10) & Chr(13)


                    Send_Sub(StringToSend)

                Next X
                Dim shoplot As String = Replace(Replace(lotno, "$", ""), "M10", "", , 1)
                Dim lot As String = Split(shoplot, "-")(1)

                lot = "-" & lot
                Dim BigStringToSend As String = "^LPI;8" & Chr(10) & Chr(13) &
        "^CREATE;LOTIDLBL;72" & Chr(10) & Chr(13) &
        "ALPHA" & Chr(10) & Chr(13) &
        "AF1;18;3;54;2;2" & Chr(10) & Chr(13) &
        "AF2;18;5;53;2;2" & Chr(10) & Chr(13) &
        "STOP" & Chr(10) & Chr(13) &
        "BARCODE" & Chr(10) & Chr(13) &
       "C3/9;X1B;H8;BF1;18;DARK;1.5;9" & Chr(10) & Chr(13) &
        "STOP" & Chr(10) & Chr(13) &
        "END" & Chr(10) & Chr(13) &
        "^EXECUTE;LOTIDLBL" & Chr(10) & Chr(13) &
        "^BF1;" & Chr(34) & "" & lotno & "$" & Chr(34) & Chr(10) & Chr(13) &
        "^AF2;" & Chr(34) & lot & Chr(34) & Chr(10) & Chr(13) &
   "^AF1;" & Chr(34) & Split(shoplot, "-")(0) & Chr(34) & Chr(10) & Chr(13) &
"^NORMAL" & Chr(10) & Chr(13)
                Send_Sub(BigStringToSend)
                NotifyIcon1.Visible = True
                NotifyIcon1.Text = "Printer Active"
                NotifyIcon1.BalloonTipText = lotno & " Printed (" & NumPrints & " Copies)" & Chr(10) & "Switch Scroll Lock off to Disable Printing"
                NotifyIcon1.ShowBalloonTip(6)

                '  TelnetClient.Close()M1012407-00037$
SKIPTHIS:
            Catch
            Finally
                CatchKeys = True
                Try
                    TelnetClient.Client.Disconnect(True)
                    TelnetClient.Client.Dispose()
                Catch : Finally
                    TelnetClient = Nothing
                End Try
            End Try

        Else

            NotifyIcon1.Visible = True
                NotifyIcon1.Text = "Printer Disabled"
            End If

    End Sub

    Private Function GetQtyPerMold(LotNo As String) As Integer
        Dim Part As String = Split(LotNo, "-")(0)
        Part = Replace(Part, "M10", "")

        For Each record In MoldQtys
            If InStr(record, Part) Then
                Return CInt(Split(record, ";")(1))
            End If
        Next
        Return 0
    End Function



    Private Function GetQtyInLot(LotNo As String) As Integer
        Dim sn As String = Replace(Replace(LotNo, "$", ""), "M", "",, 1)
        For Each record In LotQtys

            If InStr(record, sn) > 0 Then
                Return CInt(Split(record, ";")(2))
            End If
        Next
        Return 0
    End Function

    Sub Send_Sub(ByVal msg As String)

        Dim byt_to_send() As Byte = System.Text.Encoding.ASCII.GetBytes(msg)
            Dim s As Integer = TelnetClient.Client.Send(byt_to_send, 0, byt_to_send.Length, SocketFlags.None)
        Threading.Thread.Sleep(10)

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.WindowsShutDown Or e.CloseReason = CloseReason.TaskManagerClosing Then
            e.Cancel = False
        ElseIf CurrentEvalString = "1" Then
            e.Cancel = False
        Else
            e.Cancel = True
        End If
        If e.Cancel = False Then
            ' TelnetClient.Client.Close()
            ' TelnetClient.Close()
            ' Threading.Thread.Sleep(1000)
            TelnetClient = Nothing
        End If

    End Sub

    Private Sub EXITToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EXITToolStripMenuItem.Click
        CurrentEvalString = "1"
        Me.Close()

    End Sub


    Private Function SaveFlatFiles(WfResponses As List(Of AsyncResponse)) As Boolean
        Dim SavePath As String


        Dim j()() As String = WfResponses(0).Response

        SavePath = LotsCSV
        If FileIO.FileSystem.FileExists(LotsCSV) Then FileIO.FileSystem.DeleteFile(LotsCSV)
        If FileIO.FileSystem.FileExists(QtyPerCSV) Then FileIO.FileSystem.DeleteFile(QtyPerCSV)
        Dim fs As New StreamWriter(LotsCSV, True)



        For x = 1 To j.Length - 1
            Dim sr As String = ""
            For y = 0 To j(x).Length - 1
                sr = sr & j(x)(y) & ";"
            Next y
            fs.WriteLine(sr)
        Next x
        fs.Flush()
        fs.Close()

        Dim w()() As String = WfResponses(1).Response
        Dim fs1 As New StreamWriter(QtyPerCSV, True)

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

    Private Sub UpdateLotListsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UpdateLotListsToolStripMenuItem.Click
        updatedata()
    End Sub



End Class

