#Region " Imports "

Imports System.IO
Imports System.Text
Imports System.Configuration
Imports K4GDW.RandomSig.Signature

#End Region

Module modMain

	''' <summary>
	''' Local storage for the NotifyIcon
	''' </summary>
'	Private mobNotifyIcon As NotifyIcon

	''' <summary>
	''' Local storage for the ContextManu object.
	''' </summary>
	Private WithEvents mobContextMenu As ContextMenu

	''' <summary>
	''' Local storage for the Timer object.
	''' </summary>
	Private WithEvents mobTimer As Timers.Timer

	''' <summary>
	''' Sets up timer.
	''' </summary>
	Private Sub SetUpTimer()
		Try
			mobTimer = New Timers.Timer()
			With mobTimer
				.AutoReset = True
				.Interval = CInt(ConfigurationManager.AppSettings("WaitPeriod")) * 1000
				.Start()
			End With
		Catch obEx As Exception
			Throw
		End Try
	End Sub

    ''' <summary>
    ''' This function provides a way to check for existing
    ''' instance of the program.
    ''' </summary>
    ''' <returns>
    ''' A boolean indicating wether or not the
    ''' program is alread running</returns>
    ''' <remarks>
    ''' I got this code from the VBHelper email I get once a week from
    ''' VBHelper.com on 5/12/2007
    ''' </remarks>
    Private Function AlreadyRunning() As Boolean
		' Get our process name.
		Dim my_proc As Process = Process.GetCurrentProcess
		Dim my_name As String = my_proc.ProcessName

		' Get information about processes with this name.
		Dim procs() As Process = Process.GetProcessesByName(my_name)

		' If there is only one, it's us.
		If procs.Length = 1 Then Return False

		' If there is more than one process,
		' see if one has a StartTime before ours.
		Dim i As Integer
		For i = 0 To procs.Length - 1
			If procs(i).StartTime < my_proc.StartTime Then Return True
		Next i

		' If we get here, we were first.
		Return False

		'Dim isAlreadyRunning As Boolean = False
		'Dim objMutex As System.Threading.Mutex
		'objMutex = System.Threading.Mutex.OpenExisting("Application_Mutex")
		'If objMutex Is Nothing Then
		'	objMutex = New System.Threading.Mutex(True, "Application_Mutex", isAlreadyRunning)
		'End If
		'Return isAlreadyRunning
    End Function

	''' <summary>
	''' Creates the menu.
	''' </summary>
	Private Sub CreateMenu()
		Try
			mobContextMenu = New ContextMenu
			With mobContextMenu.MenuItems
                .Add(New MenuItem("Exit", _
                  New EventHandler(AddressOf Close)))
				.Add(New MenuItem("Refresh Sig", _
				  New EventHandler(AddressOf ChangeSig)))
				.Add(New MenuItem("Copy Sig", _
				 New EventHandler(AddressOf CopySig)))
				.Add(New MenuItem("Show Sig", _
				 New EventHandler(AddressOf ShowSig)))
				.Add(New MenuItem("About Signature Rotator", _
				 New EventHandler(AddressOf AboutBox)))
			End With
		Catch obEx As Exception
			Throw
		End Try
	End Sub

	''' <summary>
	''' Copies the sig.
	''' </summary>
	''' <param name="sender">The sender.</param>
	''' <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
	Private Sub CopySig(ByVal sender As Object, ByVal e As EventArgs)
		Dim i As Integer
		Dim cs As Signature = ChooseSig(i)
		Dim sb As New StringBuilder
		With sb
			.Append(cs.Preamble)
			.Append(cs.Text)
			.Append(cs.Boilerplate)
		End With
		Clipboard.SetText(sb.ToString, TextDataFormat.Text)
		Sigs.RemoveAt(i)
	End Sub

	''' <summary>
	''' Displays the about box for the program.
	''' </summary>
	''' <param name="sender">The sender.</param>
	''' <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
	Private Sub AboutBox(ByVal sender As Object, ByVal e As EventArgs)
		Dim sb As New StringBuilder
		With sb
			.Append("Thanks for trying out my signature rotator utility.")
			.Append(vbCrLf)
			.Append("Version;")
			.Append(vbTab)
			.Append(My.Application.Info.Version)
			.Append(vbCrLf)
			.Append(My.Application.Info.Copyright)
			.Append(vbCrLf)
			.Append(My.Application.Info.CompanyName)
			.Append(vbCrLf)
			.Append(My.Application.Info.Description)
		End With
		MsgBox(sb.ToString, MsgBoxStyle.Information, My.Application.Info.ProductName)
	End Sub

	''' <summary>
	''' Shows the sig.
	''' </summary>
	''' <param name="sender">The sender.</param>
	''' <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
	Private Sub ShowSig(ByVal sender As Object, ByVal e As EventArgs)
		Dim i As Integer
		Dim cs As Signature = ChooseSig(i)
		MsgBox("Signature #:  " & (i + 1).ToString & " of " & Sigs.Count.ToString & vbCrLf & vbCrLf & cs.Preamble & cs.Text & cs.Boilerplate)
		Sigs.RemoveAt(i)
	End Sub

	''' <summary>
	''' Changes the sig.
	''' </summary>
	''' <param name="sender">The sender.</param>
	''' <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
	Private Sub ChangeSig(ByVal sender As Object, ByVal e As EventArgs)
		WriteSigFile(ChooseSig)
	End Sub

	''' <summary>
	''' Closes the specified sender.
	''' </summary>
	''' <param name="sender">The sender.</param>
	''' <param name="e">The <see cref="System.EventArgs" /> instance containing the event data.</param>
	Private Sub Close(ByVal sender As Object, ByVal e As EventArgs)
		Close()
	End Sub

	''' <summary>
	''' Closes this instance.
	''' </summary>
	Private Sub Close()
		With mobNotifyIcon
			.Visible = False
			.Dispose()
		End With
		With mobContextMenu
			.Dispose()
		End With
		With mobTimer
			.Dispose()
		End With
		Application.Exit()
	End Sub

	''' <summary>
	''' Handles the Elapsed event of the mobTimer control.
	''' </summary>
	''' <param name="sender">The source of the event.</param>
	''' <param name="e">The <see cref="System.Timers.ElapsedEventArgs" /> instance containing the event data.</param>
    Public Sub mobTimer_Elapsed(ByVal sender As Object, _
       ByVal e As Timers.ElapsedEventArgs) Handles mobTimer.Elapsed
        Try
            Dim i As Integer
            WriteSigFile(ChooseSig(i))
            Sigs.RemoveAt(i)
            Sigs.TrimExcess()
        Catch obEx As Exception
            LogError(obEx)
            Close()
        End Try
    End Sub

	''' <summary>
	''' Logs errors to the windows Application event log.
	''' </summary>
	''' <param name="ex">
	''' A reference to the exception object.
	''' </param>
	''' <remarks>
	''' Writen by bjohns - 5/10/2007 1:18:57 PM
	''' </remarks>
	Private Sub LogError(ByRef ex As Exception)
		My.Application.Log.WriteEntry(ex.ToString, TraceEventType.Error)
	End Sub

	''' <summary>
	''' Logs the error.
	''' </summary>
	''' <param name="msg">The MSG.</param>
	Private Sub LogError(ByVal msg As String)
		LogError(New Exception(msg))
	End Sub

	''' <summary>
	''' Logs the error.
	''' </summary>
	''' <param name="msg">The MSG.</param>
	''' <param name="ex">The ex.</param>
	Private Sub LogError(ByVal msg As String, ByVal ex As Exception)
		LogError(New Exception(msg, ex))
	End Sub

	''' <summary>
	''' This method writes the sig to the outlook (or other mail app) sig file.
	''' </summary>
	''' <remarks>
	''' </remarks>
    Private Sub WriteSigFile(ByRef sig As Signature)
        Dim txtFile As String = SigFile
        Dim htmlFile As String = SigFile.Replace("txt", "htm")
        Dim sb As New StringBuilder
        sb.Append(sig.Preamble)
        sb.Append(sig.Text)
        If Not sig.Boilerplate = String.Empty Then
            sb.Append(sig.Boilerplate)
        End If
        sb.Append(vbCrLf)
        sb.Append(vbCrLf)
        Dim signature As String = sb.ToString
        sb.Insert(0, "<html><body><p>")
        sb.Replace(vbCrLf, "<br />")
        sb.Append("</p><p>&nbsp;</p></body></html>")
        Dim htmlSig As String = sb.ToString

        Using txtFileWriter As New StreamWriter(txtFile)
            Try
                txtFileWriter.Write(signature)
            Catch ex As IOException
                LogError("Unable to write to plain text signature file.", ex)
            Catch ex As Exception
                Throw
            Finally
                txtFileWriter.Close()
            End Try
        End Using

        Using htmlFileWriter As New StreamWriter(htmlFile)
            Try
                htmlFileWriter.Write(htmlSig)
            Catch ex As IOException
                LogError("Unable to write to html signature file.", ex)
            Catch ex As Exception
                Throw
            Finally
                htmlFileWriter.Close()
            End Try
        End Using

    End Sub

	''' <summary>
	''' Local storage for the application Icon
	''' </summary>
    Private icon As Icon

    ''' <summary>
    ''' The main routine.
    ''' </summary>
    Sub Main()
        My.Application.Log.WriteEntry("Starting Signature Rotator", TraceEventType.Start)
		icon = New Icon(My.Application.Info.DirectoryPath & "\gdw16.ico")
        mobNotifyIcon = New NotifyIcon()
        mobNotifyIcon.Icon = Icon
        mobNotifyIcon.Visible = False
        'mobContextMenu = New ContextMenu()
        ' See if another instance is already running.

        If AlreadyRunning() Then
            LogError("An attempt was made to start the application twice.")
            MsgBox("There is already an instance of the application running.", MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly)
            Close()
        End If

        CreateMenu()
        mobNotifyIcon.ContextMenu = mobContextMenu
        SetUpTimer()
        mobNotifyIcon.Visible = True
        Try
            SigFile = ConfigurationManager.AppSettings("SigfileLocation")
            QuoteFile = My.Application.Info.DirectoryPath & "\" & ConfigurationManager.AppSettings("quotefile")
			Sigs = New List(Of Signature)
            GetSigs(Sigs)
            Dim i As Integer
            WriteSigFile(ChooseSig(i))
            Sigs.RemoveAt(i)
            Sigs.TrimExcess()
            Application.Run()
        Catch ex As Exception
#If DEBUG Then
            MsgBox(ex.Message & vbCrLf & ex.ToString, MsgBoxStyle.Critical)
#End If
            LogError(ex)
            Close()
        End Try
    End Sub

End Module
