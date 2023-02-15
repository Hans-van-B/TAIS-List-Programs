Module Log
    Dim SubLevel As Integer = 0
    Dim LTrace As Integer = 2
    Dim G_TL_Sub As Integer = 2  ' Trace level for Subroutines
    Public LogFile As String

    Sub xtrace_init()
        InitTemp()

        ' Show the command-line string at the top of the log file
        My.Computer.FileSystem.WriteAllText(LogFile, Environment.CommandLine & vbNewLine, False)
        My.Computer.FileSystem.WriteAllText(LogFile, "xtrace_init" & vbNewLine, True)
    End Sub

    Sub xtrace_header()
        xtrace_subs("xtrace_header")
        xtrace_i(AppName & " V" & AppVer)
        xtrace_TimeStamp()

        xtrace_i("AppRoot = " & AppRoot)
        xtrace_i("Log level to logfile = " & LTrace.ToString, 2)
        xtrace_sube("xtrace_header")
    End Sub

    '---- xtrace ----
    Sub xtrace_root(Msg As String, TV As Integer)
        If LogFile Is Nothing Then Return
        Dim Nr As Int16
        Dim Tab As String = ""

        ' If subroutines are Not logged then tabbing is also disabeled
        If LTrace >= G_TL_Sub Then
            ' If exiting main or sublevel maint error,
            ' this if makes it more clear
            If SubLevel >= 0 Then Tab = "|"

            For Nr = 1 To SubLevel
                Tab = Tab + "  |"
            Next
        End If

        If TV <= LTrace Then
            My.Computer.FileSystem.WriteAllText(LogFile, Tab & Msg & vbNewLine, True)
        End If
    End Sub
    Sub xtrace(Msg As String)
        xtrace_root(" " & Msg, 2)
    End Sub

    Sub xtrace(Msg As String, TV As Integer)
        xtrace_root(" " & Msg, TV)
    End Sub

    '---- xtrace_i ----
    Sub xtrace_i(Msg As String)
        xtrace(" * " & Msg)
    End Sub

    Sub xtrace_i(Msg As String, TV As Integer)
        xtrace(" * " & Msg, TV)
    End Sub

    '--- xtrace TimeStamp ----
    Sub xtrace_TimeStamp()
        xtrace_i("Timestamp = " & DateTime.Now)
    End Sub

    '--- xtrace Sub ----
    Sub xtrace_subs(Msg As String)
        xtrace_root("->" & Msg & " (" & (SubLevel + 1).ToString & ")", G_TL_Sub)
        SubLevel = SubLevel + 1
    End Sub

    Sub xtrace_sube(Msg As String)
        SubLevel = SubLevel - 1
        xtrace_root("<-" & Msg & " (" & (SubLevel + 1).ToString & ")", G_TL_Sub)
    End Sub

    Sub xtrace_err(Msg As String)
        xtrace("ERROR: " & Msg)
    End Sub

    Public Sub WriteInfo(Msg)
        Form1.TextBoxInfo.AppendText(Msg & vbNewLine)
        xtrace(Msg)
    End Sub

End Module
