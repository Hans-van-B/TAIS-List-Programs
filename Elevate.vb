Imports System.Security.Principal

Module Elevate

    Sub ElevateMe()
        xtrace_subs("ElevateMe")

        If IsAdmin() Then
            xtrace_i("Already elevated")
            Form1.StartInstall()
        Else
            xtrace_i("Need to elevate")

            Dim Exe As String = Application.ExecutablePath
            Dim Exe2 = Exe.Replace("""", "")
            xtrace_i("Exe2 = " & Exe2)

            Dim Param As String
            'Param = Mid(Environment.CommandLine, Len(Exe) + 1)
            Param = "--seq=2"
            For Each Arg As String In My.Application.CommandLineArgs
                If Left(Arg, 5) = "--seq" Then Continue For
                Param = Param & " " & Arg
            Next
            xtrace_i("Param = " & Param)

            xtrace_i("ReStart Me Elevated")
            Try
                Dim Proc As New Process()
                Dim ProcStartInfo As New ProcessStartInfo(Exe2)
                ProcStartInfo.Verb = "runas"
                ProcStartInfo.Arguments = Param
                Proc.StartInfo = ProcStartInfo
                Proc.Start()

            Catch ex As Exception
                xtrace_err(ex.Message)
            End Try

        End If

        xtrace_sube("ElevateMe")
    End Sub

    Function IsAdmin() As Boolean
        xtrace_subs("IsAdmin")

        Dim Identity As WindowsIdentity = WindowsIdentity.GetCurrent()
        Dim Principal As WindowsPrincipal = New WindowsPrincipal(Identity)
        Dim Result As Boolean = Principal.IsInRole(WindowsBuiltInRole.Administrator)

        xtrace_i("IsAdmin = " & Result.ToString)
        IsAdmin = Result

        xtrace_sube("IsAdmin")
    End Function


End Module
