Module RunScript
    Sub RunPS()
        xtrace_subs("RunPS")

        Dim S32 As String = Environment.GetFolderPath(Environment.SpecialFolder.System)
        xtrace_i("System32   = " & S32)
        Dim PSExe As String = S32 & "\WindowsPowerShell\v1.0\PowerShell.exe"
        Dim Arguments = "-ExecutionPolicy Bypass -File """ & ScriptPath & """"

        xtrace_i("Start: " & PSExe & " " & Arguments)
        Dim Proc As New Process()
        Dim ProcStartInfo As New ProcessStartInfo(PSExe)
        ProcStartInfo.Arguments = Arguments
        Proc.StartInfo = ProcStartInfo
        Proc.Start()
        Proc.WaitForExit()
        Proc.Close()

        'xtrace_i("Start: " & NotePad & " " & Arguments)
        'Dim ProcStartInfo2 As New ProcessStartInfo(NotePad)
        'ProcStartInfo2.Arguments = ProgList
        'Proc.StartInfo = ProcStartInfo2
        'Proc.Start()
        'Proc.WaitForExit()
        'Proc.Close()

        xtrace_sube("RunPS")
    End Sub
End Module
