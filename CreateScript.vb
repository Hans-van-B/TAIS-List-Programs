Module CreateScript
    Public ProgList As String
    Sub CreatePSScript(TName As String)
        xtrace_subs("CreatePSScript")

        If TName = "Inst" Then PS_Script_Target = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) & "\TAIS\Util"
        If TName = "Run" Then PS_Script_Target = Temp

        If Not My.Computer.FileSystem.DirectoryExists(PS_Script_Target) Then
            My.Computer.FileSystem.CreateDirectory(PS_Script_Target)
        End If
        ScriptPath = PS_Script_Target & "\" & PSName & ".ps1"
        xtrace_i("ScriptPath = " & ScriptPath)

        CS_WriteHeader()

        ProgList = """" & Temp & "\InstalledProgramsList.txt" & """"
        CS_WriteLine("$list = " & ProgList)

        Dim Content As String = "Write-Host ""Output to: "", $list

Write-Output ""File: $list"", "" "" >$list
Write-Output ""---- AppxPackage ----------"" >>$list

Get-AppxPackage * | select name, InstallLocation >>$list

Write-Output "" "", ""---- Programs ----------"" >>$list

Get-WmiObject -Class Win32_Product | Select-Object -Property Name, Vendor  >>$list

Write-Host ""Start NotePad"""
        CS_WriteLine(Content)
        CS_WriteLine("& """ & NotePad & """ " & ProgList)

        xtrace_sube("CreatePSScript")
    End Sub

    Sub CS_WriteHeader()
        My.Computer.FileSystem.WriteAllText(ScriptPath, "# " & ScriptPath & vbNewLine, False)
        CS_WriteLine("# Created: " & DateTime.Now)
    End Sub

    Sub CS_WriteLine(Text As String)
        My.Computer.FileSystem.WriteAllText(ScriptPath, Text & vbNewLine, True)
    End Sub

End Module
