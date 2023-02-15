' https://www.vbforums.com/showthread.php?586974-Create-a-shortcut-using-IWshRuntimeLibrary
'
'From the Add Reference dialog, click the .COM tab and select the 'Windows Scripting Host Object Model' and click the 'Add' button.

Imports System.IO
Imports IWshRuntimeLibrary

Module CreateIcon
    Sub CreateStartIcon()
        xtrace_subs("CreateStartIcon")

        Dim Sm As String = Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu)
        Dim S32 As String = Environment.GetFolderPath(Environment.SpecialFolder.System)
        xtrace_i("Start Menu = " & Sm)
        xtrace_i("System32   = " & S32)
        Dim PSExe As String = S32 & "\WindowsPowerShell\v1.0\PowerShell.exe"

        Dim TargetDir As String = Sm & "\TAIS\Util"
        xtrace_i("TargetDir = " & TargetDir)
        If Not My.Computer.FileSystem.DirectoryExists(TargetDir) Then
            My.Computer.FileSystem.CreateDirectory(TargetDir)
        End If

        Dim ShortcutPath As String = TargetDir & "\" & AppName2 & ".lnk"

        Dim MyWshShell As WshShell = New WshShell
        Dim MyShortcut As IWshRuntimeLibrary.IWshShortcut
        MyShortcut = CType(MyWshShell.CreateShortcut(ShortcutPath), IWshRuntimeLibrary.IWshShortcut)
        With MyShortcut
            .TargetPath = PSExe     'Specify target executable full path
            .Arguments = "-ExecutionPolicy Bypass -File """ & ScriptPath & """"
            .Description = "Technical Application Installation Service Utility: " & AppName
            .IconLocation = PS_Script_Target & "\Installer.ico"
            .Save()
        End With

        xtrace_sube("CreateStartIcon")
    End Sub
End Module
