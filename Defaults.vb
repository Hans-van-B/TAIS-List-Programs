Module Defaults
    Dim Err1 As Boolean = False
    Sub ReadDefaults()
        xtrace_subs("ReadDefaults")
        Dim ReadFile

        If My.Computer.FileSystem.FileExists(IniFile1) Then
            WriteInfo("Read Defaults " & IniFile1)
            ReadFile = My.Computer.FileSystem.OpenTextFileReader(IniFile1)

        ElseIf My.Computer.FileSystem.FileExists(IniFile2) Then
            WriteInfo("Read Defaults " & IniFile2)
            ReadFile = My.Computer.FileSystem.OpenTextFileReader(IniFile2)

        Else
            WriteInfo("Failed to read:")
            WriteInfo("   " & IniFile1)
            WriteInfo("Or " & IniFile2)
            CreateIniFile()
            Exit Sub
        End If

        Dim Line As String
        Dim Group As String = ""
        Dim P1, P2 As Integer
        Dim DName, DVal As String

        While Not ReadFile.EndOfStream
            Try
                Line = ReadFile.ReadLine()

                '---- Remark lines ----
                If Left(Line, 1) = "#" Then
                    Continue While
                End If

                '---- Read Group ----------
                P1 = InStr(Line, "[")
                P2 = InStr(Line, "]")

                If P1 = 1 And P2 > 2 Then
                    Group = Mid(Line, 2, P2 - 2)
                    xtrace("Group = " & Group)
                    Continue While
                End If

                '---- Pick Lists ----
                ' Reset group (Optional)
                If Len(Line) < 2 Then
                    Group = ""
                    Continue While
                End If

                If Group = "COMBOBOX01" Then
                    'Form1.COMBOBOX01.Items.Add(Line)
                End If

                '---- Read Defaults
                P1 = InStr(Line, "=")
                DName = Left(Line, P1 - 1)
                DVal = Mid(Line, P1 + 1)
                xtrace("Default " & DName & "=" & DVal)

                If Group = "INIT" Then

                    If DName = "WindowState" Then
                        If UCase(DVal) = "MIN" Then
                            xtrace("Minimize form")
                            Form1.WindowState = FormWindowState.Minimized
                        End If
                    End If

                    If DName = "NotePad" Then
                        NotePad = DVal
                    End If

                End If

            Catch ex As Exception

            End Try
        End While
        ReadFile.Dispose()

        xtrace_sube("ReadDefaults")
    End Sub

    Function StringToBoolean(Val As String) As Boolean
        Val = UCase(Val).Trim

        If (Val = "TRUE") _
        Or (Val = "1") _
        Or (Val = "Y") Then
            Return True
        Else
            Return False
        End If
    End Function

    '---- Create ini file -----------------------------------------------------
    Sub CreateIniFile()
        CIF_WriteHeader()
        If Err1 Then Return
        CIF_WriteLine("NotePad=C:\Program Files (x86)\Notepad++\notepad++.exe")

    End Sub

    Sub CIF_WriteHeader()
        Try
            My.Computer.FileSystem.WriteAllText(ScriptPath, "# " & ScriptPath & vbNewLine, False)
        Catch ex As Exception
            xtrace_err(ex.Message)
            Err1 = True
            Return
        End Try
        CS_WriteLine("[INIT]")
    End Sub

    Sub CIF_WriteLine(Text As String)
        My.Computer.FileSystem.WriteAllText(ScriptPath, Text & vbNewLine, True)
    End Sub
End Module
