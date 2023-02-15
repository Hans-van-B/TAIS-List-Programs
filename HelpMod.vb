Module HelpMod

    '==== GUI Help ============================================================

    Dim HelpPage As String = Temp & "\help.html"
    Dim HelpPageHtm As String = AppRoot & "\" & AppName & ".html"
    Dim HelpPagePdf As String = AppRoot & "\" & AppName & ".pdf"

    Sub ShowHelp()
        xtrace("Show Help")

        If (System.IO.File.Exists(HelpPageHtm)) Then
            HelpPage = HelpPageHtm
        ElseIf (System.IO.File.Exists(HelpPagePdf)) Then
            HelpPage = HelpPagePdf
        Else
            xtrace("Did not find " & HelpPageHtm)
            xtrace("Did not find " & HelpPagePdf)

            xtrace("Create page")
            My.Computer.FileSystem.WriteAllText(HelpPage, "<html>" & vbNewLine, False)
            WriteHelp("<head>")
            WriteHelp("")
            WriteHelp("</head>")
            WriteHelp("<H2>" & AppName & " V" & AppVer & " Help Page</H2>")
            WriteHelp("Technical Application Installation Service Utility: " & AppName & "<br>")
            WriteHelp("<br>")
            WriteHelp("This program creates a list of all installed programs and apps.<br>")
            WriteHelp("<br>")
            WriteHelp("<br>")
            WriteHelp("The " & AppName & " log can be found here: " & LogFile & "<br>")
            WriteHelp("</body>")
            WriteHelp("</html")
        End If

        Process.Start(HelpPage, "")
    End Sub

    Sub ShowHelpAbout()
        xtrace("Show Help, about")

        xtrace("Create page")
        My.Computer.FileSystem.WriteAllText(HelpPage, "<html>" & vbNewLine, False)
        WriteHelp("<head>")
        WriteHelp("")
        WriteHelp("</head>")
        WriteHelp("<H2>" & AppName & " V" & AppVer & " Help About</H2>")
        WriteHelp("Technical Application Installation Service Utility: " & AppName & "<br>")
        WriteHelp("<br>")
        WriteHelp("This release supports both starting the remove action directly<br>")
        WriteHelp("or installing it in the start menu below &quot;TAIS&quot; which stands for &quot;Technical Application Information Service&quot;.<br>")
        WriteHelp("This compiled app may be used freely for both commercial use and non commercial use without limitations.<br>")
        WriteHelp("<br>")
        WriteHelp("<hr>")
        WriteHelp("<font size=""-1"">")
        WriteHelp("This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details. You should have received a copy of the GNU General Public License along with this program. If not, see https://www.gnu.org/licenses/. <br>")
        WriteHelp("<br>")
        WriteHelp(" The " & AppName & " log can be found here: " & LogFile & "<br>")
        WriteHelp(" Dev. : C:\Cloud1\VB.Net-Dev\TAIS-Util\TAIS-List-Programs<br>")
        WriteHelp(" Maint: C:\Cloud1\Dev\TAIS\List-Programs<br>")
        WriteHelp(" Git  : https://github.com/Hans-van-B/TAIS-List-Programs <br>")
        WriteHelp(" Inst.: " & AppRoot & "<br>")
        WriteHelp("</font>")
        WriteHelp("</body>")
        WriteHelp("</html")

        Process.Start(HelpPage, "")
    End Sub

    Sub WriteHelp(Line As String)
        My.Computer.FileSystem.WriteAllText(HelpPage, Line & vbNewLine, True)
    End Sub

End Module
