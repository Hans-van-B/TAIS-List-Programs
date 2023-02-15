Public Class Form1

    '---- Initialize application ----------------------------------------------
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        xtrace_init()
        xtrace_subs("Form1_Load")
        xtrace_header()
        WriteInfo("Log file = " & LogFile)
        xtrace("Initializing")
        ReadDefaults()
        Read_Command_Line_Arg()
        AppName2 = AppName.Replace("-", " ")
        Me.Text = AppName2 & " V" & AppVer
        If Verbose Then
            xtrace_i("Enable the settings menu")
            SettingsToolStripMenuItem.Enabled = True
        End If

        ' Button text
        BtnStartInstall.Text = "Install: " & PSName
        BtnStartRun.Text = "Run: " & PSName

        If StartSeq = 2 Then
            xtrace_i("StartSeq = 2, installation was started")
            BtnStartInstall.Visible = False
            BtnStartRun.Visible = False
            Timer1.Start()
        End If

        xtrace_sube("Form1_Load")
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        xtrace_subs("Timer1_Tick")

        ShowActivity("Applying...")
        Timer1.Stop()

        StartInstall()

        xtrace_sube("Timer1_Tick")
    End Sub

    Sub ShowActivity(Message As String)
        Panel1.Visible = True
        Panel1.Top = Panel1.Top - 50
        Label1.Text = Message
    End Sub

    '==== Main Menu ===========================================================

    '---- File, Exit
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        xtrace_subs("Menu, File, Exit")
        exit_program()
    End Sub

    '---- Show Settings -------------------------------------------------------
    Private Sub ShowSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowSettingsToolStripMenuItem.Click
        xtrace_subs("Menu, Settings, Show settings")
        If ShowSettingsToolStripMenuItem.Checked Then
            TextBoxInfo.Visible = True
            TextBoxInfo.Dock = DockStyle.Fill
        Else
            TextBoxInfo.Visible = False
        End If
        xtrace_sube("Menu, Settings, Show settings")
    End Sub

    '---- Show Log
    Private Sub ShowLogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowLogToolStripMenuItem.Click
        xtrace_subs("Menu, Settings, Show log file")
        Process.Start(LogFile)
    End Sub

    '---- Show help
    Private Sub HelpToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem1.Click
        xtrace_subs("Menu, Help, Help")
        ShowHelp()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        xtrace_subs("Menu, Help, About")
        ShowHelpAbout()
    End Sub

    '---- Start install ----
    Private Sub BtnStartInstall_Click(sender As Object, e As EventArgs) Handles BtnStartInstall.Click
        xtrace_subs("BtnStartInstall_Click")

        ElevateMe()
        exit_program()

        xtrace_sube("BtnStartInstall_Click")
    End Sub

    Public Shared Sub StartInstall()
        xtrace_subs("StartInstall")

        CreatePSScript("Inst")
        CreateStartIcon()
        wait(1)

        Form1.Label1.Text = "Finished."
        wait(2)
        exit_program()

        xtrace_sube("StartInstall")
    End Sub

    '---- Start Run ----
    Private Sub BtnStartRun_Click(sender As Object, e As EventArgs) Handles BtnStartRun.Click
        xtrace_subs("BtnStartRun_Click")

        BtnStartInstall.Visible = False
        BtnStartRun.Visible = False
        ShowActivity("Starting")

        CreatePSScript("Run")
        RunPS()

        Label1.Text = "Finished"
        wait(2)
        exit_program()

        xtrace_sube("BtnStartRun_Click")
    End Sub


End Class
