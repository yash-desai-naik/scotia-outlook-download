Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Timers
Imports System.Text.RegularExpressions

Public Class Form1
    Private notifyIcon As System.Windows.Forms.NotifyIcon

    Private WithEvents timer As New Timer()
    Private outlookApp As Outlook.Application
    Private downloadPath As String
    Private interval As Integer

    Public Sub New()
        InitializeComponent()

        ' Initialize timer
        timer.Enabled = True
        AddHandler timer.Elapsed, AddressOf Timer_Elapsed

        ' Initialize Outlook application
        outlookApp = New Outlook.Application()

        ' Retrieve download path and interval from settings
        downloadPath = My.Settings.DownloadPath
        interval = My.Settings.Interval

        ' Update UI with settings
        ToolStripStatusLabel1.Text = downloadPath
        txtInterval.Text = interval.ToString()

        ' Set timer interval
        timer.Interval = TimeSpan.FromSeconds(interval).TotalMilliseconds
    End Sub

    Private Sub Timer_Elapsed(sender As Object, e As ElapsedEventArgs)
        ' Check for new emails
        Dim inboxFolder As Outlook.Folder = outlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
        For Each item As Object In inboxFolder.Items
            If TypeOf item Is Outlook.MailItem Then
                Dim mailItem As Outlook.MailItem = DirectCast(item, Outlook.MailItem)
                ' Process only unread emails
                If Not mailItem.UnRead Then
                    Continue For
                End If

                ' Define regular expression pattern to match subject
                Dim pattern As String = "(RE: |RV: )?(?<Prefix>.+?) - Dodd-Frank DeMinimis Extract Request \| (?<StartDate>.+?) - (?<EndDate>.+?)( - SwapOne)?"

                ' Check if the subject matches the pattern
                Dim match As Match = Regex.Match(mailItem.Subject, pattern, RegexOptions.IgnoreCase)

                If match.Success Then
                    ' Extract information from the subject
                    Dim prefix As String = match.Groups("Prefix").Value.Trim()
                    Dim startDate As String = match.Groups("StartDate").Value.Trim()
                    Dim endDate As String = match.Groups("EndDate").Value.Trim()

                    ' Display a message box with subject details
                    Dim message As String = $"Email received with subject:{Environment.NewLine}{Environment.NewLine}" &
                                            $"Prefix: {prefix}{Environment.NewLine}" &
                                            $"Start Date: {startDate}{Environment.NewLine}" &
                                            $"End Date: {endDate}"
                    MessageBox.Show(message, "New Email", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' Download attachments
                    For Each attachment As Outlook.Attachment In mailItem.Attachments
                        ' Save the attachment to the selected location
                        attachment.SaveAsFile(System.IO.Path.Combine(downloadPath, attachment.FileName))
                    Next

                    ' Optionally, you can mark the email as read once processed
                    mailItem.UnRead = False
                    ' Save changes
                    mailItem.Save()

                End If


            End If
        Next
    End Sub


    Private Sub btnSelectDownloadPath_Click(sender As Object, e As EventArgs) Handles btnSelectDownloadPath.Click
        ' Open folder browser dialog to select download location
        Dim folderBrowserDialog As New FolderBrowserDialog()
        If folderBrowserDialog.ShowDialog() = DialogResult.OK Then
            downloadPath = folderBrowserDialog.SelectedPath
            ToolStripStatusLabel1.Text = downloadPath

            ' Save download path to settings
            My.Settings.DownloadPath = downloadPath
            My.Settings.Save()
        End If
    End Sub

    Private Sub btnUpdateInterval_Click(sender As Object, e As EventArgs) Handles btnUpdateInterval.Click
        ' Update interval based on the value entered in the interval textbox
        Dim newInterval As Integer
        If Integer.TryParse(txtInterval.Text, newInterval) AndAlso newInterval > 0 AndAlso newInterval <= Integer.MaxValue Then
            ' Set timer interval
            timer.Interval = TimeSpan.FromSeconds(newInterval).TotalMilliseconds

            ' Save interval to settings
            interval = newInterval
            My.Settings.Interval = interval
            My.Settings.Save()
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize NotifyIcon
        notifyIcon = New System.Windows.Forms.NotifyIcon()
        notifyIcon.Icon = Me.Icon
        notifyIcon.Text = Me.Text

        ' Add a context menu to the NotifyIcon (optional)
        Dim contextMenu As New System.Windows.Forms.ContextMenu()
        contextMenu.MenuItems.Add("Restore", AddressOf RestoreForm)
        contextMenu.MenuItems.Add("Exit", AddressOf ExitApplication)
        notifyIcon.ContextMenu = contextMenu
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
            notifyIcon.Visible = True
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        ' Ensure that the application exits properly by disposing of the NotifyIcon
        notifyIcon.Dispose()
    End Sub

    Private Sub RestoreForm(sender As Object, e As EventArgs)
        ' Show the form and hide the NotifyIcon when "Restore" is clicked
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        notifyIcon.Visible = False
    End Sub

    Private Sub ExitApplication(sender As Object, e As EventArgs)
        ' Exit the application when "Exit" is clicked
        Me.Close()
    End Sub
End Class
