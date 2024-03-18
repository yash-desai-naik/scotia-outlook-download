Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Timers
Imports System.Text.RegularExpressions
Imports System.IO

Public Class Form1
    Private notifyIcon As System.Windows.Forms.NotifyIcon

    Private WithEvents timer As New Timer()
    Private outlookApp As Outlook.Application
    Private downloadPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

    Private interval As Integer

    Public Sub New()
        InitializeComponent()


        ' Initialize timer
        timer.Enabled = True
        AddHandler timer.Elapsed, AddressOf Timer_Elapsed

        ' Initialize Outlook application
        outlookApp = New Outlook.Application()

        ' Retrieve  interval from settings
        interval = My.Settings.Interval

        downloadPath = Path.Combine(downloadPath, "scotia-automation")
        MsgBox("attempt to create: " & downloadPath)
        EnsureCreation(downloadPath)

        Try
            MsgBox("CreateYearMonthFolders")
            ' Create folders for YYYY\MMM in download path
            CreateYearMonthFolders()
        Catch ex As Exception
            MsgBox("Sorr, " & ex.Message)
        End Try


        ' Update UI with settings
        ToolStripStatusLabel1.Text = downloadPath
        txtInterval.Text = interval.ToString()

        ' Set timer interval
        timer.Interval = TimeSpan.FromSeconds(interval).TotalMilliseconds
    End Sub

    Private Sub CreateYearMonthFolders()
        MsgBox("In CreateYearMonthFolders")
        Dim currentDate As Date = Date.Now
        Dim yearFolder As String = Path.Combine(downloadPath, currentDate.ToString("yyyy"))
        MsgBox("attempt to create: " & yearFolder)
        EnsureCreation(yearFolder)
        Dim prevMonthFolder As String = Path.Combine(yearFolder, currentDate.AddMonths(-1).ToString("MMM"))
        MsgBox("attempt to create: " & prevMonthFolder)
        EnsureCreation(prevMonthFolder)

        '' Create year folder if it doesn't exist
        'If Not Directory.Exists(yearFolder) Then
        '    Directory.CreateDirectory(yearFolder)
        'End If

        '' Create month folder if it doesn't exist
        'If Not Directory.Exists(prevMonthFolder) Then
        '    Directory.CreateDirectory(prevMonthFolder)
        'End If

        ' Initialize additional folders under the month folder
        Dim additionalFolders As String() = {
            "Latam De Minimis Calculation",
            "OPICS Scotia Investments Jamaica Limited",
            "Supporting Files K2 and Murex",
            "SCOTS",
            "Calculations"
        }

        For Each folderName As String In additionalFolders
            Dim folderPath As String = Path.Combine(prevMonthFolder, folderName)
            'If Not Directory.Exists(folderPath) Then
            '    Directory.CreateDirectory(folderPath)
            'End If
            MsgBox("attempt to create: " & folderPath)
            EnsureCreation(folderPath)
        Next
    End Sub

    Private Sub SaveAttachment(attachment As Outlook.Attachment, targetFolder As String)
        ' Save the attachment to the target folder
        'If Not Directory.Exists(targetFolder) Then
        '    Directory.CreateDirectory(targetFolder)
        'End If

        Dim filePath As String = Path.Combine(targetFolder, attachment.FileName)
        'EnsureCreation(filePath, method:="file")
        attachment.SaveAsFile(filePath)
    End Sub

    Private Sub Timer_Elapsed(sender As Object, e As ElapsedEventArgs)
        Dim inboxFolder As Outlook.Folder
        ' Check if outlookApp is initialized
        If outlookApp Is Nothing Then
            Try
                ' Attempt to create a new instance of Outlook
                outlookApp = CreateObject("Outlook.Application")
            Catch ex As Exception
                ' Handle any errors that occur during the creation of Outlook
                MessageBox.Show("Unable to open Outlook: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try

            ' Check if Outlook was successfully opened
            If outlookApp Is Nothing Then
                ' Outlook could not be opened
                MessageBox.Show("Unable to open Outlook.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If
        End If

        ' Attempt to access Outlook objects with a retry mechanism
        Dim retryCount As Integer = 0
        Do
            Try
                ' Attempt to access the inbox folder
                inboxFolder = outlookApp.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
                ' Proceed with checking for new emails
                Exit Do ' Exit the loop if successful
            Catch ex As System.Runtime.InteropServices.COMException When ex.ErrorCode = &H8001010A AndAlso retryCount < 10
                ' The application is busy, so wait for a short time and retry
                System.Threading.Thread.Sleep(1000) ' Wait for 1 second before retrying
                retryCount += 1 ' Increment the retry count
            Catch ex As Exception
                ' Handle any other exceptions
                MessageBox.Show("Error accessing Outlook: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End Try
        Loop While retryCount < 10 ' Retry a maximum of 10 times

        ' If the loop exits without success after 10 retries, display an error message
        If retryCount >= 10 Then
            MessageBox.Show("Unable to access Outlook. Please try again later.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

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
                        ' Define base target folder
                        Dim currentDate As Date = Date.Now
                        Dim yearFolder As String = Path.Combine(downloadPath, currentDate.ToString("yyyy"))
                        Dim prevMonthFolder As String = Path.Combine(yearFolder, currentDate.AddMonths(-1).ToString("MMM"))
                        Dim targetFolder As String = prevMonthFolder ' Base target folder
                        EnsureCreation(targetFolder)

                        ' Append subfolders based on conditions
                        Dim fileName As String = attachment.FileName.ToUpper()
                        If fileName.Contains("US PERSON") OrElse fileName.Contains("US_PERSON") Then
                            targetFolder = Path.Combine(targetFolder, "Latam De Minimis Calculation", "CFTC Deminimis LatAm Extracts", "US Person List")
                            EnsureCreation(targetFolder)
                        ElseIf fileName.StartsWith("CARTERA") OrElse
                           fileName.StartsWith("DEMINIMISREPORT") OrElse
                           fileName.StartsWith("DERIVATIVES") OrElse
                           fileName.StartsWith("DODD-FRANK") OrElse
                           fileName.StartsWith("MINIMIS CALCULATION TEMPLATE") Then
                            targetFolder = Path.Combine(targetFolder, "Latam De Minimis Calculation", "CFTC Deminimis LatAm Extracts")
                            EnsureCreation(targetFolder)
                        ElseIf fileName.StartsWith("FX") OrElse
                           fileName.StartsWith("FOREX") Then
                            targetFolder = Path.Combine(targetFolder, "OPICS Scotia Investments Jamaica Limited")
                            EnsureCreation(targetFolder)
                        ElseIf fileName.Contains("DF_DEMINIMIS_EXTRACT") Then
                            targetFolder = Path.Combine(targetFolder, "Supporting Files K2 and Murex", "Murex")
                            EnsureCreation(targetFolder)
                        End If

                        ' Save attachment to the target folder
                        SaveAttachment(attachment, targetFolder)
                    Next

                    ' Optionally, you can mark the email as read once processed
                    mailItem.UnRead = False
                    ' Save changes
                    mailItem.Save()

                End If

            End If
        Next
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

    Public Function EnsureCreation(path As String, Optional ByVal method As String = "dir") As Boolean


        If method = "dir" Then
            Try
                If Not Directory.Exists(path) Then
                    Dim directoryInfo As New DirectoryInfo(path)
                    directoryInfo.Create() ' Create directory with intermediate directories if needed
                    MsgBox("Folder Createed: " & path)
                    Return True
                End If
            Catch ex As Exception
                Console.WriteLine($"Error creating directory: {path} ({ex.Message})")
                Return False
            End Try
        ElseIf method = "file" Then
            Try
                Dim fileStream As New FileStream(path, FileMode.Create)
                fileStream.Close() ' Create an empty file
                Return True
            Catch ex As Exception
                Console.WriteLine($"Error creating file: {path} ({ex.Message})")
                Return False
            End Try
        Else
            Console.WriteLine($"Invalid method: {method}. Supported methods are 'dir' and 'file'.")
            Return False
        End If
    End Function
End Class
