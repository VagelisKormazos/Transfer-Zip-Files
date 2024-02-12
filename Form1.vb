
Imports System.IO

Public Class Form1
    Dim InitialFolderPath As String = "c:\example"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Module1.test()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Module1.testARJ()
    End Sub

    Private Sub btnUnzip_Click(sender As Object, e As EventArgs) Handles btnUnzip.Click
        Dim zipFilePath = "c:\example\52940.ARJ"
        Dim destinationPath = "c:\example\extract_arj"
        UnzipWith7Zip(zipFilePath, destinationPath)
    End Sub

    Private Sub UnzipWith7Zip(zipFilePath As String, destinationPath As String)
        Try

            Dim processInfo As New ProcessStartInfo()
            processInfo.FileName = "7z.exe" ' assuming 7-Zip is in the system PATH
            processInfo.Arguments = $"x ""{zipFilePath}"" -o""{destinationPath}"" -y" ' -y is used to auto-respond yes to any prompts
            processInfo.WindowStyle = ProcessWindowStyle.Hidden ' Hide the command prompt window

            Dim process As Process = Process.Start(processInfo)
            process.WaitForExit()

            If process.ExitCode = 0 Then
                MessageBox.Show("Extraction successful!")
            Else
                MessageBox.Show("Extraction failed.")
            End If
        Catch ex As Exception
            MessageBox.Show($"An error occurred: {ex.Message}")
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim FolderPath As String = "c:\example\"
        Dim files() As FileInfo
        Dim directoryInfo As New DirectoryInfo(InitialFolderPath)

        files = directoryInfo.GetFiles()

        Dim zipFilePath As String = "C:\TMS_FILES\CAPTURE"   ' Ορίζουμε την τελική διαδρομή ως υποφάκελο του φακέλου εφαρμογής

        For Each fileinf As FileInfo In files


            Dim fileName As String = fileinf.Name
            Dim finalDestinationPath As String = Path.Combine(zipFilePath, fileinf.Name.Substring(0, fileName.IndexOf(".")), "Images", fileinf.LastWriteTime.ToString("yyyy-MM-dd"))
            Dim finalLoadPath As String = Path.Combine(FolderPath, fileName)
            MsgBox(finalLoadPath)
            MsgBox(finalDestinationPath)
            UnzipWith7Zip(finalLoadPath, finalDestinationPath)

        Next

    End Sub
End Class
