Imports System.IO
Imports System.IO.Compression
Imports System.Data.SqlClient

Module Module1
    Dim connectionString As String = "Server=TMS-DEV-2\MSSQLSERVER_2014; Database=GENBASE;User Id=sa;Password=medical;"
    Dim InitialFolderPath As String = "c:\example"
    Public Class ExamDetails
        Public Property PatCode As String
        Public Property PatLast As String
        Public Property PatFirst As String
        Public Property POBMP1 As String

        Public Property POBMP2 As String
        Public Property POBMP3 As String
        Public Property POBMP4 As String
        Public Property IMG1 As String
        Public Property IMG2 As String
        Public Property IMG3 As String
        Public Property IMG4 As String
        Public Property IMG5 As String
        Public Property IMG6 As String
    End Class

    Sub OpenConnection()
        Using conn As New SqlConnection(connectionString)
            Try
                conn.Open()
                MessageBox.Show(conn.State.ToString())
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Using
    End Sub

    Public Sub CloseConnection()
        Using conn As New SqlConnection(connectionString)
            Try
                conn.Close()
                MessageBox.Show("Connection Closed. Connection State: " & conn.State.ToString())
            Catch ex As Exception
                MessageBox.Show("Error closing connection: " & ex.Message)
            End Try
        End Using
    End Sub



    Public Sub UnzipWith7Zip(zipFilePath As String, destinationPath As String)
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

    Public Sub transferData()

        Dim FolderPath As String = "c:\example\"
        Dim zipFilePath As String = "C:\TMS_FILES\CAPTURE"
        Dim files() As FileInfo
        Dim directoryInfo As New DirectoryInfo(InitialFolderPath)

        files = directoryInfo.GetFiles()

        For Each fileinf As FileInfo In files

            Dim fileName As String = fileinf.Name
            Dim finalDestinationPath As String = Path.Combine(zipFilePath, fileinf.Name.Substring(0, fileName.IndexOf(".")), "Images", fileinf.LastWriteTime.ToString("yyyy-MM-dd"))
            Dim finalLoadPath As String = Path.Combine(FolderPath, fileName)
            MsgBox(finalLoadPath)
            MsgBox(finalDestinationPath)
            UnzipWith7Zip(finalLoadPath, finalDestinationPath)

        Next

    End Sub


    Public Sub transferTheData(data As String, patientId As String)
        Dim FolderPath As String = "c:\example\"
        Dim zipFilePath As String = "C:\TMS_FILES\CAPTURE"
        Dim files() As FileInfo
        Dim directoryInfo As New DirectoryInfo(InitialFolderPath)

        files = directoryInfo.GetFiles()

        Dim fileFound As Boolean = False

        For Each fileinf As FileInfo In files
            If (fileinf.Name = data + ".ARJ") Then
                Dim fileName As String = fileinf.Name
                Dim finalDestinationPath As String = Path.Combine(zipFilePath, patientId, "Images", fileinf.LastWriteTime.ToString("yyyy-MM-dd"))
                Dim finalLoadPath As String = Path.Combine(FolderPath, fileName)
                MsgBox(finalLoadPath)
                MsgBox(finalDestinationPath)
                UnzipWith7Zip(finalLoadPath, finalDestinationPath)
                fileFound = True
                Exit For
            End If
        Next

        If Not fileFound Then
            MsgBox("File not found: " & data & ".ARJ")
        End If

    End Sub

    Public Sub test()
        Dim startPath As String = "c:\example\start"
        Dim zipPath As String = "c:\example\result.zip"
        Dim extractPath As String = "c:\example\extract"

    End Sub

    Public Sub testARJ()
        Dim zipPath As String = "c:\example\52940.ARJ"
        Dim extractPath As String = "c:\example\extract_arj"

    End Sub

End Module