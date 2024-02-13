
Imports System.Data.SqlClient
Imports System.IO
Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Module1.OpenConnection()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Module1.CloseConnection()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Module1.transferData()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim connectionString As String = "Server=TMS-DEV-2\MSSQLSERVER_2014; Database=GENBASE;User Id=sa;Password=medical;"
        Dim queryString As String = "SELECT PF.[PATCODE], PF.[PATLAST], PF.[PATFIRST], PO.[POBMP1], PO.[POBMP2], PO.[POBMP3], PO.[POBMP4], PO.[IMG1], PO.[IMG2], PO.[IMG3], PO.[IMG4], PO.[IMG5], PO.[IMG6] FROM [GENBASE].[dbo].[PAT1FILE] PF LEFT JOIN [GENBASE].[dbo].[POTHERS] PO ON PF.[PATCODE] = PO.[PATCODE]"
        Dim examList As New List(Of ExamDetails)

        Using connection As New SqlConnection(connectionString)
            Dim command As New SqlCommand(queryString, connection)
            Try
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                While reader.Read()
                    Dim exam As New ExamDetails()
                    exam.PatCode = reader("PATCODE").ToString()
                    exam.PatLast = reader("PATLAST").ToString()
                    exam.PatFirst = reader("PATFIRST").ToString()
                    exam.POBMP1 = If(Not IsDBNull(reader("POBMP1")), reader("POBMP1").ToString(), String.Empty)
                    exam.POBMP2 = If(Not IsDBNull(reader("POBMP2")), reader("POBMP2").ToString(), String.Empty)
                    exam.POBMP3 = If(Not IsDBNull(reader("POBMP3")), reader("POBMP3").ToString(), String.Empty)
                    exam.POBMP4 = If(Not IsDBNull(reader("POBMP4")), reader("POBMP4").ToString(), String.Empty)
                    exam.IMG1 = If(Not IsDBNull(reader("IMG1")), reader("IMG1").ToString(), String.Empty)
                    exam.IMG2 = If(Not IsDBNull(reader("IMG2")), reader("IMG2").ToString(), String.Empty)
                    exam.IMG3 = If(Not IsDBNull(reader("IMG3")), reader("IMG3").ToString(), String.Empty)
                    exam.IMG4 = If(Not IsDBNull(reader("IMG4")), reader("IMG4").ToString(), String.Empty)
                    exam.IMG5 = If(Not IsDBNull(reader("IMG5")), reader("IMG5").ToString(), String.Empty)
                    exam.IMG6 = If(Not IsDBNull(reader("IMG6")), reader("IMG6").ToString(), String.Empty)
                    examList.Add(exam)
                End While

                reader.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Using

        DataGridView1.DataSource = examList
    End Sub

    Public Sub btnUnzip_Click(sender As Object, e As EventArgs) Handles btnUnzip.Click
        Dim zipFilePath = "c:\example\52940.ARJ"
        Dim destinationPath = "c:\example\extract_arj"
        UnzipWith7Zip(zipFilePath, destinationPath)
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim connectionString As String = "Server=TMS-DEV-2\MSSQLSERVER_2014; Database=GENBASE;User Id=sa;Password=medical;"
        Dim queryString As String = "SELECT PF.[PATCODE], PF.[PATLAST], PF.[PATFIRST], PO.[POBMP1], PO.[POBMP2], PO.[POBMP3], PO.[POBMP4], PO.[IMG1], PO.[IMG2], PO.[IMG3], PO.[IMG4], PO.[IMG5], PO.[IMG6] FROM [GENBASE].[dbo].[PAT1FILE] PF LEFT JOIN [GENBASE].[dbo].[POTHERS] PO ON PF.[PATCODE] = PO.[PATCODE]"
        Dim examList As New List(Of ExamDetails)

        Using connection As New SqlConnection(connectionString)
            Dim command As New SqlCommand(queryString, connection)
            Try
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()

                For i As Integer = 2 To 5
                    If reader.Read() Then
                        Dim nextExam As New ExamDetails()
                        nextExam.PatCode = reader("PATCODE").ToString()
                        nextExam.POBMP1 = If(Not IsDBNull(reader("POBMP1")), reader("POBMP1").ToString(), String.Empty)
                        nextExam.POBMP2 = If(Not IsDBNull(reader("POBMP2")), reader("POBMP2").ToString(), String.Empty)
                        nextExam.POBMP3 = If(Not IsDBNull(reader("POBMP3")), reader("POBMP3").ToString(), String.Empty)
                        nextExam.POBMP4 = If(Not IsDBNull(reader("POBMP4")), reader("POBMP4").ToString(), String.Empty)

                        transferTheData(nextExam.POBMP1, nextExam.PatCode)
                        transferTheData(nextExam.POBMP2, nextExam.PatCode)
                        transferTheData(nextExam.POBMP3, nextExam.PatCode)
                        transferTheData(nextExam.POBMP4, nextExam.PatCode)
                    Else
                        Exit For ' Εάν δεν υπάρχουν άλλες εγγραφές, βγαίνουμε από τον βρόχο
                    End If
                Next

                reader.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Using

    End Sub

End Class
