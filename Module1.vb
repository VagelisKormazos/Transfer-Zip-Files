Imports System.IO
Imports System.IO.Compression


Module Module1

    Sub test()
        Dim startPath As String = "c:\example\start"
        Dim zipPath As String = "c:\example\result.zip"
        Dim extractPath As String = "c:\example\extract"

        ZipFile.CreateFromDirectory(startPath, zipPath)

        ZipFile.ExtractToDirectory(zipPath, extractPath)
    End Sub

    Sub testARJ()
        Dim zipPath As String = "c:\example\52940.ARJ"
        Dim extractPath As String = "c:\example\extract_arj"
        ZipFile.ExtractToDirectory(zipPath, extractPath)
    End Sub

End Module