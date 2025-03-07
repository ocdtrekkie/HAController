Module modLibrary
    ''' <summary>
    ''' Retrieves an ebook from the repository and emails it to the user
    ''' </summary>
    ''' <param name="strBarcode">Material barcode</param>
    ''' <returns>Temp: Type of material</returns>
    Function CheckOutEbook(ByVal strBarcode As String, ByVal strRequestor As String) As String
        If My.Settings.Library_Enable = True Then
            If IsNumeric(strBarcode) = True Then
                'File Selection Section
                Dim SelectedFile As System.IO.FileInfo
                Dim SelectedFolder As New System.IO.DirectoryInfo(My.Settings.Library_Repository & strBarcode)
                If SelectedFolder.Exists Then
                    Dim SelectedFilesInFolder As String()
                    SelectedFilesInFolder = IO.Directory.GetFiles(SelectedFolder.FullName, "*.pdf", IO.SearchOption.TopDirectoryOnly)
                    If SelectedFilesInFolder.Count = 1 Then
                        SelectedFile = New System.IO.FileInfo(SelectedFilesInFolder(0))
                    ElseIf SelectedFilesInFolder.Count > 1 Then
                        My.Application.Log.WriteEntry("Multiple PDFs found in folder, remote checkout is not supported")
                        SelectedFile = Nothing
                    Else
                        SelectedFilesInFolder = IO.Directory.GetFiles(SelectedFolder.FullName, "*.epub", IO.SearchOption.TopDirectoryOnly)
                        If SelectedFilesInFolder.Count = 1 Then
                            SelectedFile = New System.IO.FileInfo(SelectedFilesInFolder(0))
                        ElseIf SelectedFilesInFolder.Count > 1 Then
                            My.Application.Log.WriteEntry("Multiple EPUBs found in folder, remote checkout is not supported")
                            SelectedFile = Nothing
                        Else
                            My.Application.Log.WriteEntry("Folder contains neither PDFs nor EPUBs, remote checkout is not supported")
                            SelectedFile = Nothing
                        End If
                    End If
                Else
                    SelectedFile = New System.IO.FileInfo(My.Settings.Library_Repository & strBarcode & ".pdf")
                    If SelectedFile.Exists = False Then
                        SelectedFile = New System.IO.FileInfo(My.Settings.Library_Repository & strBarcode & ".epub")
                        If SelectedFile.Exists = False Then
                            My.Application.Log.WriteEntry("No folders, PDFs, or EPUBs were found, remote checkout is not supported")
                            SelectedFile = Nothing
                        End If
                    End If
                End If
                'File Transmission Section
                If SelectedFile IsNot Nothing Then
                    My.Application.Log.WriteEntry("Requested file for barcode " & strBarcode & " found at " & SelectedFile.FullName & ". Size is " & SelectedFile.Length & " bytes.")
                    If SelectedFile.Length > (5 * 1024 * 1024) Then
                        Return "Requested file is too large to send" 'TODO: Support compressing larger files
                    Else
                        modMail.Send("Your library materials", "Thank you for checking out barcode " & strBarcode, SelectedFile.FullName, strRequestor)
                        Return "Requested file sent"
                    End If
                Else
                    Return "Acceptable file to return was not found"
                End If
            Else
                Return "Invalid barcode format"
            End If
        Else
            My.Application.Log.WriteEntry("Library module is disabled")
            Return "Disabled"
        End If
    End Function

    ''' <summary>
    ''' Strips the hyphens from an ISBN number, determines if it is ISBN10 or ISBN13, and then returns the converted value.
    ''' </summary>
    ''' <param name="strInputISBN">ISBN to convert</param>
    ''' <returns>ISBN converted to other format</returns>
    Function ConvertISBN(ByVal strInputISBN As String) As String
        strInputISBN = Replace(strInputISBN, "-", "")
        If strInputISBN.Length = 10 Then
            Return GetISBN13(strInputISBN)
        ElseIf strInputISBN.Length = 13 Then
            Return GetISBN10(strInputISBN)
        Else
            My.Application.Log.WriteEntry("Malformed ISBN entered", TraceEventType.Warning)
            Return "BAD ISBN"
        End If
    End Function

    Function Disable() As String
        Unload()
        My.Settings.Library_Enable = False
        My.Application.Log.WriteEntry("Library module is disabled")
        Return "Library module disabled"
    End Function

    Function Enable() As String
        My.Settings.Library_Enable = True
        My.Application.Log.WriteEntry("Library module is enabled")
        Load()
        Return "Library module enabled"
    End Function

    ''' <summary>
    ''' Converts a 13-digit ISBN to a 10-digit ISBN.
    ''' </summary>
    ''' <param name="strISBN13">ISBN13 to convert</param>
    ''' <returns>ISBN10</returns>
    Function GetISBN10(ByVal strISBN13 As String) As String
        ' Credit to jmaddalone at https://snipplr.com/view/5127/convert-isbn-13-to-10/
        If strISBN13.Substring(0, 3) <> "978" Then
            Return "NO ISBN10"
        Else
            Dim a As Integer
            Dim b As Integer
            Dim c As Integer
            Dim d As Integer
            Dim e As Integer
            Dim f As Integer
            Dim g As Integer
            Dim h As Integer
            Dim i As Integer
            Dim j As Integer
            Dim k As Integer
            Dim l As Integer
            Dim m As Integer
            Dim n As Integer
            Dim o As Integer
            Dim o2 As String = ""
            Dim n2 As Integer
            Dim isbnarr(12) As Integer
            For i = 0 To 12
                isbnarr(i) = CInt(Mid(strISBN13, i + 1, 1))
            Next
            a = isbnarr(0)
            b = isbnarr(1)
            c = isbnarr(2)
            d = isbnarr(3)
            e = isbnarr(4)
            f = isbnarr(5)
            g = isbnarr(6)
            h = isbnarr(7)
            i = isbnarr(8)
            j = isbnarr(9)
            k = isbnarr(10)
            l = isbnarr(11)
            m = isbnarr(12)
            n = (d * 10) + (9 * e) + (8 * f) + (7 * g) + (6 * h) + (5 * i) + (4 * j) + (3 * k) + (2 * l)
            n2 = Int((n / 11) + 1)
            o = (11 * n2) - n
            If o = 10 Then
                o2 = "X"
            ElseIf o = 11 Then
                o2 = "0"
            End If
            Return CStr(d & e & f & g & h & i & j & k & l & o2)
        End If
    End Function

    ''' <summary>
    ''' Converts a 10-digit ISBN to a 13-digit ISBN.
    ''' </summary>
    ''' <param name="strISBN10">ISBN10 to convert</param>
    ''' <returns>ISBN13</returns>
    Function GetISBN13(ByVal strISBN10 As String) As String
        ' Credit to Mehrdad, Asghari at https://www.codeproject.com/tips/75999/%2fTips%2f75999%2fConvert-ISBN10-To-ISBN-13
        Dim isbn10 As String = "978" & strISBN10.Substring(0, 9)
        Dim isbn10_1 As Integer = Convert.ToInt32(isbn10.Substring(0, 1))
        Dim isbn10_2 As Integer = Convert.ToInt32(Convert.ToInt32(isbn10.Substring(1, 1)) * 3)
        Dim isbn10_3 As Integer = Convert.ToInt32(isbn10.Substring(2, 1))
        Dim isbn10_4 As Integer = Convert.ToInt32(Convert.ToInt32(isbn10.Substring(3, 1)) * 3)
        Dim isbn10_5 As Integer = Convert.ToInt32(isbn10.Substring(4, 1))
        Dim isbn10_6 As Integer = Convert.ToInt32(Convert.ToInt32(isbn10.Substring(5, 1)) * 3)
        Dim isbn10_7 As Integer = Convert.ToInt32(isbn10.Substring(6, 1))
        Dim isbn10_8 As Integer = Convert.ToInt32(Convert.ToInt32(isbn10.Substring(7, 1)) * 3)
        Dim isbn10_9 As Integer = Convert.ToInt32(isbn10.Substring(8, 1))
        Dim isbn10_10 As Integer = Convert.ToInt32(Convert.ToInt32(isbn10.Substring(9, 1)) * 3)
        Dim isbn10_11 As Integer = Convert.ToInt32(isbn10.Substring(10, 1))
        Dim isbn10_12 As Integer = Convert.ToInt32(Convert.ToInt32(isbn10.Substring(11, 1)) * 3)
        Dim k As Integer = (isbn10_1 + isbn10_2 + isbn10_3 + isbn10_4 + isbn10_5 + isbn10_6 + isbn10_7 + isbn10_8 + isbn10_9 + isbn10_10 + isbn10_11 + isbn10_12)
        Dim checkdigit As Integer = 10 - ((isbn10_1 + isbn10_2 + isbn10_3 + isbn10_4 + isbn10_5 + isbn10_6 + isbn10_7 + isbn10_8 + isbn10_9 + isbn10_10 + isbn10_11 + isbn10_12) Mod 10)
        If checkdigit = 10 Then checkdigit = 0
        Return isbn10 & checkdigit.ToString()
    End Function

    Function Load() As String
        If My.Settings.Library_Enable = True Then
            My.Application.Log.WriteEntry("Loading library module")
            If My.Settings.Library_Repository = "" Then
                My.Application.Log.WriteEntry("No library location set, asking for it")
                My.Settings.Library_Repository = InputBox("Enter the folder location of your library data.", "Library Repository")
            End If
            Return "Library module loaded"
        Else
            My.Application.Log.WriteEntry("Library module is disabled, module not loaded")
            Return "Library module is disabled, module not loaded"
        End If
    End Function

    Function Unload() As String
        My.Application.Log.WriteEntry("Unloading library module")
        Return "Library module unloaded"
    End Function
End Module
