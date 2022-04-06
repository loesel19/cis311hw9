Imports Microsoft.Office.Interop
Public Class Form1
    Dim myStudents As List(Of clsStudent)

    Public Sub populateList()
        myStudents.Add(New clsStudent("V.A.", "Borstellis", {25, 25, 25, 25}, 100.0))
        myStudents.Add(New clsStudent("A.S.", "Reid", {20, 21, 20, 18}, 75.0))
        myStudents.Add(New clsStudent("C.U.", "Tyler", {19, 20, 21, 24}, 75.5))
        myStudents.Add(New clsStudent("H.A.", "Renee", {20, 23, 23, 25}, 80.5))
        myStudents.Add(New clsStudent("I.A.", "Douglas", {24, 23, 25, 25}, 95.0))
        myStudents.Add(New clsStudent("M.A.", "Elenaips", {23, 24, 23, 21}, 94.5))
        myStudents.Add(New clsStudent("A.L.", "Emmet", {21, 19, 18, 15}, 73.0))
        myStudents.Add(New clsStudent("S.U.", "James", {21, 24, 23, 22}, 87.5))
        myStudents.Add(New clsStudent("S.H.", "Issacs", {23, 24, 21, 21}, 93.0))
        myStudents.Add(New clsStudent("B.I.", "Opus", {23, 24, 25, 23}, 97.5))
        myStudents.Add(New clsStudent("T.R.", "Alski", {24, 25, 25, 23}, 95.5))
        myStudents.Add(New clsStudent("H.E.", "Zeus", {23, 24, 23, 23}, 77.0))
        myStudents.Add(New clsStudent("S.C.", "Ustaf", {24, 23, 24, 25}, 91.0))
        myStudents.Add(New clsStudent("K.I.", "Chrint", {23, 23, 24, 21}, 89.0))
        myStudents.Add(New clsStudent("J.E.", "Yaz", {25, 24, 23, 24}, 92.5))
        myStudents.Add(New clsStudent("F.R.", "Franks", {23, 19, 18, 23}, 88.5))
        myStudents.Add(New clsStudent("W.I.", "Walton", {24, 23, 23, 19}, 90.0))
        myStudents.Add(New clsStudent("K.A.", "Gilch", {24, 23, 25, 24}, 92.0))
        myStudents.Add(New clsStudent("R.O.", "Little", {23, 24, 23, 24}, 94.0))
        myStudents.Add(New clsStudent("S.A.", "Xerxes", {24, 23, 25, 23}, 94.0))
        myStudents.Add(New clsStudent("W.I.", "Harris", {23, 24, 25, 23}, 92.0))
        myStudents.Add(New clsStudent("T.I.", "Vargo", {24, 23, 25, 25}, 99.0))
        myStudents.Add(New clsStudent("I.E.", "Interas", {24, 23, 25, 25}, 97.5))
        myStudents.Add(New clsStudent("T.O.", "Kiliens", {23, 19, 18, 18}, 73.0))
        myStudents.Add(New clsStudent("E.R.", "Manrose", {23, 24, 25, 23}, 84.0))
        myStudents.Add(New clsStudent("W.A.", "Nelson", {23, 24, 25, 23}, 87.0))
        myStudents.Add(New clsStudent("K.U.", "Quaras", {23, 24, 25, 23}, 96.5))
    End Sub
    Public Function getExcelReference()
        Dim CheckExcel As Object
        Dim anExcelDoc As Excel.Application

        'see if excel is already open
        Try
            CheckExcel = GetObject(, "Excel.Application")
        Catch ex As Exception

        End Try

        'see if we found a running instance of excel
        If CheckExcel Is Nothing Then
            anExcelDoc = New Excel.Application()
        Else
            anExcelDoc = CheckExcel
        End If

        'show excel
        anExcelDoc.visible = True

        'we want to add new workbook and sheet
        anExcelDoc.Workbooks.Add()
        anExcelDoc.Sheets.Add()

        'while here we might as well add the column headers
        anExcelDoc.Cells(1, 1) = "Initials"
        anExcelDoc.Cells(1, 2) = "Name"
        anExcelDoc.Cells(1, 3) = "Grade 1"
        anExcelDoc.Cells(1, 4) = "Grade 2"
        anExcelDoc.Cells(1, 5) = "Grade 3"
        anExcelDoc.Cells(1, 6) = "Grade 4"
        anExcelDoc.Cells(1, 7) = "Grade Total"
        anExcelDoc.Cells(1, 8) = "Exam"
        anExcelDoc.Cells(1, 9) = "Final Grade"

        'and since this subroutine is called after the students are loaded we can
        'add the row titles for average, stdev, min and max.
        anExcelDoc.Cells(2, myStudents.Count + 1) = "Aver:"
        anExcelDoc.Cells(2, myStudents.Count + 2) = "St Dev:"
        anExcelDoc.Cells(2, myStudents.Count + 3) = "Min:"
        anExcelDoc.Cells(2, myStudents.Count + 4) = "Max:"

        Return anExcelDoc
    End Function
    Public Sub loadStudentData(anExcelDoc As Excel.Application)
        For i As Integer = 2 To myStudents.Count + 2
            Dim currStudent As clsStudent = myStudents(i)
            anExcelDoc.Cells(i, 1) = currStudent.getInitials
            anExcelDoc.Cells(i, 2) = currStudent.getLastName
            anExcelDoc.Cells(i, 3) = (currStudent.getScores)(0)
            anExcelDoc.Cells(i, 4) = (currStudent.getScores)(1)
            anExcelDoc.Cells(i, 5) = (currStudent.getScores)(2)
            anExcelDoc.Cells(i, 6) = (currStudent.getScores)(3)
            anExcelDoc.Cells(i, 7) = String.Format("=SUM(C{0}:F{0})", i)
            anExcelDoc.Cells(i, 8) = currStudent.getExam
            anExcelDoc.Cells(i, 9) = String.Format("=ROUND((0.4 * G{0}) + (0.6 * H{0}))")
        Next
    End Sub
    Public Sub frm1_load(sender As Object, e As EventArgs) Handles Me.Load
        populateList()
        Dim anExcelDoc = getExcelReference()
        loadStudentData(anExcelDoc)
    End Sub
End Class
