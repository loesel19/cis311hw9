Imports Microsoft.Office.Interop
Public Class Form1
    Dim myStudents As List(Of clsStudent)

    Public Sub populateList()
        'make a new reference for student list
        myStudents = New List(Of clsStudent)
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
            '  CheckExcel = GetObject(, "Excel.Application")
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


        Return anExcelDoc
    End Function
    Public Sub loadStudentData(anExcelDoc As Excel.Application)
        'loop through each student in the list, we start at 2 for our loop counter since that will be the row
        'we start entering data at. We then can just use our getter methods to get that data and place it in the cells
        'the total grade and final grade are both functions that we place in the cells
        For i As Integer = 2 To myStudents.Count + 1
            Dim currStudent As clsStudent = myStudents(i - 2)
            anExcelDoc.Cells(i, 1) = currStudent.getInitials
            anExcelDoc.Cells(i, 2) = currStudent.getLastName
            anExcelDoc.Cells(i, 3) = (currStudent.getScores)(0)
            anExcelDoc.Cells(i, 4) = (currStudent.getScores)(1)
            anExcelDoc.Cells(i, 5) = (currStudent.getScores)(2)
            anExcelDoc.Cells(i, 6) = (currStudent.getScores)(3)
            anExcelDoc.Cells(i, 7) = String.Format("=SUM(C{0}:F{0})", i)
            anExcelDoc.Cells(i, 8) = currStudent.getExam
            anExcelDoc.Cells(i, 9) = String.Format("=ROUND((0.4 * G{0}) + (0.6 * H{0}), 1)", i)
        Next
        'we can add row titles for statistics now to get it out of the way as well
        anExcelDoc.Cells(myStudents.Count + 3, 2) = "Aver:"
        anExcelDoc.Cells(myStudents.Count + 4, 2) = "St Dev:"
        anExcelDoc.Cells(myStudents.Count + 5, 2) = "Min:"
        anExcelDoc.Cells(myStudents.Count + 6, 2) = "Max:"
    End Sub

    Public Sub putStatisticalFormulas(anExcelDoc As Excel.Application)
        'add in average, stdev, min and max functions in the corresponding cells
        'all + 3 rows are for average
        anExcelDoc.Cells(myStudents.Count + 3, 3) = String.Format("=AVERAGE(C2:C{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 3, 4) = String.Format("=AVERAGE(D2:D{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 3, 5) = String.Format("=AVERAGE(E2:E{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 3, 6) = String.Format("=AVERAGE(F2:F{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 3, 7) = String.Format("=AVERAGE(G2:G{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 3, 8) = String.Format("=AVERAGE(H2:H{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 3, 9) = String.Format("=AVERAGE(I2:I{0})", myStudents.Count + 1)
        '+4 for stdev
        anExcelDoc.Cells(myStudents.Count + 4, 3) = String.Format("=STDEV(C2:C{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 4, 4) = String.Format("=STDEV(D2:D{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 4, 5) = String.Format("=STDEV(E2:E{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 4, 6) = String.Format("=STDEV(F2:F{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 4, 7) = String.Format("=STDEV(G2:G{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 4, 8) = String.Format("=STDEV(H2:H{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 4, 9) = String.Format("=STDEV(I2:I{0})", myStudents.Count + 1)
        '+5 For Min
        anExcelDoc.Cells(myStudents.Count + 5, 3) = String.Format("=MIN(C2:C{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 5, 4) = String.Format("=MIN(D2:D{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 5, 5) = String.Format("=MIN(E2:E{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 5, 6) = String.Format("=MIN(F2:F{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 5, 7) = String.Format("=MIN(G2:G{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 5, 8) = String.Format("=MIN(H2:H{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 5, 9) = String.Format("=MIN(I2:I{0})", myStudents.Count + 1)
        '+6 for Max
        anExcelDoc.Cells(myStudents.Count + 6, 3) = String.Format("=MAX(C2:C{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 6, 4) = String.Format("=MAX(D2:D{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 6, 5) = String.Format("=MAX(E2:E{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 6, 6) = String.Format("=MAX(F2:F{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 6, 7) = String.Format("=MAX(G2:G{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 6, 8) = String.Format("=MAX(H2:H{0})", myStudents.Count + 1)
        anExcelDoc.Cells(myStudents.Count + 6, 9) = String.Format("=MAX(I2:I{0})", myStudents.Count + 1)
    End Sub
    Public Sub frm1_load(sender As Object, e As EventArgs) Handles Me.Load
        populateList()
        Dim anExcelDoc = getExcelReference()
        loadStudentData(anExcelDoc)
        putStatisticalFormulas(anExcelDoc)
    End Sub
End Class
