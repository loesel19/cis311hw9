Imports Microsoft.Office.Interop
Public Class Form1
    '--------------------------------------------------------------------------------
    '-                      File Name: Form1                                        -
    '-                      Part of Project: Excel Linking Application (cis311 HW9) -
    '--------------------------------------------------------------------------------
    '-                      Written By: Andrew A. Loesel                            -
    '-                      Written On: April 7, 2022                               -
    '--------------------------------------------------------------------------------
    '- File Purpose:                                                                -
    '-                                                                              -
    '- This file contains all functionality of the program. We take care of all I/O -
    '- in this File. 
    '--------------------------------------------------------------------------------
    '- Program Purpose:                                                             -
    '-                                                                              -
    '- The purpose of this program is to load some data into a list on the front end-
    '- and then send that data to an excel sheet. We also will program formulas into-
    '- the excel sheet to take care of averages, stdevs, min and max for the student-
    '- scores that we are working with.                                             -
    '--------------------------------------------------------------------------------
    '- Global Variable Dictionary (alphabetically):                                 -
    '- myStudents - a generic list of clsStudents. Holds all the student data we    -
    '-              work with in this application.                                  -
    '--------------------------------------------------------------------------------

    'GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES
    'GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES GLOBAL VARIABLES

    Dim myStudents As List(Of clsStudent)

    'SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS
    'SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS SUBPROGRAMS
    Public Sub populateList()
        '------------------------------------------------------------------------------
        '-                      Subprogram Name: populateList                         -
        '------------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                          -
        '-                      Written On: April 7, 2022                             -
        '------------------------------------------------------------------------------
        '- Subprogram Purpose:                                                        -
        '-                                                                            -
        '- The purpose of this subroutine is to load in some default data to our list -
        '- of students.                                                               -
        '------------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                                 -
        '- None                                                                       -
        '------------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                                -
        '- None                                                                       -
        '------------------------------------------------------------------------------
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
        myStudents.Add(New clsStudent("A.A.", "Loesel", {23, 25, 25, 25}, 100))
    End Sub
    Public Function getExcelReference()
        '------------------------------------------------------------------------------
        '-                      Subprogram Name: getExcelReference                    -
        '------------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                          -
        '-                      Written On: April 7, 2022                             -
        '------------------------------------------------------------------------------
        '- Subprogram Purpose:                                                        -
        '-                                                                            -
        '- The purpose of this subroutine is to see if excel is already in the device -
        '- memory. If it is we grab that reference, if not we create a new one. We    -
        '- then add a new workbook and sheet to excel to display our new data.        -
        '------------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                                 -
        '- None                                                                       -
        '------------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                                -
        '- anExcelDoc - a reference to the excel application that we will work with.  -
        '- checkExcel - the object we try to grab an existing excel application with. -
        '------------------------------------------------------------------------------
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
            'excel is already open so we can just add a sheet

        End If


        'we want to add new workbook and sheet
        anExcelDoc.Workbooks.Add()
        anExcelDoc.Sheets.Add()



        Return anExcelDoc
    End Function
    Public Sub loadStudentData(anExcelDoc As Excel.Application)
        '------------------------------------------------------------------------------
        '-                      Subprogram Name: loadStudentData                      -
        '------------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                          -
        '-                      Written On: April 7, 2022                             -
        '------------------------------------------------------------------------------
        '- Subprogram Purpose:                                                        -
        '-                                                                            -
        '- This subprograms purpose is to load student data into the excel sheet. We  -
        '- first loop through all of our students and programatically add their data  -
        '- into the corresponding cells. We then add some headers in fixed positions  -
        '- on the excel sheet.                                                        -
        '------------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                                 -
        '- anExcelDoc - a reference to the excel object that we will be adding data   -
        '- to.                                                                        -
        '------------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                                -
        '- currStudent - the Current clsStudent whose data we are putting in the cells-
        '-               of our excel page in the for loop.                           -
        '------------------------------------------------------------------------------
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

    End Sub

    Public Sub putStatisticalFormulas(anExcelDoc As Excel.Application)
        '------------------------------------------------------------------------------
        '-                      Subprogram Name: putStatisticalFormulas               -
        '------------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                          -
        '-                      Written On: April 7, 2022                             -
        '------------------------------------------------------------------------------
        '- Subprogram Purpose:                                                        -
        '-                                                                            -
        '- The porpose of this subprogram is to put excel formulas in the proper cells-
        '- so that excel can handle average, standard deviation, min and max calculati-
        '- ons for us directly on the sheet. We then make the sheet visible.          -
        '------------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                                 -
        '- anExcelDoc - a reference to the excel object that we will be adding data   -
        '- to.                                                                        -
        '------------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                                -
        '- None                                                                       -
        '------------------------------------------------------------------------------
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

        'show excel
        anExcelDoc.Visible = True
    End Sub
    Public Sub displayDataList()
        '------------------------------------------------------------------------------
        '-                      Subprogram Name: displayDataList                      -
        '------------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                          -
        '-                      Written On: April 7, 2022                             -
        '------------------------------------------------------------------------------
        '- Subprogram Purpose:                                                        -
        '-                                                                            -
        '- The purpose of this subprogram is display all the students in myStudents   -
        '- in our listbox. We change the listbox font to a monospace font where every -
        '- character is the same width. We then loop through each student and add     -
        '- their information to a formatted string which is then added to our listbox -
        '- items.                                                                     -
        '------------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                                 -
        '- None                                                                       -
        '------------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                                -
        '- str - this will be a formatted string that we will add to the listbox.     -
        '------------------------------------------------------------------------------
        'change font of listbox to a monospace font like Consolas to make the listbox look tidier
        lstStudentData.Font = New Font("Consolas", 12, FontStyle.Regular)
        'since we need to format the student data in the listbox we will need to string.format it
        Dim str As String
        For Each student As clsStudent In myStudents
            str = String.Format("{0,-4}   {1,-12}{2,-17}{3,-3}",
                                student.getInitials, student.getLastName, student.getScores(0) & ", " & student.getScores(1) &
                               ", " & student.getScores(2) & ", " & student.getScores(3), student.getExam)
            lstStudentData.Items.Add(str)
        Next


    End Sub

    Public Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        '------------------------------------------------------------------------------
        '-                      Subprogram Name: btnAdd_Click                         -
        '------------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                          -
        '-                      Written On: April 7, 2022                             -
        '------------------------------------------------------------------------------
        '- Subprogram Purpose:                                                        -
        '-                                                                            -
        '- The purpose of this subprogram is to add a new student into our list and   -
        '- make sure their data is displayed and used in excel. We first try to create-
        '- a new student object from the textbox values, if an exception is triggered -
        '- during this we print out a message to tell the user to use the specified   -
        '- format in the textbox hints. We then clear all of our textboxes.           -
        '------------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                                 -
        '- sender - identifies which control used the event.                          -
        '- e - Holds the EventArgs object sent to the routine.                        -
        '------------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                                -
        '- newStudent - this is a clsStudent object that we want to add to our list.  -
        '- scores() - an integer array of the students homework scores.               -
        '- strScores() - a string array that we get from splitting the textbox for    -
        '-               score input by comma.                                        -
        '------------------------------------------------------------------------------
        Try
            Dim strScores() = txtScores.Text.Split(",")
            Dim scores() As Integer = {CInt(strScores(0)), CInt(strScores(1)), CInt(strScores(2)), CInt(strScores(3))}
            Dim newStudent As New clsStudent(txtInitials.Text, txtLastName.Text, scores, CInt(txtExam.Text))
            myStudents.Add(newStudent)
            lstStudentData.Items.Clear()
            displayDataList()
        Catch ex As Exception
            MessageBox.Show("input in the format specified by the hints.", "Incorrect Input")
        End Try
        'clear textboxes
        txtInitials.Clear()
        txtLastName.Clear()
        txtScores.Clear()
        txtExam.Clear()


    End Sub
    Public Sub btnViewInExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        '------------------------------------------------------------------------------
        '-                      Subprogram Name: btnViewInExcel_Click                 -
        '------------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                          -
        '-                      Written On: April 7, 2022                             -
        '------------------------------------------------------------------------------
        '- Subprogram Purpose:                                                        -
        '-                                                                            -
        '- The purpose of this subprogram is to open up the excel document that we    -
        '- have a reference to. so we first get that reference, then we load the      -
        '- student data into the sheet, and then we put the formulas in the sheet.    -
        '- putStatisticalFormulas() will make the sheet visible.                      -
        '------------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                                 -
        '- sender - identifies which control used the event.                          -
        '- e - Holds the EventArgs object sent to the routine.                        -
        '------------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                                -
        '- anExcelDoc - a reference to the excel object that we will be adding data   -
        '- to.                                                                        -
        '------------------------------------------------------------------------------
        Dim anExcelDoc = getExcelReference()
        loadStudentData(anExcelDoc)
        putStatisticalFormulas(anExcelDoc)
    End Sub

    Public Sub frm1_load(sender As Object, e As EventArgs) Handles Me.Load
        '------------------------------------------------------------------------------
        '-                      Subprogram Name: frm1_load                            -
        '------------------------------------------------------------------------------
        '-                      Written By: Andrew A. Loesel                          -
        '-                      Written On: April 7, 2022                             -
        '------------------------------------------------------------------------------
        '- Subprogram Purpose:                                                        -
        '-                                                                            -
        '- The purpose of this subprogram is to control the program when the form laod-
        '- s. We just call populateList to get our student data loaded into the list. -
        '- we then get an excel reference and display the list data in the listbox    -
        '- by calling displayDataList.                                                -
        '------------------------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):                                 -
        '- sender - identifies which control used the event.                          -
        '- e - Holds the EventArgs object sent to the routine.                        -
        '------------------------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):                                -
        '- anExcelDoc - a reference to the excel object that we will be adding data   -
        '- to.                                                                        -
        '------------------------------------------------------------------------------
        populateList()
        Dim anExcelDoc = getExcelReference()
        displayDataList()
    End Sub
End Class
