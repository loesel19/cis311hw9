Public Class clsStudent
    Private initials As String
    Private lastName As String
    Private scores As Integer()
    Private exam As Decimal

    Public Sub setInitials(initials As String)
        Me.initials = initials
    End Sub
    Public Sub setLastName(lastName As String)
        Me.lastName = lastName
    End Sub
    Public Sub setScores(scores As Integer())
        Me.scores = scores
    End Sub
    Public Sub setExam(exam As Decimal)
        Me.exam = exam
    End Sub
    Public Function getInitials() As String
        Return Me.initials
    End Function
    Public Function getLastName() As String
        Return Me.lastName
    End Function
    Public Function getScores() As Integer()
        Return Me.scores
    End Function
    Public Function getExam() As Decimal
        Return Me.exam
    End Function
    Public Sub New(initials As String, lastName As String, scores As Integer(), exam As Decimal)
        setInitials(initials)
        setLastName(lastName)
        setScores(scores)
        setExam(exam)
    End Sub
End Class
