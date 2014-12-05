Option Explicit On
Option Compare Text

Imports System.Data
Imports System.Data.SqlClient

Public Class Course

    Dim mCourseID As Integer
    Dim mCourseCode As String
    Dim mCourseName As String
    Dim mCourseDuration As Integer

    Public Property CourseID() As Integer
        Get
            CourseID = mCourseID
        End Get
        Set(ByVal value As Integer)
            mCourseID = value
            LoadData(mCourseID)

        End Set
    End Property

    Public Property CourseCode() As String
        Get
            CourseCode = mCourseCode
        End Get
        Set(ByVal value As String)
            mCourseCode = value
        End Set
    End Property

    Public Property CourseName() As String
        Get
            CourseName = mCourseName
        End Get
        Set(ByVal value As String)
            mCourseName = value
        End Set
    End Property

    Public Property CourseDuration() As Integer
        Get
            CourseDuration = mCourseDuration
        End Get
        Set(ByVal value As Integer)
            mCourseDuration = value
        End Set
    End Property

    Private Sub pLoadDefaults()
        mCourseID = 0
        mCourseCode = ""
        mCourseName = ""
        mCourseDuration = 0
    End Sub

    Public Sub LoadData(Optional ByVal aCourseID As String = "", Optional ByVal ErrDetail As String = "")
        Dim iRS As SqlDataReader = Nothing

        If aCourseID <> vbNullString Then mCourseID = aCourseID
        If mCourseID = vbNullString Then ErrDetail = "Invalid CourseID" : Exit Sub

        Try
            iRS = New SqlCommand("SELECT * FROM Course WHERE CourseID='" & mCourseID & "'", gConnection).ExecuteReader

            iRS.Read()
            mCourseCode = iRS("CourseCode").ToString
            mCourseName = iRS("CourseName").ToString
            mCourseDuration = iRS("CourseDuration").ToString

            iRS.Close()
            iRS = Nothing

        Catch ex As Exception
            If Not iRS Is Nothing Then
                If iRS.IsClosed = False Then iRS.Close()
            End If
            ErrDetail = ex.Message
        End Try

    End Sub


    Public Sub Save(ByVal IsNewEntry As Boolean, Optional ByVal ErrDetail As String = "")
        Dim iSQL As String = ""

        Dim iCourseID As Integer
        Dim i As Integer

        Try
            If IsNewEntry Then
                iCourseID = CInt(GetData("SELECT MAX(CourseID) + 1 FROM Course"))
                If iCourseID = 0 Then iCourseID = 1
                iSQL = "INSERT INTO Course (CourseID,CourseCode,CourseName,CourseDuration) VALUES (" & _
                iCourseID & ",'" & mCourseCode & "','" & mCourseName & "'," & mCourseDuration & ")"
                i = New SqlCommand(iSQL, gConnection).ExecuteNonQuery()
                mCourseID = iCourseID
            Else
                iSQL = "UPDATE Course SET CourseName'" & mCourseName & "'" & _
                    ",CourseCode='" & mCourseCode & "'" & _
                    ",CourseDuration=" & mCourseDuration & _
                    " WHERE CourseID=" & mCourseID
                i = New SqlCommand(iSQL, gConnection).ExecuteNonQuery()
            End If

        Catch ex As Exception
            ErrDetail = ex.Message
        End Try

    End Sub

    Public Sub New()
        pLoadDefaults()
    End Sub
End Class
