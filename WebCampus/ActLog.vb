Option Explicit On
Option Compare Text

Imports System.Data
Imports System.Data.SqlClient

Public Class ActLog

    Dim mActID As Integer
    Dim mDetail As String
    Dim mActDate As Date
    Dim mLoginID As String

    Public Property ActID() As Integer
        Get
            ActID = mActID
        End Get
        Set(ByVal value As Integer)
            mActID = value
        End Set
    End Property

    Public Property Detail() As String
        Get
            Detail = mDetail
        End Get
        Set(ByVal value As String)
            mDetail = value
        End Set
    End Property

    Public Property LoginID() As String
        Get
            LoginID = mLoginID
        End Get
        Set(ByVal value As String)
            mLoginID = value
        End Set
    End Property

    Public Property ActDate() As Date
        Get
            ActDate = mActDate
        End Get
        Set(ByVal value As Date)
            mActDate = value
        End Set
    End Property

End Class
