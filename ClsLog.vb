'Author: Yan Naing Aye
'WebSite: http://cool-emerald.blogspot.sg/
'Updated: 2009 June 25
'-----------------------------------------------------------------------------
Imports System.IO
Public Class ClsLog
    Private mEnableLog As Boolean = False
    Private mLogFileDirectory As String = ""
    Private mLogFilePath As String = ""
    Private mLogFileLifeSpan As Integer = 0
    Private mLogFileFirstName As String = "AppName"
    Public Sub New()
        mEnableLog = False
        mLogFileDirectory = My.Application.Info.DirectoryPath
        mLogFileLifeSpan = 0
        mLogFileFirstName = My.Application.Info.AssemblyName
    End Sub
    Public Property LogFileLifeSpan() As Integer
        Get
            Return mLogFileLifeSpan
        End Get
        Set(ByVal value As Integer)
            mLogFileLifeSpan = IIf(value >= 0, value, 0)
        End Set
    End Property
    Public Property LogFileFirstName() As String
        Get
            Return mLogFileFirstName
        End Get
        Set(ByVal value As String)            
            mLogFileFirstName = value
        End Set
    End Property
    Public Property LogEnable() As Boolean
        Get
            Return mEnableLog
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                If Directory.Exists(Me.LogFileDirectory) Then
                    mEnableLog = value
                Else
                    mEnableLog = False
                    Throw New Exception("Invalid file directory.")
                End If
            Else
                mEnableLog = value
            End If
        End Set
    End Property
    Public Property LogFileDirectory() As String
        Get
            Return mLogFileDirectory
        End Get
        Set(ByVal value As String)
            value = Trim(value)
            If Directory.Exists(value) Then
                Dim i As Integer = value.Length - 1
                If (((value(i)) = "\") OrElse ((value(i)) = "/")) Then
                    value = value.Substring(0, i)
                End If
                mLogFileDirectory = value
            Else
                Throw New Exception("Invalid file directory.")
            End If
        End Set
    End Property
    Public ReadOnly Property LogFilePath() As String
        Get
            Return mLogFileDirectory & "\" & mLogFileFirstName & Format(Now, "-yyyy-MMM-dd") & ".log"
        End Get
    End Property
    Public Sub WriteLog(ByVal LogEntry As String)
        If mEnableLog = True Then
            mLogFilePath = mLogFileDirectory & "\" & mLogFileFirstName & Format(Now, "-yyyy-MMM-dd") & ".log"
            Dim objStreamWriter As StreamWriter = New StreamWriter(mLogFilePath, True)
            Try
                objStreamWriter.WriteLine(LogEntry)
            Catch ex As Exception
            Finally
                objStreamWriter.Close()
            End Try
        End If
    End Sub
    Public Sub WriteTimeAndLog(ByVal LogEntry As String)
        WriteLog(Now.ToLongTimeString & " " & LogEntry)
    End Sub
    Public Sub CleanupFiles()
        If mLogFileLifeSpan > 0 Then 'if life span is zero, there will be no cleaning up
            Try
                Dim LogFiles() As String = Directory.GetFiles(mLogFileDirectory)
                For Each LogFile As String In LogFiles
                    If (DateDiff("d", File.GetLastWriteTime(LogFile), Now) > mLogFileLifeSpan) _
                    AndAlso (Right(LogFile, 4) = ".log") Then
                        File.Delete(LogFile)
                    End If
                Next
            Catch ex As Exception
            End Try
        End If
    End Sub
End Class
