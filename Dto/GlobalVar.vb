Imports System.Timers
Imports System.IO

Public NotInheritable Class GlovalVar
    '로그 화면 전달을 위한 이벤트

End Class



#Region "Enum"
Public Enum enLogLevel
    Success
    Fail
    Info
End Enum

Public Enum enFileExtension
    csv
    log
End Enum

Public Enum enFileTransferMethod

    SFTP
    FTPS_Explicit
    FTPS_Implicit
    FTP
End Enum
#End Region

#Region "Global Event Trigger Timer"
Public NotInheritable Class adjustTimer
    Public Shared timer As Timer = New Timer()

    Public Shared Sub setInteval(value As Double)
        timer.Interval = value
    End Sub
End Class
#End Region

#Region "File Name Determiner"
Public NotInheritable Class FileNameDeterminer



#Region "Get Result File"
    Public Shared Function GetResultFileName(prefix As String, dateFormat As String) As String
        If String.IsNullOrWhiteSpace(dateFormat) Then
            dateFormat = "yyyy-MM-dd"
        End If

        Dim yesterdayDate As String = DateTime.Today.AddDays(-1).ToString(dateFormat)

        Dim localFileName As String = $"{prefix}_{yesterdayDate}"

        Return localFileName
    End Function

    Public Shared Function GetResultFileName(prefix As String, dateFormat As String, extension As enFileExtension) As String
        If String.IsNullOrWhiteSpace(dateFormat) Then
            dateFormat = "yyyy-MM-dd"
        End If

        Dim yesterdayDate As String = DateTime.Today.AddDays(-1).ToString(dateFormat)

        Dim localFileName As String = $"{prefix}_{yesterdayDate}.{extension}"

        Return localFileName
    End Function
#End Region

#Region "Get Directory Path"

    Public Shared Function GetDateDirectoryPath(title As String) As String

        Dim year As String = DateTime.Today.AddDays(-1).ToString("yyyy")
        Dim month As String = DateTime.Today.AddDays(-1).ToString("MM")

        Dim DirPath As String = AppDomain.CurrentDomain.BaseDirectory & $"{title}\{year}\{month}\"

        Return DirPath
    End Function

    Public Shared Function RemoteDirectoryPath(remoteDirectory As String) As String

        ' 마지막 글자가 \가 아니면 추가함.
        If remoteDirectory.Length > 0 Then
            If remoteDirectory(remoteDirectory.Length - 1) <> "/" Then
                remoteDirectory = remoteDirectory & "/"
            End If
        End If
        Return remoteDirectory
    End Function

#End Region

End Class
#End Region

#Region "LogOverLoading"
Public NotInheritable Class Logger

    Public Delegate Sub delLogSender(eLevel As enLogLevel, strLog As String)
    Public Shared Event eLogSender As delLogSender

    Public Shared Sub Log(eLevel As enLogLevel, LogDesc As String)
        Dim dTime As DateTime = DateTime.Now
        Dim LogInfo As String = $"{dTime: yyyy-MM-dd hh:mm:ss.fff} [{eLevel.ToString()}] {LogDesc}"
        fn_LogWrite(LogInfo, eLevel)
        fn_LogConsole(LogInfo)
        RaiseEvent eLogSender(eLevel, LogDesc)
    End Sub

    Public Shared Sub Log(dTime As DateTime, eLevel As enLogLevel, LogDesc As String)
        Dim LogInfo As String = $"{dTime: yyyy-MM-dd hh:mm:ss.fff} [{eLevel.ToString()}] {LogDesc}"
        fn_LogWrite(LogInfo, eLevel)
        fn_LogConsole(LogInfo)
    End Sub

    Private Shared Sub fn_LogConsole(str As String)
        Console.WriteLine(str)
    End Sub

    Private Shared Sub fn_LogWrite(str As String, eLevel As enLogLevel)

        Dim year As String = DateTime.Today.ToString("yyyy")
        Dim month As String = DateTime.Today.ToString("MM")
        Dim day As String = DateTime.Today.ToString("dd")

        Dim DirPath As String = AppDomain.CurrentDomain.BaseDirectory & $"\Service\{year}\{month}"

        Dim FilePath As String = DirPath & $"\{eLevel}_{year}{month}{day}.log"

        Dim di As DirectoryInfo = New DirectoryInfo(DirPath)
        Dim fi As FileInfo = New FileInfo(FilePath)
        Try

            If Not di.Exists Then
                Directory.CreateDirectory(DirPath)
            End If

            If Not fi.Exists Then
                Using sw As StreamWriter = New StreamWriter(FilePath)
                    sw.WriteLine(str)
                    sw.Close()

                End Using
            Else
                Using sw As StreamWriter = File.AppendText(FilePath)
                    sw.WriteLine(str)
                    sw.Close()
                End Using
            End If

        Catch ex As Exception

        End Try

    End Sub

End Class
#End Region