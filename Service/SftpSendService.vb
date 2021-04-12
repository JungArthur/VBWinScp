
Imports System.IO
Imports System.Net
Imports System.Text
Imports WinSCP
Public Class SftpSendService
    Public Property RemoteFullPath As String
    Public Property Result As Boolean = False

    Public Sub New(job As xmlJob)

        Dim localFileName As String = $"{FileNameDeterminer.GetResultFileName(job._JOB_FILE_PREFIX, job._JOB_FILE_DATE_FORMAT, enFileExtension.csv)}"

        Dim localFullPath As String = FileNameDeterminer.GetDateDirectoryPath(job._JOB_TITLE) & localFileName
        RemoteFullPath = FileNameDeterminer.RemoteDirectoryPath(job._JOB_SFTP_REMOTE_DIRECTORY) & localFileName

        If job._JOB_SFTP_METHOD = "SFTP" Then

            '=======================================================================
            Try
                ' Setup session options
                Dim sessionOptions As New SessionOptions
                With sessionOptions
                    .Protocol = Protocol.Sftp
                    .HostName = job._JOB_SFTP_SERVER_NAME
                    .PortNumber = job._JOB_SFTP_SERVER_PORT
                    .UserName = job._JOB_SFTP_USER_ID
                    .Password = job._JOB_SFTP_PASSWORD
                    .GiveUpSecurityAndAcceptAnySshHostKey = True
                End With

                Using session As New Session
                    ' Connect
                    session.Open(sessionOptions)

                    ' Upload files
                    Dim transferOptions As New TransferOptions
                    transferOptions.TransferMode = TransferMode.Binary

                    Dim transferResult As TransferOperationResult
                    transferResult =
                    session.PutFiles(localFullPath, RemoteFullPath, False, transferOptions)

                    ' Throw on any error
                    transferResult.Check()

                    ' Print results
                    For Each transfer In transferResult.Transfers
                        Logger.Log(enLogLevel.Success, $"{job._JOB_TITLE} SFTP 업로드 성공")
                    Next
                End Using

            Catch ex As Exception
                Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE} SFTP 업로드 실패 / {ex}")
            End Try
            '=======================================================================

        ElseIf job._JOB_SFTP_METHOD = "FTPS_Explicit" Then
            '=======================================================================
            Try
                ' Setup session options
                Dim sessionOptions As New SessionOptions
                With sessionOptions
                    .Protocol = Protocol.Ftp
                    .FtpSecure = FtpSecure.Explicit
                    .HostName = job._JOB_SFTP_SERVER_NAME
                    .PortNumber = job._JOB_SFTP_SERVER_PORT
                    .UserName = job._JOB_SFTP_USER_ID
                    .Password = job._JOB_SFTP_PASSWORD
                    .GiveUpSecurityAndAcceptAnyTlsHostCertificate = True
                End With

                Using session As New Session
                    ' Connect
                    session.Open(sessionOptions)

                    ' Upload files
                    Dim transferOptions As New TransferOptions
                    transferOptions.TransferMode = TransferMode.Binary

                    Dim transferResult As TransferOperationResult
                    transferResult =
                    session.PutFiles(localFullPath, RemoteFullPath, False, transferOptions)

                    ' Throw on any error
                    transferResult.Check()

                    ' Print results
                    For Each transfer In transferResult.Transfers
                        Logger.Log(enLogLevel.Success, $"{job._JOB_TITLE} FTPS 업로드 성공")
                    Next
                End Using

            Catch ex As Exception
                Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE} FTPS 업로드 실패 / {ex}")
            End Try
            '=======================================================================
        ElseIf job._JOB_SFTP_METHOD = "FTPS_Implicit" Then
            '=======================================================================
            Try
                ' Setup session options
                Dim sessionOptions As New SessionOptions
                With sessionOptions
                    .Protocol = Protocol.Ftp
                    .FtpSecure = FtpSecure.Implicit
                    .HostName = job._JOB_SFTP_SERVER_NAME
                    .PortNumber = job._JOB_SFTP_SERVER_PORT
                    .UserName = job._JOB_SFTP_USER_ID
                    .Password = job._JOB_SFTP_PASSWORD
                    .GiveUpSecurityAndAcceptAnyTlsHostCertificate = True
                End With

                Using session As New Session
                    ' Connect
                    session.Open(sessionOptions)

                    ' Upload files
                    Dim transferOptions As New TransferOptions
                    transferOptions.TransferMode = TransferMode.Binary

                    Dim transferResult As TransferOperationResult
                    transferResult =
                    session.PutFiles(localFullPath, RemoteFullPath, False, transferOptions)

                    ' Throw on any error
                    transferResult.Check()

                    ' Print results
                    For Each transfer In transferResult.Transfers
                        Logger.Log(enLogLevel.Success, $"{job._JOB_TITLE} FTPS 업로드 성공")
                    Next
                End Using

            Catch ex As Exception
                Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE} FTPS 업로드 실패 / {ex}")
            End Try
            '=======================================================================
        ElseIf job._JOB_SFTP_METHOD = "FTP" Then

            Try
                ' Get the object used to communicate with the server.
                Dim request As FtpWebRequest = CType(WebRequest.Create($"ftp://{job._JOB_SFTP_SERVER_NAME}:{job._JOB_SFTP_SERVER_PORT}/{RemoteFullPath}"), FtpWebRequest)
                request.Method = WebRequestMethods.Ftp.UploadFile
                request.EnableSsl = False
                request.Credentials = New NetworkCredential(job._JOB_SFTP_USER_ID, job._JOB_SFTP_PASSWORD)
                'Only FTPS End

                ' Copy the contents of the file to the request stream.
                Dim fileContents As Byte()

                Using sourceStream As StreamReader = New StreamReader(localFullPath)
                    fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd())
                End Using

                request.ContentLength = fileContents.Length

                Using requestStream As Stream = request.GetRequestStream()
                    requestStream.Write(fileContents, 0, fileContents.Length)
                End Using

                Using response As FtpWebResponse = CType(request.GetResponse(), FtpWebResponse)
                    Console.WriteLine($"Upload File Complete, status {response.StatusDescription}")
                End Using
                Logger.Log(enLogLevel.Success, $"{job._JOB_TITLE} FTP 업로드 성공")
            Catch ex As Exception
                Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE} FTP 업로드 실패 / {ex}")
            End Try

        Else
            Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE} 파일 업로드 실패 / 선택된 Method는 지원하지 않음")
        End If

    End Sub

End Class
