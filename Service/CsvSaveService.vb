Imports System.IO
Public Class CsvSaveService

    'csvWrite(job._JOB_TITLE, job._JOB_FILE_PREFIX)

    Public Property ResultFilePath As String
    Public Property Result As Boolean = False

    Public Sub New(job As xmlJob, sqlJob As SqlSelectService)

        Dim DirPath As String = FileNameDeterminer.GetDateDirectoryPath(job._JOB_TITLE)
        ResultFilePath = DirPath & $"{FileNameDeterminer.GetResultFileName(job._JOB_FILE_PREFIX, job._JOB_FILE_DATE_FORMAT, enFileExtension.csv)}"

        Try

            Dim di As DirectoryInfo = New DirectoryInfo(DirPath)
            Dim fi As FileInfo = New FileInfo(ResultFilePath)

            If Not di.Exists Then
                Directory.CreateDirectory(DirPath)
            End If

            If Not fi.Exists Then
                Using sw As StreamWriter = New StreamWriter(File.Create(ResultFilePath), Text.Encoding.UTF8) 'New StreamWriter(New FileStream(ResultFilePath, FileMode.CreateNew), Text.Encoding.UTF8)

                    'Head 채우기
                    sw.WriteLine(String.Join(",", sqlJob.headVList))

                    For Each letter As String In sqlJob.contentVList
                        sw.WriteLine(letter)
                    Next

                    sw.Close()
                End Using
            Else
                Using sw As StreamWriter = New StreamWriter(File.Create(ResultFilePath), Text.Encoding.UTF8)
                    'Head 채우기
                    sw.WriteLine(String.Join(",", sqlJob.headVList))

                    For Each letter As String In sqlJob.contentVList
                        sw.WriteLine(letter)
                    Next

                    sw.Close()
                End Using
            End If

            Logger.Log(enLogLevel.Success, $"{job._JOB_TITLE} File 작성 성공")
            Result = True

        Catch ex As Exception

            Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE} File 작성 실패")
            Return
        End Try

    End Sub


End Class
