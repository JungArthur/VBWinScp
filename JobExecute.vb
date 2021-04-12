
Public Class JobExecute

    Dim Jobs As List(Of xmlJob)

    Public Sub New(job As xmlJob)
        Me.Jobs = New List(Of xmlJob)
        Me.Jobs.Add(job)
        'Logger.Log(enLogLevel.Info, $"Job 실행 Job 개수 : {Me.Jobs.Count}")
    End Sub

    Public Sub New(list As List(Of xmlJob))
        Me.Jobs = list
        'Logger.Log(enLogLevel.Info, $"Job 실행 Job 개수 : {Me.Jobs.Count}")
    End Sub

    Public Sub TestPoint()

        Logger.Log(enLogLevel.Info, $"Job 시작")

        For Each job As xmlJob In Jobs

            Dim sqlJob As SqlSelectService = SqlSelectPoint(job)

            If sqlJob.Result Then

                Dim csvJob As CsvSaveService = SaveCsvPoint(job, sqlJob)

                Dim sftpJob As SftpSendService = SftpSendPoint(job)

                If Not String.IsNullOrWhiteSpace(job._JOB_INSERT_SERVER_NAME) And Not String.IsNullOrWhiteSpace(job._JOB_INSERT_TABLE_NAME) And sqlJob.contentVList.Count > 0 Then
                    Dim insertJob As SqlInsertService = SqlInsertPoint(job, sqlJob)
                End If

            End If

        Next

        Logger.Log(enLogLevel.Info, $"Job 종료")

    End Sub

    Public Sub StartPoint()

        For Each job As xmlJob In Jobs

            Dim sqlJob As SqlSelectService = SqlSelectPoint(job)

            If sqlJob.Result Then

                Dim csvJob As CsvSaveService = SaveCsvPoint(job, sqlJob)

                Dim sftpJob As SftpSendService = SftpSendPoint(job)

                If Not String.IsNullOrWhiteSpace(job._JOB_INSERT_SERVER_NAME) And Not String.IsNullOrWhiteSpace(job._JOB_INSERT_TABLE_NAME) And sqlJob.contentVList.Count > 0 Then
                    Dim insertJob As SqlInsertService = SqlInsertPoint(job, sqlJob)
                End If

            End If

        Next

        Logger.Log(enLogLevel.Info, $"Job 종료")
    End Sub

    Private Function SqlSelectPoint(job As xmlJob) As SqlSelectService
        Dim sqlJob As SqlSelectService = New SqlSelectService(job)
        If Not IsNothing(sqlJob) Then
            Return sqlJob
        Else
            Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE} SqlPoint Return Nothing")
            Return Nothing
        End If
    End Function

    Private Function SaveCsvPoint(job As xmlJob, sqlJob As SqlSelectService) As CsvSaveService
        Dim csvJob As CsvSaveService = New CsvSaveService(job, sqlJob)
        If Not IsNothing(csvJob) Then
            Return csvJob
        Else
            Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE}  SaveCsvPoint Return Nothing")
            Return Nothing
        End If
    End Function


    Private Function SftpSendPoint(job As xmlJob) As SftpSendService
        Dim sftpJob As SftpSendService = New SftpSendService(job)
        If Not IsNothing(sftpJob) Then
            Return sftpJob
        Else
            Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE} SftpPoint Return Nothing")
            Return Nothing
        End If
    End Function


    Private Function SqlInsertPoint(job As xmlJob, sqlJob As SqlSelectService) As SqlInsertService
        '실행할지말지 결정하는 인자 추가
        Dim insertResult As New SqlInsertService(job, sqlJob)
        If Not IsNothing(insertResult) Then
            Return insertResult
        Else
            Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE} SqlInsertPoint Return Nothing")
            Return Nothing
        End If
    End Function




End Class
