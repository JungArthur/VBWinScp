Imports System.Data.SqlClient
Imports System.Text
Public Class SqlInsertService


    Public Property Result As Boolean = False


    Public Sub New(job As xmlJob, sqlJob As SqlSelectService)

        Dim SqlServerName As String = job._JOB_INSERT_SERVER_NAME
        Dim InitialDB As String = job._JOB_INSERT_INITIAL_DB
        Dim UserID As String = job._JOB_INSERT_USER_ID
        Dim Password As String = job._JOB_INSERT_PASSWORD

        'Dim QueryTemplate As String = $"INSERT INTO PMI_Recon100 (ProductCode, PrdDate, ReceivedQty, ProductionCenterCode, cdt) VALUES (N'1', N'2021-01-01', 1, N'2', GETDATE());"

        'Dim Query As String = "INSERT INTO PMI_Recon100 (ProductCode, PrdDate, ReceivedQty, ProductionCenterCode, cdt) VALUES (N'1', N'2021-01-01', 1, N'2', GETDATE());
        '                        INSERT INTO PMI_Recon100 (ProductCode, PrdDate, ReceivedQty, ProductionCenterCode, cdt) VALUES (N'3', N'2021-01-01', 1, N'4', GETDATE())"

        Dim connectionString As String = "Data Source=" & SqlServerName & ";Initial Catalog=" & InitialDB & ";User ID=" & UserID & ";Password=" & Password & ";Connection Timeout=10"

        Using connection As New SqlConnection(connectionString)
            Dim command As SqlCommand = Nothing

            Try
                command = New SqlCommand(QueryBuilder(job, sqlJob), connection)
                connection.Open()
                command.ExecuteNonQuery()
            Catch e As Exception
                'Logger.Log(enLogLevel.Fail, $"InsertDB SQLSERVER 연결 실패")
                Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE} InsertDB {e}")
                If Not IsNothing(connection) Then
                    connection.Close()
                End If
                Return
            End Try

            Logger.Log(enLogLevel.Success, $"SQLSERVER Insert 쿼리 실행 성공")

            connection.Close()
            Result = True
        End Using

    End Sub
    Public Function QueryBuilder(job As xmlJob, sqlJob As SqlSelectService) As String


        Dim TABLE_NAME As String = job._JOB_INSERT_TABLE_NAME
        Dim COLUMN_NAMES As String = job._JOB_INSERT_COLUMN_NAMES
        Dim CDT_ENABLED As Boolean = Boolean.Parse(job._JOB_INSERT_CDT_ENABLED)

        Dim QueryTemplate As String = "INSERT INTO {0} ({1}) VALUES ({2});"

        Dim sb As New StringBuilder

        'columnList 결정
        Dim columnList As String

        'Head 채우기
        If String.IsNullOrWhiteSpace(COLUMN_NAMES) Then
            ' Defualt ColumnList 사용
            columnList = String.Join(",", sqlJob.headVList)
        Else
            columnList = COLUMN_NAMES
        End If

        'Column List에 CDT 추가여부결정
        If CDT_ENABLED Then
            columnList = columnList & ",cdt"
        End If


        For Each letter As String In sqlJob.contentVList

            Dim strtemp As StringBuilder = New StringBuilder()

            Dim initFlag As Boolean = False
            For Each entity As String In letter.Split(",")
                If initFlag Then
                    '초기 이후 조건
                    strtemp.Append($",N'{entity}'")
                Else
                    '초기 조건
                    initFlag = True
                    strtemp.Append($"N'{entity}'")
                End If
            Next

            If CDT_ENABLED Then
                sb.Append(String.Format(QueryTemplate, TABLE_NAME, columnList, strtemp.ToString() & ", DATEADD(day, -1, GETDATE())"))
            Else
                sb.Append(String.Format(QueryTemplate, TABLE_NAME, columnList, strtemp))
            End If
        Next

        'Logger.Log(enLogLevel.Success, sb.ToString())
        Return sb.ToString()
    End Function

End Class