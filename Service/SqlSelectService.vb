
Imports System.Data.SqlClient



Public Class SqlSelectService

    Public Property contentVList As List(Of String)
    Public Property headVList As List(Of String)

    Public Property Result As Boolean = False

    Public Sub New(job As xmlJob)

        Dim connectionString As String = "Data Source=" & job._JOB_SQL_SERVER_NAME & ";Initial Catalog=" & job._JOB_SQL_INITIAL_DB & ";User ID=" & job._JOB_SQL_USER_ID & ";Password=" & job._JOB_SQL_PASSWORD & ";Connection Timeout=10"

        Using connection As New SqlConnection(connectionString)

            contentVList = New List(Of String)
            headVList = New List(Of String)

            Dim command As SqlCommand = Nothing
            Dim reader As SqlDataReader = Nothing

            Try
                command = New SqlCommand(job._JOB_SQL_QUERY, connection)
                connection.Open()
                reader = command.ExecuteReader()
            Catch e As Exception
                'Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE} SQLSERVER 연결 실패")
                Logger.Log(enLogLevel.Fail, $"{job._JOB_TITLE} InsertDB {e.ToString}")
                If Not IsNothing(reader) Then
                    reader.Close()
                End If
                If Not IsNothing(connection) Then
                    connection.Close()
                End If
                Return
            End Try

            Dim fieldCnt As Integer = reader.FieldCount

            Logger.Log(enLogLevel.Info, $"SQLSERVER 연결 성공 : {job._JOB_TITLE} 쿼리 컬럼 갯수 : {fieldCnt}")

            headVList = getFieldNames(reader, fieldCnt)

            ' Call Read before accessing data.
            While reader.Read()

                Dim record As IDataRecord = CType(reader, IDataRecord)

                Dim strV As String = ""

                For index As Integer = 0 To fieldCnt - 1
                    If index = 0 Then
                        strV += String.Format("{0}", record(index))
                    Else
                        strV += String.Format(",{0}", record(index))
                    End If
                Next
                contentVList.Add(strV)
            End While

            reader.Close()

            connection.Close()

            Logger.Log(enLogLevel.Success, $"{job._JOB_TITLE} SQL 조회 성공 : {contentVList.Count}")

            Result = True
        End Using
    End Sub

    'reader와 count를 입력받아서 column field List 반환
    Private Function getFieldNames(ByVal reader As SqlDataReader, ByVal fieldCnt As Int16) As List(Of String)
        Dim tempList As List(Of String) = New List(Of String)

        If fieldCnt <= 0 Then
            Return tempList
        Else
            For index As Integer = 0 To fieldCnt - 1
                tempList.Add(reader.GetName(index))
            Next
        End If

        Return tempList
    End Function


End Class

