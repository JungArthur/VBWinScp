Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization

<XmlRoot(ElementName:="ROOT")>
Public Class xmlJobs

    <XmlArray(ElementName:="JOBS")>
    Public Property Jobs As List(Of xmlJob)

End Class

<XmlType(TypeName:="JOB")>
Public Class xmlJob

    <XmlElement(ElementName:="TITLE")>
    Public Property _JOB_TITLE As String

    <XmlElement(ElementName:="FILE_PREFIX")>
    Public Property _JOB_FILE_PREFIX As String
    <XmlElement(ElementName:="FILE_DATE_FORMAT")>
    Public Property _JOB_FILE_DATE_FORMAT As String


    <XmlElement(ElementName:="SQL_SERVER_NAME")>
    Public Property _JOB_SQL_SERVER_NAME As String
    <XmlElement(ElementName:="SQL_INITIAL_DB")>
    Public Property _JOB_SQL_INITIAL_DB As String
    <XmlElement(ElementName:="SQL_USER_ID")>
    Public Property _JOB_SQL_USER_ID As String
    <XmlElement(ElementName:="SQL_PASSWORD")>
    Public Property _JOB_SQL_PASSWORD As String
    <XmlElement(ElementName:="SQL_QUERY")>
    Public Property _JOB_SQL_QUERY As String


    <XmlElement(ElementName:="SFTP_METHOD")>
    Public Property _JOB_SFTP_METHOD As String
    <XmlElement(ElementName:="SFTP_SERVER_NAME")>
    Public Property _JOB_SFTP_SERVER_NAME As String
    <XmlElement(ElementName:="SFTP_SERVER_PORT")>
    Public Property _JOB_SFTP_SERVER_PORT As String
    <XmlElement(ElementName:="SFTP_USER_ID")>
    Public Property _JOB_SFTP_USER_ID As String
    <XmlElement(ElementName:="SFTP_PASSWORD")>
    Public Property _JOB_SFTP_PASSWORD As String
    <XmlElement(ElementName:="SFTP_REMOTE_DIRECTORY")>
    Public Property _JOB_SFTP_REMOTE_DIRECTORY As String


    <XmlElement(ElementName:="INSERT_SERVER_NAME")>
    Public Property _JOB_INSERT_SERVER_NAME As String
    <XmlElement(ElementName:="INSERT_INITIAL_DB")>
    Public Property _JOB_INSERT_INITIAL_DB As String
    <XmlElement(ElementName:="INSERT_USER_ID")>
    Public Property _JOB_INSERT_USER_ID As String
    <XmlElement(ElementName:="INSERT_PASSWORD")>
    Public Property _JOB_INSERT_PASSWORD As String
    <XmlElement(ElementName:="INSERT_TABLE_NAME")>
    Public Property _JOB_INSERT_TABLE_NAME As String
    <XmlElement(ElementName:="INSERT_COLUMN_NAMES")>
    Public Property _JOB_INSERT_COLUMN_NAMES As String
    <XmlElement(ElementName:="INSERT_CDT_ENABLED")>
    Public Property _JOB_INSERT_CDT_ENABLED As String

End Class




Public Class CXMLControl

    ' Dictionary 및 XML의 항목을 정의 ( static(정적) 변수로 사용 : 프로그램 실행 시 메모리에 바로 할당

    ' Job Setting
    Public _JOB_TITLE As String = "TITLE"
    Public _JOB_FILE_PREFIX As String = "FILE_PREFIX"
    Public _JOB_FILE_DATE_FORMAT As String = "FILE_DATE_FORMAT"

    ' SQL Setting
    Public _JOB_SQL_SERVER_NAME As String = "SQL_SERVER_NAME"
    Public _JOB_SQL_INITIAL_DB As String = "SQL_INITIAL_DB"
    Public _JOB_SQL_USER_ID As String = "SQL_USER_ID"
    Public _JOB_SQL_PASSWORD As String = "SQL_PASSWORD"
    Public _JOB_SQL_QUERY As String = "SQL_QUERY"

    ' SFTP Setting
    Public _JOB_SFTP_METHOD As String = "SFTP_METHOD"
    Public _JOB_SFTP_SERVER_NAME As String = "SFTP_SERVER_NAME"
    Public _JOB_SFTP_SERVER_PORT As String = "SFTP_SERVER_PORT"
    Public _JOB_SFTP_USER_ID As String = "SFTP_USER_ID"
    Public _JOB_SFTP_PASSWORD As String = "SFTP_PASSWORD"
    Public _JOB_SFTP_REMOTE_DIRECTORY As String = "SFTP_REMOTE_DIRECTORY"

    ' INSERT Setting
    Public _JOB_INSERT_SERVER_NAME As String = "INSERT_SERVER_NAME"
    Public _JOB_INSERT_INITIAL_DB As String = "INSERT_INITIAL_DB"
    Public _JOB_INSERT_USER_ID As String = "INSERT_USER_ID"
    Public _JOB_INSERT_PASSWORD As String = "INSERT_PASSWORD"
    Public _JOB_INSERT_TABLE_NAME As String = "INSERT_TABLE_NAME"
    Public _JOB_INSERT_COLUMN_NAMES As String = "INSERT_COLUMN_NAMES"
    Public _JOB_INSERT_CDT_ENABLED As String = "INSERT_CDT_ENABLED"




    '' <summary>
    '' XML을 저장 하기 위해 사용
    '' </summary>
    '' <param name="strXMLPath">저장 할 XML File의 경로 및 파일명</param>
    '' <param name="DXMLConfig">XML로 저장 할 항목</param>
    Public Sub fXML_Writer(strXMLPath As String, DXMLList As List(Of xmlJob))

        ' using 범위 내에 XmlWriter를 정의 하여 using을 벗어 나게 될 경우 자동으로 Dispose 하여 메모리를 관리
        Using wr As XmlWriter = XmlWriter.Create(strXMLPath)
            ' XML 형태의 Data를 작성 (결과 및 예제 확인)
            wr.WriteStartDocument()

            ' Setting#01
            wr.WriteStartElement("ROOT")
            wr.WriteStartElement("JOBS")
            ' wr.WriteAttributeString("ID", "0001")  ' attribute 쓰기

            For Each DXMLConfig As xmlJob In DXMLList
                wr.WriteStartElement("JOB")
                wr.WriteElementString(_JOB_TITLE, DXMLConfig._JOB_TITLE)
                wr.WriteElementString(_JOB_FILE_PREFIX, DXMLConfig._JOB_FILE_PREFIX)
                wr.WriteElementString(_JOB_FILE_DATE_FORMAT, DXMLConfig._JOB_FILE_DATE_FORMAT)

                wr.WriteElementString(_JOB_SQL_SERVER_NAME, DXMLConfig._JOB_SQL_SERVER_NAME)
                wr.WriteElementString(_JOB_SQL_INITIAL_DB, DXMLConfig._JOB_SQL_INITIAL_DB)
                wr.WriteElementString(_JOB_SQL_USER_ID, DXMLConfig._JOB_SQL_USER_ID)
                wr.WriteElementString(_JOB_SQL_PASSWORD, DXMLConfig._JOB_SQL_PASSWORD)
                wr.WriteElementString(_JOB_SQL_QUERY, DXMLConfig._JOB_SQL_QUERY)

                wr.WriteElementString(_JOB_SFTP_METHOD, DXMLConfig._JOB_SFTP_METHOD)
                wr.WriteElementString(_JOB_SFTP_SERVER_NAME, DXMLConfig._JOB_SFTP_SERVER_NAME)
                wr.WriteElementString(_JOB_SFTP_SERVER_PORT, DXMLConfig._JOB_SFTP_SERVER_PORT)
                wr.WriteElementString(_JOB_SFTP_USER_ID, DXMLConfig._JOB_SFTP_USER_ID)
                wr.WriteElementString(_JOB_SFTP_PASSWORD, DXMLConfig._JOB_SFTP_PASSWORD)
                wr.WriteElementString(_JOB_SFTP_REMOTE_DIRECTORY, DXMLConfig._JOB_SFTP_REMOTE_DIRECTORY)

                wr.WriteElementString(_JOB_INSERT_SERVER_NAME, DXMLConfig._JOB_INSERT_SERVER_NAME)
                wr.WriteElementString(_JOB_INSERT_INITIAL_DB, DXMLConfig._JOB_INSERT_INITIAL_DB)
                wr.WriteElementString(_JOB_INSERT_USER_ID, DXMLConfig._JOB_INSERT_USER_ID)
                wr.WriteElementString(_JOB_INSERT_PASSWORD, DXMLConfig._JOB_INSERT_PASSWORD)
                wr.WriteElementString(_JOB_INSERT_TABLE_NAME, DXMLConfig._JOB_INSERT_TABLE_NAME)
                wr.WriteElementString(_JOB_INSERT_COLUMN_NAMES, DXMLConfig._JOB_INSERT_COLUMN_NAMES)
                wr.WriteElementString(_JOB_INSERT_CDT_ENABLED, DXMLConfig._JOB_INSERT_CDT_ENABLED)

                wr.WriteEndElement()
            Next
            wr.WriteEndElement()
            wr.WriteEndElement()
            wr.WriteEndDocument()
            wr.Close()
        End Using
    End Sub

    '' <summary>
    '' XML을 읽어 오기 위해 사용
    '' </summary>
    '' <param name="strXMLPath">읽어 올 XML File의 경로 및 파일명</param>
    '' <returns></returns>
    Public Function fXML_Reader(strXMLPath As String) As List(Of xmlJob)

        Dim deserializer As New XmlSerializer(GetType(xmlJobs))
        Dim search As xmlJobs = Nothing

        Using reader As New StreamReader(strXMLPath)
            search = DirectCast(deserializer.Deserialize(reader), xmlJobs)
            reader.Close()
        End Using

        Return search.Jobs   ' 작성 한 Dictionary를 반환
    End Function

End Class
