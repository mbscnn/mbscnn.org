Imports System.IO
Imports Com.Azion.NET.VB
Imports System.Collections
Imports System.Data

Public Class generateMetaData

    'Shared m_sTableNameSQL As String

    Dim m_aryTableNames As System.Collections.ArrayList = New System.Collections.ArrayList
    Dim m_aryTableMetaData As System.Collections.ArrayList = New System.Collections.ArrayList

    Shared m_dbManager As DatabaseManager

    Dim m_sb As System.Text.StringBuilder

    Shared sCurDir As String = Microsoft.VisualBasic.FileSystem.CurDir
    Shared ipos As Integer = sCurDir.LastIndexOf("\")
    Shared m_fileName As String

    Shared Sub main()
        Properties.bConsoleMode = True
        Properties.setConfiguration("C:\inetpub\vhosts\mbscnn.org\httpdocs\MBSCConf\MBSC.config")

        Dim sTableNameSQL As String = ""
        Dim sOenTableNameSQL As String = ""

        '產生Table Meta or  View Meta
        Dim bTable As Boolean = True '產生Table Meta
        Dim bView As Boolean = True '產生View Meta
        ' Q.NET.VB.Properties.setConfiguration()
        If bTable Then
            'Table
            'sTableNameSQL = "select table_name from all_tables where all_tables.owner='" & Properties.getSchemaName.ToUpper & "' "

            ''跑單筆自己改
            'sOenTableNameSQL = " and  table_name like 'AP_BATCH_001%'"
            'for ibm db2
            'sTableNameSQL = "Select tabname from SYSCATV82.SNAPTAB  where tabschema='SFAUSR'  and  tabname = 'SFA_COP_CIF_CUST' or tabname = 'SFA_COP_CIF_ADDR' or tabname = 'SFA_COP_CIF_PHONE'"
            ' sTableNameSQL = "Select tabname from SYSSTAT.TABLES  where tabname = 'CMSR05'"

            'sTableNameSQL &= sOenTableNameSQL

            sTableNameSQL = "SELECT UPPER(table_name) FROM INFORMATION_SCHEMA.TABLES where table_schema='" & Properties.getSchemaName.ToUpper & "' "
            Try
                m_dbManager = DatabaseManager.getInstance()
                'm_dbManager = DatabaseManager.getInstance("Provider=IBMDADB2;Data Source=DBPOC;Database=DBPOC;User ID=db2admin;Password=titanqq")
                'm_dbManager = DatabaseManager.getInstance("Provider=IBMDADB2;Data Source=ACLDB;Database=ACLDB;Hostname=10.8.208.58;Protocol=TCPIP;Port=50000;User ID=aclpoc;Password=aclpoc")
                Dim meta As generateMetaData = New generateMetaData
                meta.getTableNames(sTableNameSQL)
                meta.geneCode()
            Catch ex As Exception
                Throw ex
            Finally
                m_dbManager.releaseConnection()
            End Try

        End If


        If bView Then
            'VIEW
            'select view_name as table_name from all_views  where all_views.owner='BOTELOAN' and all_views.view_name='APVW_BRNLBDIST'
            sTableNameSQL = "SELECT UPPER(table_name) FROM INFORMATION_SCHEMA.VIEWS where table_schema='" & Properties.getSchemaName.ToUpper & "' "
            'sTableNameSQL = "select view_name as table_name from all_views  where all_views.owner= '" & Properties.getSchemaName.ToUpper & "' "
            '跑單筆自己改
            'sOenTableNameSQL = " and all_views.view_name='VIEW_JCIC_91_74_FINAL%'"
            'sTableNameSQL &= sOenTableNameSQL

            Try
                m_dbManager = DatabaseManager.getInstance()
                Dim meta As generateMetaData = New generateMetaData
                meta.getViewTableNames(sTableNameSQL)
                meta.geneCode(True)
            Catch ex As Exception

            Finally
                m_dbManager.releaseConnection()
            End Try
        End If
    End Sub

    Private Sub getTableNames(ByVal sTableNameSQL As String)
        Dim objDs As DataSet = DBObject.ExecuteDataset(m_dbManager.getConnection, CommandType.Text, sTableNameSQL, Nothing)
        For iCount As Integer = 1 To objDs.Tables(0).DefaultView.Count

            For Each col As DataColumn In objDs.Tables(0).Columns
                Dim sColumnName As String = CType(objDs.Tables(0).DefaultView.Item(iCount - 1).Item(col.ColumnName), String)
                m_aryTableNames.Add(sColumnName)

                'Console.WriteLine("[" & sColumnName & "]=" & sColumnName)
            Next
        Next
    End Sub

    Private Sub getViewTableNames(ByVal sTableNameSQL As String)
        Dim objDs As DataSet = DBObject.ExecuteDataset(m_dbManager.getConnection, CommandType.Text, sTableNameSQL, Nothing)
        For iCount As Integer = 1 To objDs.Tables(0).DefaultView.Count

            For Each col As DataColumn In objDs.Tables(0).Columns
                Dim sColumnName As String = CType(objDs.Tables(0).DefaultView.Item(iCount - 1).Item(col.ColumnName), String)
                m_aryTableNames.Add(sColumnName)
                'Console.WriteLine("[" & sColumnName & "]=" & sColumnName)
            Next
        Next
    End Sub

    Private Function getColumnName(ByVal sTableName As String) As System.Collections.ArrayList

        Dim schemaTable As DataTable
        'For Each sTableName As String In aryTableNames
        schemaTable = DbUtility.getTable(sTableName, Me.m_dbManager)
        'Console.WriteLine("==============" & sTableName & "================")
        Dim hData As System.Collections.Hashtable
        m_aryTableMetaData = New System.Collections.ArrayList

        For i As Integer = 0 To schemaTable.Rows.Count - 1
            hData = New System.Collections.Hashtable
            ' attribute meta data (column name, type name)
            constructAttrMeta(schemaTable.Rows(i), hData)
            m_aryTableMetaData.Add(hData)
        Next i

        'Console.WriteLine("====================================")
        ' Next
        Return m_aryTableMetaData
    End Function


    Sub constructAttrMeta(ByVal objRows As DataRow, ByVal hData As System.Collections.Hashtable)
        Try
            hData.Add("COLUMN_NAME", objRows!ColumnName.ToString) 'COLUMN_NAME
            hData.Add("DB_TYPE", DbUtility.getFieldType(objRows!DataType.ToString)) 'DB_TYPE
            hData.Add("PROVIDER_TYPE", objRows!ProviderType) 'PROVIDER_TYPE
            hData.Add("PROVIDER_TYPE_NAME", ProviderFactory.getFieldType(objRows!ProviderType))  'PROVIDER_TYPE_NAME
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Sub geneCode(Optional ByVal bView As Boolean = False)
        Dim iTable As Integer = 1
        For Each sTableName As String In m_aryTableNames
            Console.WriteLine(iTable & "." & "[" & sTableName & "]")
            iTable += 1
            m_sb = New System.Text.StringBuilder
            'm_sb.Append("'########################################" & Microsoft.VisualBasic.vbCrLf)
            'm_sb.Append("'#                                      #" & Microsoft.VisualBasic.vbCrLf)
            'm_sb.Append("'# Cteate Time :" & Now.ToString & "  #" & Microsoft.VisualBasic.vbCrLf)
            'm_sb.Append("'#                                      #" & Microsoft.VisualBasic.vbCrLf)
            'm_sb.Append("'########################################" & Microsoft.VisualBasic.vbCrLf & Microsoft.VisualBasic.vbCrLf)
            m_sb.Append("Public Class " & sTableName & "_dbMeta" & Microsoft.VisualBasic.vbCrLf)
            m_sb.Append("        Implements com.Azion.NET.VB.DBMetaData" & Microsoft.VisualBasic.vbCrLf)
            m_sb.Append("   Private  arry As New System.Collections.ArrayList" & Microsoft.VisualBasic.vbCrLf)
            m_sb.Append("   Private  m_arryPrimaryKeys As New System.Collections.ArrayList" & Microsoft.VisualBasic.vbCrLf)

            m_sb.Append("     Sub init() Implements com.Azion.NET.VB.DBMetaData.init" & Microsoft.VisualBasic.vbCrLf)
            Try
                Dim arry As System.Collections.ArrayList = getColumnName(sTableName)
                For i As Integer = 0 To arry.Count - 1
                    Dim hData As System.Collections.Hashtable = arry.Item(i)
                    Dim COLUMN_NAME As String = hData.Item("COLUMN_NAME") 'COLUMN_NAME 欄位名稱
                    Dim DATA_TYPE As String = hData.Item("DB_TYPE") 'DB_TYPE 欄位型態代碼
                    Dim PROVIDER_TYPE As String = hData.Item("PROVIDER_TYPE") 'PROVIDER_TYPE 欄位型態
                    Dim PROVIDER_TYPE_NAME As String = hData.Item("PROVIDER_TYPE_NAME")  'PROVIDER_TYPE_NAME

                    Dim j As Integer = i + 1
                    m_sb.Append("			Static hMetaData" & j & " As new System.Collections.Hashtable" & Microsoft.VisualBasic.vbCrLf)
                    m_sb.Append("			hMetaData" & j & ".add(""COLUMN_NAME"", """ & COLUMN_NAME & """ )" & Microsoft.VisualBasic.vbCrLf)
                    m_sb.Append("			hMetaData" & j & ".add(""DB_TYPE"", " & DATA_TYPE & " )" & Microsoft.VisualBasic.vbCrLf)
                    m_sb.Append("			hMetaData" & j & ".add(""PROVIDER_TYPE""," & PROVIDER_TYPE & " )" & Microsoft.VisualBasic.vbCrLf)
                    m_sb.Append("			hMetaData" & j & ".add(""PROVIDER_TYPE_NAME"", " & PROVIDER_TYPE_NAME & " )" & Microsoft.VisualBasic.vbCrLf)
                    m_sb.Append("			arry.add(hMetaData" & j & ")" & Microsoft.VisualBasic.vbCrLf)
                Next
                m_sb.Append("     End Sub" & Microsoft.VisualBasic.vbCrLf)

                m_sb.Append("     Public Function getMetaArray() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getMetaArray" & Microsoft.VisualBasic.vbCrLf)
                m_sb.Append("         Return arry" & Microsoft.VisualBasic.vbCrLf)
                m_sb.Append("     End Function" & Microsoft.VisualBasic.vbCrLf)

                m_sb.Append("     Public sub setPrimaryKeys() Implements com.Azion.NET.VB.DBMetaData.setPrimaryKeys" & Microsoft.VisualBasic.vbCrLf)

                If Not bView Then
                    arry = DbUtility.getPrimaryKey(sTableName, Me.m_dbManager)
                    For i As Integer = 0 To arry.Count - 1

                        m_sb.Append("       m_arryPrimaryKeys.add(""" & arry.Item(i) & """)" & Microsoft.VisualBasic.vbCrLf)

                    Next
                End If
                m_sb.Append("     End Sub" & Microsoft.VisualBasic.vbCrLf)

                m_sb.Append("     Public Function getPrimaryKeys() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getPrimaryKeys" & Microsoft.VisualBasic.vbCrLf)
                m_sb.Append("         Return m_arryPrimaryKeys" & Microsoft.VisualBasic.vbCrLf)
                m_sb.Append("     End Function" & Microsoft.VisualBasic.vbCrLf)

                m_sb.Append("End Class" & Microsoft.VisualBasic.vbCrLf)
                ' Console.WriteLine(m_sb.ToString)

                write(sTableName, m_sb.ToString, bView)
            Catch ex As Exception

            End Try
        Next
    End Sub

    Shared Sub write(ByVal sClassName As String, ByVal sData As String, Optional ByVal bView As Boolean = False)
        Try
            If bView Then
                m_fileName = sCurDir.Substring(0, ipos + 1) & "VIEW" & "\"
            ElseIf sClassName.ToUpper.IndexOf("JC") <> -1 Then
                m_fileName = sCurDir.Substring(0, ipos + 1) & "Tables\JC"
            ElseIf sClassName.ToUpper.IndexOf("FLOW") <> -1 Then
                m_fileName = sCurDir.Substring(0, ipos + 1) & "Tables\FLOW"
            ElseIf sClassName.ToUpper.IndexOf("_") <> -1 Then
                m_fileName = sCurDir.Substring(0, ipos + 1) & "Tables\" & sClassName.Split("_")(0).ToUpper
            Else
                m_fileName = sCurDir.Substring(0, ipos + 1) & "Tables\OTHER"
            End If

            If Not (Directory.Exists(m_fileName)) Then
                Directory.CreateDirectory(m_fileName)
            End If
            Dim fileName As String = m_fileName & "\" & sClassName & "_dbMeta.vb"
            ' Dim s As System.IO.TextWriter = New System.IO.TextWriter
            Dim sw As TextWriter = New StreamWriter(fileName)
            'Console.WriteLine(sData)
            sw.Write(sData)
            sw.Close()
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try


    End Sub

End Class
