Public Class SY_FUNCTION_CODE_HIS_dbMeta
        Implements com.Azion.NET.VB.DBMetaData
   Private  arry As New System.Collections.ArrayList
   Private  m_arryPrimaryKeys As New System.Collections.ArrayList
     Sub init() Implements com.Azion.NET.VB.DBMetaData.init
			Static hMetaData1 As new System.Collections.Hashtable
			hMetaData1.add("COLUMN_NAME", "FUNCCODE" )
			hMetaData1.add("DB_TYPE", 11 )
			hMetaData1.add("PROVIDER_TYPE",3 )
			hMetaData1.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Int32 )
			arry.add(hMetaData1)
			Static hMetaData2 As new System.Collections.Hashtable
			hMetaData2.add("COLUMN_NAME", "FUCNAME" )
			hMetaData2.add("DB_TYPE", 16 )
			hMetaData2.add("PROVIDER_TYPE",253 )
			hMetaData2.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData2)
			Static hMetaData3 As new System.Collections.Hashtable
			hMetaData3.add("COLUMN_NAME", "FUCURL" )
			hMetaData3.add("DB_TYPE", 16 )
			hMetaData3.add("PROVIDER_TYPE",253 )
			hMetaData3.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData3)
			Static hMetaData4 As new System.Collections.Hashtable
			hMetaData4.add("COLUMN_NAME", "PARENT" )
			hMetaData4.add("DB_TYPE", 11 )
			hMetaData4.add("PROVIDER_TYPE",3 )
			hMetaData4.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Int32 )
			arry.add(hMetaData4)
			Static hMetaData5 As new System.Collections.Hashtable
			hMetaData5.add("COLUMN_NAME", "DISABLED" )
			hMetaData5.add("DB_TYPE", 16 )
			hMetaData5.add("PROVIDER_TYPE",254 )
			hMetaData5.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData5)
			Static hMetaData6 As new System.Collections.Hashtable
			hMetaData6.add("COLUMN_NAME", "SORTCTRL" )
			hMetaData6.add("DB_TYPE", 11 )
			hMetaData6.add("PROVIDER_TYPE",3 )
			hMetaData6.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Int32 )
			arry.add(hMetaData6)
			Static hMetaData7 As new System.Collections.Hashtable
			hMetaData7.add("COLUMN_NAME", "HOFLAG" )
			hMetaData7.add("DB_TYPE", 16 )
			hMetaData7.add("PROVIDER_TYPE",254 )
			hMetaData7.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData7)
			Static hMetaData8 As new System.Collections.Hashtable
			hMetaData8.add("COLUMN_NAME", "CASEID" )
			hMetaData8.add("DB_TYPE", 16 )
			hMetaData8.add("PROVIDER_TYPE",254 )
			hMetaData8.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData8)
			Static hMetaData9 As new System.Collections.Hashtable
			hMetaData9.add("COLUMN_NAME", "STEP_NO" )
			hMetaData9.add("DB_TYPE", 16 )
			hMetaData9.add("PROVIDER_TYPE",253 )
			hMetaData9.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData9)
			Static hMetaData10 As new System.Collections.Hashtable
			hMetaData10.add("COLUMN_NAME", "SUBFLOW_SEQ" )
			hMetaData10.add("DB_TYPE", 11 )
			hMetaData10.add("PROVIDER_TYPE",3 )
			hMetaData10.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Int32 )
			arry.add(hMetaData10)
			Static hMetaData11 As new System.Collections.Hashtable
			hMetaData11.add("COLUMN_NAME", "SUBFLOW_COUNT" )
			hMetaData11.add("DB_TYPE", 11 )
			hMetaData11.add("PROVIDER_TYPE",3 )
			hMetaData11.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Int32 )
			arry.add(hMetaData11)
			Static hMetaData12 As new System.Collections.Hashtable
			hMetaData12.add("COLUMN_NAME", "APPROVED" )
			hMetaData12.add("DB_TYPE", 16 )
			hMetaData12.add("PROVIDER_TYPE",254 )
			hMetaData12.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData12)
			Static hMetaData13 As new System.Collections.Hashtable
			hMetaData13.add("COLUMN_NAME", "XMLDATA" )
			hMetaData13.add("DB_TYPE", 16 )
			hMetaData13.add("PROVIDER_TYPE",752 )
			hMetaData13.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Text )
			arry.add(hMetaData13)
			Static hMetaData14 As new System.Collections.Hashtable
			hMetaData14.add("COLUMN_NAME", "OPERATION" )
			hMetaData14.add("DB_TYPE", 16 )
			hMetaData14.add("PROVIDER_TYPE",254 )
			hMetaData14.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData14)
     End Sub
     Public Function getMetaArray() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getMetaArray
         Return arry
     End Function
     Public sub setPrimaryKeys() Implements com.Azion.NET.VB.DBMetaData.setPrimaryKeys
       m_arryPrimaryKeys.add("FUNCCODE")
       m_arryPrimaryKeys.add("CASEID")
       m_arryPrimaryKeys.add("STEP_NO")
       m_arryPrimaryKeys.add("SUBFLOW_SEQ")
       m_arryPrimaryKeys.add("SUBFLOW_COUNT")
     End Sub
     Public Function getPrimaryKeys() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getPrimaryKeys
         Return m_arryPrimaryKeys
     End Function
End Class
