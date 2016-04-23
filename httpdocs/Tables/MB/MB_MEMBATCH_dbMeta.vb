Public Class MB_MEMBATCH_dbMeta
        Implements com.Azion.NET.VB.DBMetaData
   Private  arry As New System.Collections.ArrayList
   Private  m_arryPrimaryKeys As New System.Collections.ArrayList
     Sub init() Implements com.Azion.NET.VB.DBMetaData.init
			Static hMetaData1 As new System.Collections.Hashtable
			hMetaData1.add("COLUMN_NAME", "MB_MEMSEQ" )
			hMetaData1.add("DB_TYPE", 7 )
			hMetaData1.add("PROVIDER_TYPE",246 )
			hMetaData1.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData1)
			Static hMetaData2 As new System.Collections.Hashtable
			hMetaData2.add("COLUMN_NAME", "MB_SEQ" )
			hMetaData2.add("DB_TYPE", 7 )
			hMetaData2.add("PROVIDER_TYPE",246 )
			hMetaData2.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData2)
			Static hMetaData3 As new System.Collections.Hashtable
			hMetaData3.add("COLUMN_NAME", "MB_BATCH" )
			hMetaData3.add("DB_TYPE", 7 )
			hMetaData3.add("PROVIDER_TYPE",246 )
			hMetaData3.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData3)
			Static hMetaData4 As new System.Collections.Hashtable
			hMetaData4.add("COLUMN_NAME", "MB_ELECT" )
			hMetaData4.add("DB_TYPE", 16 )
			hMetaData4.add("PROVIDER_TYPE",254 )
			hMetaData4.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData4)
			Static hMetaData5 As new System.Collections.Hashtable
			hMetaData5.add("COLUMN_NAME", "MB_CHKFLAG" )
			hMetaData5.add("DB_TYPE", 16 )
			hMetaData5.add("PROVIDER_TYPE",254 )
			hMetaData5.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData5)
			Static hMetaData6 As new System.Collections.Hashtable
			hMetaData6.add("COLUMN_NAME", "MB_RESP" )
			hMetaData6.add("DB_TYPE", 16 )
			hMetaData6.add("PROVIDER_TYPE",254 )
			hMetaData6.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData6)
			Static hMetaData7 As new System.Collections.Hashtable
			hMetaData7.add("COLUMN_NAME", "MB_APVDATE" )
			hMetaData7.add("DB_TYPE", 6 )
			hMetaData7.add("PROVIDER_TYPE",10 )
			hMetaData7.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Date )
			arry.add(hMetaData7)
     End Sub
     Public Function getMetaArray() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getMetaArray
         Return arry
     End Function
     Public sub setPrimaryKeys() Implements com.Azion.NET.VB.DBMetaData.setPrimaryKeys
       m_arryPrimaryKeys.add("MB_MEMSEQ")
       m_arryPrimaryKeys.add("MB_SEQ")
       m_arryPrimaryKeys.add("MB_BATCH")
     End Sub
     Public Function getPrimaryKeys() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getPrimaryKeys
         Return m_arryPrimaryKeys
     End Function
End Class
