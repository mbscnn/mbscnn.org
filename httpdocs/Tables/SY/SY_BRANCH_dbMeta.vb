Public Class SY_BRANCH_dbMeta
        Implements com.Azion.NET.VB.DBMetaData
   Private  arry As New System.Collections.ArrayList
   Private  m_arryPrimaryKeys As New System.Collections.ArrayList
     Sub init() Implements com.Azion.NET.VB.DBMetaData.init
			Static hMetaData1 As new System.Collections.Hashtable
			hMetaData1.add("COLUMN_NAME", "BRA_DEPNO" )
			hMetaData1.add("DB_TYPE", 11 )
			hMetaData1.add("PROVIDER_TYPE",3 )
			hMetaData1.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Int32 )
			arry.add(hMetaData1)
			Static hMetaData2 As new System.Collections.Hashtable
			hMetaData2.add("COLUMN_NAME", "BRID" )
			hMetaData2.add("DB_TYPE", 16 )
			hMetaData2.add("PROVIDER_TYPE",253 )
			hMetaData2.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData2)
			Static hMetaData3 As new System.Collections.Hashtable
			hMetaData3.add("COLUMN_NAME", "PARENT" )
			hMetaData3.add("DB_TYPE", 11 )
			hMetaData3.add("PROVIDER_TYPE",3 )
			hMetaData3.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Int32 )
			arry.add(hMetaData3)
			Static hMetaData4 As new System.Collections.Hashtable
			hMetaData4.add("COLUMN_NAME", "BRCNAME" )
			hMetaData4.add("DB_TYPE", 16 )
			hMetaData4.add("PROVIDER_TYPE",253 )
			hMetaData4.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData4)
			Static hMetaData5 As new System.Collections.Hashtable
			hMetaData5.add("COLUMN_NAME", "BRCADDR" )
			hMetaData5.add("DB_TYPE", 16 )
			hMetaData5.add("PROVIDER_TYPE",253 )
			hMetaData5.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData5)
			Static hMetaData6 As new System.Collections.Hashtable
			hMetaData6.add("COLUMN_NAME", "BRATEL" )
			hMetaData6.add("DB_TYPE", 16 )
			hMetaData6.add("PROVIDER_TYPE",253 )
			hMetaData6.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData6)
			Static hMetaData7 As new System.Collections.Hashtable
			hMetaData7.add("COLUMN_NAME", "BRTEL" )
			hMetaData7.add("DB_TYPE", 16 )
			hMetaData7.add("PROVIDER_TYPE",253 )
			hMetaData7.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData7)
			Static hMetaData8 As new System.Collections.Hashtable
			hMetaData8.add("COLUMN_NAME", "BRFAX" )
			hMetaData8.add("DB_TYPE", 16 )
			hMetaData8.add("PROVIDER_TYPE",253 )
			hMetaData8.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData8)
			Static hMetaData9 As new System.Collections.Hashtable
			hMetaData9.add("COLUMN_NAME", "BRENAME" )
			hMetaData9.add("DB_TYPE", 16 )
			hMetaData9.add("PROVIDER_TYPE",253 )
			hMetaData9.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData9)
			Static hMetaData10 As new System.Collections.Hashtable
			hMetaData10.add("COLUMN_NAME", "BREADDR" )
			hMetaData10.add("DB_TYPE", 16 )
			hMetaData10.add("PROVIDER_TYPE",253 )
			hMetaData10.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData10)
			Static hMetaData11 As new System.Collections.Hashtable
			hMetaData11.add("COLUMN_NAME", "DISABLED" )
			hMetaData11.add("DB_TYPE", 16 )
			hMetaData11.add("PROVIDER_TYPE",254 )
			hMetaData11.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData11)
			Static hMetaData12 As new System.Collections.Hashtable
			hMetaData12.add("COLUMN_NAME", "BRCCITY" )
			hMetaData12.add("DB_TYPE", 16 )
			hMetaData12.add("PROVIDER_TYPE",253 )
			hMetaData12.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData12)
			Static hMetaData13 As new System.Collections.Hashtable
			hMetaData13.add("COLUMN_NAME", "BRCAREA" )
			hMetaData13.add("DB_TYPE", 16 )
			hMetaData13.add("PROVIDER_TYPE",253 )
			hMetaData13.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData13)
			Static hMetaData14 As new System.Collections.Hashtable
			hMetaData14.add("COLUMN_NAME", "HOFLAG" )
			hMetaData14.add("DB_TYPE", 16 )
			hMetaData14.add("PROVIDER_TYPE",254 )
			hMetaData14.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData14)
			Static hMetaData15 As new System.Collections.Hashtable
			hMetaData15.add("COLUMN_NAME", "EPSDEP" )
			hMetaData15.add("DB_TYPE", 16 )
			hMetaData15.add("PROVIDER_TYPE",253 )
			hMetaData15.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData15)
     End Sub
     Public Function getMetaArray() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getMetaArray
         Return arry
     End Function
     Public sub setPrimaryKeys() Implements com.Azion.NET.VB.DBMetaData.setPrimaryKeys
       m_arryPrimaryKeys.add("BRA_DEPNO")
     End Sub
     Public Function getPrimaryKeys() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getPrimaryKeys
         Return m_arryPrimaryKeys
     End Function
End Class
