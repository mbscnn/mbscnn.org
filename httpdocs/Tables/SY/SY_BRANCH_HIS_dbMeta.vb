Public Class SY_BRANCH_HIS_dbMeta
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
			hMetaData2.add("COLUMN_NAME", "CASEID" )
			hMetaData2.add("DB_TYPE", 16 )
			hMetaData2.add("PROVIDER_TYPE",254 )
			hMetaData2.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData2)
			Static hMetaData3 As new System.Collections.Hashtable
			hMetaData3.add("COLUMN_NAME", "STEP_NO" )
			hMetaData3.add("DB_TYPE", 16 )
			hMetaData3.add("PROVIDER_TYPE",253 )
			hMetaData3.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData3)
			Static hMetaData4 As new System.Collections.Hashtable
			hMetaData4.add("COLUMN_NAME", "SUBFLOW_SEQ" )
			hMetaData4.add("DB_TYPE", 11 )
			hMetaData4.add("PROVIDER_TYPE",3 )
			hMetaData4.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Int32 )
			arry.add(hMetaData4)
			Static hMetaData5 As new System.Collections.Hashtable
			hMetaData5.add("COLUMN_NAME", "SUBFLOW_COUNT" )
			hMetaData5.add("DB_TYPE", 11 )
			hMetaData5.add("PROVIDER_TYPE",3 )
			hMetaData5.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Int32 )
			arry.add(hMetaData5)
			Static hMetaData6 As new System.Collections.Hashtable
			hMetaData6.add("COLUMN_NAME", "BRID" )
			hMetaData6.add("DB_TYPE", 16 )
			hMetaData6.add("PROVIDER_TYPE",253 )
			hMetaData6.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData6)
			Static hMetaData7 As new System.Collections.Hashtable
			hMetaData7.add("COLUMN_NAME", "PARENT" )
			hMetaData7.add("DB_TYPE", 11 )
			hMetaData7.add("PROVIDER_TYPE",3 )
			hMetaData7.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Int32 )
			arry.add(hMetaData7)
			Static hMetaData8 As new System.Collections.Hashtable
			hMetaData8.add("COLUMN_NAME", "BRCNAME" )
			hMetaData8.add("DB_TYPE", 16 )
			hMetaData8.add("PROVIDER_TYPE",253 )
			hMetaData8.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData8)
			Static hMetaData9 As new System.Collections.Hashtable
			hMetaData9.add("COLUMN_NAME", "BRCADDR" )
			hMetaData9.add("DB_TYPE", 16 )
			hMetaData9.add("PROVIDER_TYPE",253 )
			hMetaData9.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData9)
			Static hMetaData10 As new System.Collections.Hashtable
			hMetaData10.add("COLUMN_NAME", "BRATEL" )
			hMetaData10.add("DB_TYPE", 16 )
			hMetaData10.add("PROVIDER_TYPE",253 )
			hMetaData10.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData10)
			Static hMetaData11 As new System.Collections.Hashtable
			hMetaData11.add("COLUMN_NAME", "BRTEL" )
			hMetaData11.add("DB_TYPE", 16 )
			hMetaData11.add("PROVIDER_TYPE",253 )
			hMetaData11.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData11)
			Static hMetaData12 As new System.Collections.Hashtable
			hMetaData12.add("COLUMN_NAME", "BRFAX" )
			hMetaData12.add("DB_TYPE", 16 )
			hMetaData12.add("PROVIDER_TYPE",253 )
			hMetaData12.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData12)
			Static hMetaData13 As new System.Collections.Hashtable
			hMetaData13.add("COLUMN_NAME", "BRENAME" )
			hMetaData13.add("DB_TYPE", 16 )
			hMetaData13.add("PROVIDER_TYPE",253 )
			hMetaData13.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData13)
			Static hMetaData14 As new System.Collections.Hashtable
			hMetaData14.add("COLUMN_NAME", "BREADDR" )
			hMetaData14.add("DB_TYPE", 16 )
			hMetaData14.add("PROVIDER_TYPE",253 )
			hMetaData14.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData14)
			Static hMetaData15 As new System.Collections.Hashtable
			hMetaData15.add("COLUMN_NAME", "DISABLED" )
			hMetaData15.add("DB_TYPE", 16 )
			hMetaData15.add("PROVIDER_TYPE",254 )
			hMetaData15.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData15)
			Static hMetaData16 As new System.Collections.Hashtable
			hMetaData16.add("COLUMN_NAME", "BRCCITY" )
			hMetaData16.add("DB_TYPE", 16 )
			hMetaData16.add("PROVIDER_TYPE",253 )
			hMetaData16.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData16)
			Static hMetaData17 As new System.Collections.Hashtable
			hMetaData17.add("COLUMN_NAME", "BRCAREA" )
			hMetaData17.add("DB_TYPE", 16 )
			hMetaData17.add("PROVIDER_TYPE",253 )
			hMetaData17.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData17)
			Static hMetaData18 As new System.Collections.Hashtable
			hMetaData18.add("COLUMN_NAME", "HOFLAG" )
			hMetaData18.add("DB_TYPE", 16 )
			hMetaData18.add("PROVIDER_TYPE",254 )
			hMetaData18.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData18)
			Static hMetaData19 As new System.Collections.Hashtable
			hMetaData19.add("COLUMN_NAME", "EPSDEP" )
			hMetaData19.add("DB_TYPE", 16 )
			hMetaData19.add("PROVIDER_TYPE",253 )
			hMetaData19.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData19)
			Static hMetaData20 As new System.Collections.Hashtable
			hMetaData20.add("COLUMN_NAME", "APPROVED" )
			hMetaData20.add("DB_TYPE", 16 )
			hMetaData20.add("PROVIDER_TYPE",254 )
			hMetaData20.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData20)
			Static hMetaData21 As new System.Collections.Hashtable
			hMetaData21.add("COLUMN_NAME", "XMLDATA" )
			hMetaData21.add("DB_TYPE", 16 )
			hMetaData21.add("PROVIDER_TYPE",752 )
			hMetaData21.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Text )
			arry.add(hMetaData21)
			Static hMetaData22 As new System.Collections.Hashtable
			hMetaData22.add("COLUMN_NAME", "OPERATION" )
			hMetaData22.add("DB_TYPE", 16 )
			hMetaData22.add("PROVIDER_TYPE",254 )
			hMetaData22.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData22)
			Static hMetaData23 As new System.Collections.Hashtable
			hMetaData23.add("COLUMN_NAME", "MGR_BRID" )
			hMetaData23.add("DB_TYPE", 16 )
			hMetaData23.add("PROVIDER_TYPE",253 )
			hMetaData23.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData23)
     End Sub
     Public Function getMetaArray() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getMetaArray
         Return arry
     End Function
     Public sub setPrimaryKeys() Implements com.Azion.NET.VB.DBMetaData.setPrimaryKeys
       m_arryPrimaryKeys.add("BRA_DEPNO")
       m_arryPrimaryKeys.add("CASEID")
       m_arryPrimaryKeys.add("STEP_NO")
       m_arryPrimaryKeys.add("SUBFLOW_SEQ")
       m_arryPrimaryKeys.add("SUBFLOW_COUNT")
     End Sub
     Public Function getPrimaryKeys() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getPrimaryKeys
         Return m_arryPrimaryKeys
     End Function
End Class
