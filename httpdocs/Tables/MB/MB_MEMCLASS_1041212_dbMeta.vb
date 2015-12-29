Public Class MB_MEMCLASS_1041212_dbMeta
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
			hMetaData3.add("COLUMN_NAME", "MB_CHKFLAG" )
			hMetaData3.add("DB_TYPE", 16 )
			hMetaData3.add("PROVIDER_TYPE",254 )
			hMetaData3.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData3)
			Static hMetaData4 As new System.Collections.Hashtable
			hMetaData4.add("COLUMN_NAME", "MB_NAPRSN" )
			hMetaData4.add("DB_TYPE", 16 )
			hMetaData4.add("PROVIDER_TYPE",253 )
			hMetaData4.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData4)
			Static hMetaData5 As new System.Collections.Hashtable
			hMetaData5.add("COLUMN_NAME", "MB_CREDATETIME" )
			hMetaData5.add("DB_TYPE", 6 )
			hMetaData5.add("PROVIDER_TYPE",12 )
			hMetaData5.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.DateTime )
			arry.add(hMetaData5)
			Static hMetaData6 As new System.Collections.Hashtable
			hMetaData6.add("COLUMN_NAME", "MB_APVDATETIME" )
			hMetaData6.add("DB_TYPE", 6 )
			hMetaData6.add("PROVIDER_TYPE",12 )
			hMetaData6.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.DateTime )
			arry.add(hMetaData6)
			Static hMetaData7 As new System.Collections.Hashtable
			hMetaData7.add("COLUMN_NAME", "MB_FWMK" )
			hMetaData7.add("DB_TYPE", 16 )
			hMetaData7.add("PROVIDER_TYPE",254 )
			hMetaData7.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData7)
			Static hMetaData8 As new System.Collections.Hashtable
			hMetaData8.add("COLUMN_NAME", "MB_MEMO" )
			hMetaData8.add("DB_TYPE", 16 )
			hMetaData8.add("PROVIDER_TYPE",253 )
			hMetaData8.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData8)
			Static hMetaData9 As new System.Collections.Hashtable
			hMetaData9.add("COLUMN_NAME", "MB_AMT" )
			hMetaData9.add("DB_TYPE", 7 )
			hMetaData9.add("PROVIDER_TYPE",246 )
			hMetaData9.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData9)
			Static hMetaData10 As new System.Collections.Hashtable
			hMetaData10.add("COLUMN_NAME", "MB_AMTMEMO" )
			hMetaData10.add("DB_TYPE", 16 )
			hMetaData10.add("PROVIDER_TYPE",253 )
			hMetaData10.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData10)
			Static hMetaData11 As new System.Collections.Hashtable
			hMetaData11.add("COLUMN_NAME", "MB_SORTNO" )
			hMetaData11.add("DB_TYPE", 7 )
			hMetaData11.add("PROVIDER_TYPE",246 )
			hMetaData11.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData11)
			Static hMetaData12 As new System.Collections.Hashtable
			hMetaData12.add("COLUMN_NAME", "MB_OBJECT" )
			hMetaData12.add("DB_TYPE", 16 )
			hMetaData12.add("PROVIDER_TYPE",253 )
			hMetaData12.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData12)
			Static hMetaData13 As new System.Collections.Hashtable
			hMetaData13.add("COLUMN_NAME", "MB_CDATE" )
			hMetaData13.add("DB_TYPE", 6 )
			hMetaData13.add("PROVIDER_TYPE",10 )
			hMetaData13.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Date )
			arry.add(hMetaData13)
			Static hMetaData14 As new System.Collections.Hashtable
			hMetaData14.add("COLUMN_NAME", "MB_BATCH" )
			hMetaData14.add("DB_TYPE", 7 )
			hMetaData14.add("PROVIDER_TYPE",246 )
			hMetaData14.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData14)
			Static hMetaData15 As new System.Collections.Hashtable
			hMetaData15.add("COLUMN_NAME", "MB_CHGDATE" )
			hMetaData15.add("DB_TYPE", 6 )
			hMetaData15.add("PROVIDER_TYPE",12 )
			hMetaData15.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.DateTime )
			arry.add(hMetaData15)
			Static hMetaData16 As new System.Collections.Hashtable
			hMetaData16.add("COLUMN_NAME", "MB_RESP" )
			hMetaData16.add("DB_TYPE", 16 )
			hMetaData16.add("PROVIDER_TYPE",254 )
			hMetaData16.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData16)
     End Sub
     Public Function getMetaArray() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getMetaArray
         Return arry
     End Function
     Public sub setPrimaryKeys() Implements com.Azion.NET.VB.DBMetaData.setPrimaryKeys
     End Sub
     Public Function getPrimaryKeys() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getPrimaryKeys
         Return m_arryPrimaryKeys
     End Function
End Class
