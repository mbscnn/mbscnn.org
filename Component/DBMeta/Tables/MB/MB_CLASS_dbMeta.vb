Public Class MB_CLASS_dbMeta
        Implements com.Azion.NET.VB.DBMetaData
   Private  arry As New System.Collections.ArrayList
   Private  m_arryPrimaryKeys As New System.Collections.ArrayList
     Sub init() Implements com.Azion.NET.VB.DBMetaData.init
			Static hMetaData1 As new System.Collections.Hashtable
			hMetaData1.add("COLUMN_NAME", "MB_SEQ" )
			hMetaData1.add("DB_TYPE", 7 )
			hMetaData1.add("PROVIDER_TYPE",246 )
			hMetaData1.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData1)
			Static hMetaData2 As new System.Collections.Hashtable
			hMetaData2.add("COLUMN_NAME", "MB_BATCH" )
			hMetaData2.add("DB_TYPE", 7 )
			hMetaData2.add("PROVIDER_TYPE",246 )
			hMetaData2.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData2)
			Static hMetaData3 As new System.Collections.Hashtable
			hMetaData3.add("COLUMN_NAME", "MB_PLACE" )
			hMetaData3.add("DB_TYPE", 16 )
			hMetaData3.add("PROVIDER_TYPE",253 )
			hMetaData3.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData3)
			Static hMetaData4 As new System.Collections.Hashtable
			hMetaData4.add("COLUMN_NAME", "MB_SDATE" )
			hMetaData4.add("DB_TYPE", 6 )
			hMetaData4.add("PROVIDER_TYPE",10 )
			hMetaData4.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Date )
			arry.add(hMetaData4)
			Static hMetaData5 As new System.Collections.Hashtable
			hMetaData5.add("COLUMN_NAME", "MB_EDATE" )
			hMetaData5.add("DB_TYPE", 6 )
			hMetaData5.add("PROVIDER_TYPE",10 )
			hMetaData5.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Date )
			arry.add(hMetaData5)
			Static hMetaData6 As new System.Collections.Hashtable
			hMetaData6.add("COLUMN_NAME", "MB_SWEEK" )
			hMetaData6.add("DB_TYPE", 7 )
			hMetaData6.add("PROVIDER_TYPE",246 )
			hMetaData6.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData6)
			Static hMetaData7 As new System.Collections.Hashtable
			hMetaData7.add("COLUMN_NAME", "MB_EWEEK" )
			hMetaData7.add("DB_TYPE", 7 )
			hMetaData7.add("PROVIDER_TYPE",246 )
			hMetaData7.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData7)
			Static hMetaData8 As new System.Collections.Hashtable
			hMetaData8.add("COLUMN_NAME", "MB_CLASS_NAME" )
			hMetaData8.add("DB_TYPE", 16 )
			hMetaData8.add("PROVIDER_TYPE",253 )
			hMetaData8.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData8)
			Static hMetaData9 As new System.Collections.Hashtable
			hMetaData9.add("COLUMN_NAME", "MB_TEACHER" )
			hMetaData9.add("DB_TYPE", 16 )
			hMetaData9.add("PROVIDER_TYPE",253 )
			hMetaData9.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData9)
			Static hMetaData10 As new System.Collections.Hashtable
			hMetaData10.add("COLUMN_NAME", "MB_MEMO" )
			hMetaData10.add("DB_TYPE", 16 )
			hMetaData10.add("PROVIDER_TYPE",253 )
			hMetaData10.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData10)
			Static hMetaData11 As new System.Collections.Hashtable
			hMetaData11.add("COLUMN_NAME", "MB_YES" )
			hMetaData11.add("DB_TYPE", 16 )
			hMetaData11.add("PROVIDER_TYPE",254 )
			hMetaData11.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData11)
			Static hMetaData12 As new System.Collections.Hashtable
			hMetaData12.add("COLUMN_NAME", "MB_FULL" )
			hMetaData12.add("DB_TYPE", 7 )
			hMetaData12.add("PROVIDER_TYPE",246 )
			hMetaData12.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData12)
			Static hMetaData13 As new System.Collections.Hashtable
			hMetaData13.add("COLUMN_NAME", "MB_WAIT" )
			hMetaData13.add("DB_TYPE", 7 )
			hMetaData13.add("PROVIDER_TYPE",246 )
			hMetaData13.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData13)
			Static hMetaData14 As new System.Collections.Hashtable
			hMetaData14.add("COLUMN_NAME", "MB_APV_CNT" )
			hMetaData14.add("DB_TYPE", 7 )
			hMetaData14.add("PROVIDER_TYPE",246 )
			hMetaData14.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData14)
			Static hMetaData15 As new System.Collections.Hashtable
			hMetaData15.add("COLUMN_NAME", "MB_APV" )
			hMetaData15.add("DB_TYPE", 16 )
			hMetaData15.add("PROVIDER_TYPE",254 )
			hMetaData15.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData15)
			Static hMetaData16 As new System.Collections.Hashtable
			hMetaData16.add("COLUMN_NAME", "MB_OPEN" )
			hMetaData16.add("DB_TYPE", 16 )
			hMetaData16.add("PROVIDER_TYPE",254 )
			hMetaData16.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData16)
			Static hMetaData17 As new System.Collections.Hashtable
			hMetaData17.add("COLUMN_NAME", "MB_SAPLY" )
			hMetaData17.add("DB_TYPE", 6 )
			hMetaData17.add("PROVIDER_TYPE",10 )
			hMetaData17.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Date )
			arry.add(hMetaData17)
			Static hMetaData18 As new System.Collections.Hashtable
			hMetaData18.add("COLUMN_NAME", "MB_EAPLY" )
			hMetaData18.add("DB_TYPE", 6 )
			hMetaData18.add("PROVIDER_TYPE",10 )
			hMetaData18.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Date )
			arry.add(hMetaData18)
			Static hMetaData19 As new System.Collections.Hashtable
			hMetaData19.add("COLUMN_NAME", "MB_CDAYS" )
			hMetaData19.add("DB_TYPE", 7 )
			hMetaData19.add("PROVIDER_TYPE",246 )
			hMetaData19.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData19)
			Static hMetaData20 As new System.Collections.Hashtable
			hMetaData20.add("COLUMN_NAME", "MB_ALERT1_DAY" )
			hMetaData20.add("DB_TYPE", 7 )
			hMetaData20.add("PROVIDER_TYPE",246 )
			hMetaData20.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData20)
			Static hMetaData21 As new System.Collections.Hashtable
			hMetaData21.add("COLUMN_NAME", "MB_ALERT1" )
			hMetaData21.add("DB_TYPE", 16 )
			hMetaData21.add("PROVIDER_TYPE",254 )
			hMetaData21.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData21)
			Static hMetaData22 As new System.Collections.Hashtable
			hMetaData22.add("COLUMN_NAME", "MB_ALERT2_DAY" )
			hMetaData22.add("DB_TYPE", 7 )
			hMetaData22.add("PROVIDER_TYPE",246 )
			hMetaData22.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData22)
			Static hMetaData23 As new System.Collections.Hashtable
			hMetaData23.add("COLUMN_NAME", "MB_ALERT2" )
			hMetaData23.add("DB_TYPE", 16 )
			hMetaData23.add("PROVIDER_TYPE",254 )
			hMetaData23.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData23)
			Static hMetaData24 As new System.Collections.Hashtable
			hMetaData24.add("COLUMN_NAME", "REGTIME" )
			hMetaData24.add("DB_TYPE", 7 )
			hMetaData24.add("PROVIDER_TYPE",246 )
			hMetaData24.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData24)
			Static hMetaData25 As new System.Collections.Hashtable
			hMetaData25.add("COLUMN_NAME", "CONTACT" )
			hMetaData25.add("DB_TYPE", 16 )
			hMetaData25.add("PROVIDER_TYPE",253 )
			hMetaData25.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData25)
			Static hMetaData26 As new System.Collections.Hashtable
			hMetaData26.add("COLUMN_NAME", "CONTEL" )
			hMetaData26.add("DB_TYPE", 16 )
			hMetaData26.add("PROVIDER_TYPE",253 )
			hMetaData26.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData26)
			Static hMetaData27 As new System.Collections.Hashtable
			hMetaData27.add("COLUMN_NAME", "CLASS_PLACE" )
			hMetaData27.add("DB_TYPE", 16 )
			hMetaData27.add("PROVIDER_TYPE",253 )
			hMetaData27.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData27)
			Static hMetaData28 As new System.Collections.Hashtable
			hMetaData28.add("COLUMN_NAME", "TRAFFIC_DESC" )
			hMetaData28.add("DB_TYPE", 16 )
			hMetaData28.add("PROVIDER_TYPE",253 )
			hMetaData28.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData28)
     End Sub
     Public Function getMetaArray() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getMetaArray
         Return arry
     End Function
     Public sub setPrimaryKeys() Implements com.Azion.NET.VB.DBMetaData.setPrimaryKeys
       m_arryPrimaryKeys.add("MB_SEQ")
       m_arryPrimaryKeys.add("MB_BATCH")
     End Sub
     Public Function getPrimaryKeys() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getPrimaryKeys
         Return m_arryPrimaryKeys
     End Function
End Class
