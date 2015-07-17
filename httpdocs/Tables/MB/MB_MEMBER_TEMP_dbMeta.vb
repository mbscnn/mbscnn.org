Public Class MB_MEMBER_TEMP_dbMeta
        Implements com.Azion.NET.VB.DBMetaData
   Private  arry As New System.Collections.ArrayList
   Private  m_arryPrimaryKeys As New System.Collections.ArrayList
     Sub init() Implements com.Azion.NET.VB.DBMetaData.init
			Static hMetaData1 As new System.Collections.Hashtable
			hMetaData1.add("COLUMN_NAME", "MB_DATE" )
			hMetaData1.add("DB_TYPE", 7 )
			hMetaData1.add("PROVIDER_TYPE",246 )
			hMetaData1.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData1)
			Static hMetaData2 As new System.Collections.Hashtable
			hMetaData2.add("COLUMN_NAME", "MB_NAME" )
			hMetaData2.add("DB_TYPE", 16 )
			hMetaData2.add("PROVIDER_TYPE",253 )
			hMetaData2.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData2)
			Static hMetaData3 As new System.Collections.Hashtable
			hMetaData3.add("COLUMN_NAME", "MB_SEX" )
			hMetaData3.add("DB_TYPE", 16 )
			hMetaData3.add("PROVIDER_TYPE",254 )
			hMetaData3.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData3)
			Static hMetaData4 As new System.Collections.Hashtable
			hMetaData4.add("COLUMN_NAME", "MB_BIRTH" )
			hMetaData4.add("DB_TYPE", 6 )
			hMetaData4.add("PROVIDER_TYPE",10 )
			hMetaData4.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Date )
			arry.add(hMetaData4)
			Static hMetaData5 As new System.Collections.Hashtable
			hMetaData5.add("COLUMN_NAME", "MB_IDENTIFY" )
			hMetaData5.add("DB_TYPE", 16 )
			hMetaData5.add("PROVIDER_TYPE",253 )
			hMetaData5.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData5)
			Static hMetaData6 As new System.Collections.Hashtable
			hMetaData6.add("COLUMN_NAME", "MB_MOBIL" )
			hMetaData6.add("DB_TYPE", 16 )
			hMetaData6.add("PROVIDER_TYPE",253 )
			hMetaData6.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData6)
			Static hMetaData7 As new System.Collections.Hashtable
			hMetaData7.add("COLUMN_NAME", "MB_TEL" )
			hMetaData7.add("DB_TYPE", 16 )
			hMetaData7.add("PROVIDER_TYPE",253 )
			hMetaData7.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData7)
			Static hMetaData8 As new System.Collections.Hashtable
			hMetaData8.add("COLUMN_NAME", "MB_EMAIL" )
			hMetaData8.add("DB_TYPE", 16 )
			hMetaData8.add("PROVIDER_TYPE",253 )
			hMetaData8.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData8)
			Static hMetaData9 As new System.Collections.Hashtable
			hMetaData9.add("COLUMN_NAME", "MB_ID" )
			hMetaData9.add("DB_TYPE", 16 )
			hMetaData9.add("PROVIDER_TYPE",253 )
			hMetaData9.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData9)
			Static hMetaData10 As new System.Collections.Hashtable
			hMetaData10.add("COLUMN_NAME", "MB_EDU" )
			hMetaData10.add("DB_TYPE", 16 )
			hMetaData10.add("PROVIDER_TYPE",253 )
			hMetaData10.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData10)
			Static hMetaData11 As new System.Collections.Hashtable
			hMetaData11.add("COLUMN_NAME", "MB_REFER" )
			hMetaData11.add("DB_TYPE", 16 )
			hMetaData11.add("PROVIDER_TYPE",253 )
			hMetaData11.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData11)
			Static hMetaData12 As new System.Collections.Hashtable
			hMetaData12.add("COLUMN_NAME", "MB_CITY" )
			hMetaData12.add("DB_TYPE", 16 )
			hMetaData12.add("PROVIDER_TYPE",253 )
			hMetaData12.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData12)
			Static hMetaData13 As new System.Collections.Hashtable
			hMetaData13.add("COLUMN_NAME", "MB_VLG" )
			hMetaData13.add("DB_TYPE", 16 )
			hMetaData13.add("PROVIDER_TYPE",253 )
			hMetaData13.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData13)
			Static hMetaData14 As new System.Collections.Hashtable
			hMetaData14.add("COLUMN_NAME", "MB_ADDR" )
			hMetaData14.add("DB_TYPE", 16 )
			hMetaData14.add("PROVIDER_TYPE",253 )
			hMetaData14.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData14)
			Static hMetaData15 As new System.Collections.Hashtable
			hMetaData15.add("COLUMN_NAME", "MB_CITY1" )
			hMetaData15.add("DB_TYPE", 16 )
			hMetaData15.add("PROVIDER_TYPE",253 )
			hMetaData15.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData15)
			Static hMetaData16 As new System.Collections.Hashtable
			hMetaData16.add("COLUMN_NAME", "MB_VLG1" )
			hMetaData16.add("DB_TYPE", 16 )
			hMetaData16.add("PROVIDER_TYPE",253 )
			hMetaData16.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData16)
			Static hMetaData17 As new System.Collections.Hashtable
			hMetaData17.add("COLUMN_NAME", "MB_ADDR1" )
			hMetaData17.add("DB_TYPE", 16 )
			hMetaData17.add("PROVIDER_TYPE",253 )
			hMetaData17.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData17)
			Static hMetaData18 As new System.Collections.Hashtable
			hMetaData18.add("COLUMN_NAME", "MB_AREA" )
			hMetaData18.add("DB_TYPE", 16 )
			hMetaData18.add("PROVIDER_TYPE",253 )
			hMetaData18.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData18)
			Static hMetaData19 As new System.Collections.Hashtable
			hMetaData19.add("COLUMN_NAME", "MB_LEADER" )
			hMetaData19.add("DB_TYPE", 16 )
			hMetaData19.add("PROVIDER_TYPE",253 )
			hMetaData19.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData19)
			Static hMetaData20 As new System.Collections.Hashtable
			hMetaData20.add("COLUMN_NAME", "MB_DELFLG" )
			hMetaData20.add("DB_TYPE", 16 )
			hMetaData20.add("PROVIDER_TYPE",254 )
			hMetaData20.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData20)
			Static hMetaData21 As new System.Collections.Hashtable
			hMetaData21.add("COLUMN_NAME", "CHGUID" )
			hMetaData21.add("DB_TYPE", 16 )
			hMetaData21.add("PROVIDER_TYPE",253 )
			hMetaData21.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData21)
			Static hMetaData22 As new System.Collections.Hashtable
			hMetaData22.add("COLUMN_NAME", "CHGDATE" )
			hMetaData22.add("DB_TYPE", 6 )
			hMetaData22.add("PROVIDER_TYPE",12 )
			hMetaData22.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.DateTime )
			arry.add(hMetaData22)
			Static hMetaData23 As new System.Collections.Hashtable
			hMetaData23.add("COLUMN_NAME", "MB_ZIP" )
			hMetaData23.add("DB_TYPE", 16 )
			hMetaData23.add("PROVIDER_TYPE",253 )
			hMetaData23.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData23)
			Static hMetaData24 As new System.Collections.Hashtable
			hMetaData24.add("COLUMN_NAME", "MB_ZIP1" )
			hMetaData24.add("DB_TYPE", 16 )
			hMetaData24.add("PROVIDER_TYPE",253 )
			hMetaData24.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData24)
			Static hMetaData25 As new System.Collections.Hashtable
			hMetaData25.add("COLUMN_NAME", "MB_FAMILY" )
			hMetaData25.add("DB_TYPE", 16 )
			hMetaData25.add("PROVIDER_TYPE",254 )
			hMetaData25.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData25)
     End Sub
     Public Function getMetaArray() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getMetaArray
         Return arry
     End Function
     Public sub setPrimaryKeys() Implements com.Azion.NET.VB.DBMetaData.setPrimaryKeys
       m_arryPrimaryKeys.add("MB_DATE")
     End Sub
     Public Function getPrimaryKeys() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getPrimaryKeys
         Return m_arryPrimaryKeys
     End Function
End Class
