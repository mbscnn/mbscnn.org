Public Class MB_MEMBER_dbMeta
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
			hMetaData20.add("COLUMN_NAME", "MB_MONK" )
			hMetaData20.add("DB_TYPE", 16 )
			hMetaData20.add("PROVIDER_TYPE",254 )
			hMetaData20.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData20)
			Static hMetaData21 As new System.Collections.Hashtable
			hMetaData21.add("COLUMN_NAME", "MB_MONKNAME" )
			hMetaData21.add("DB_TYPE", 16 )
			hMetaData21.add("PROVIDER_TYPE",253 )
			hMetaData21.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData21)
			Static hMetaData22 As new System.Collections.Hashtable
			hMetaData22.add("COLUMN_NAME", "MB_MONKTECH" )
			hMetaData22.add("DB_TYPE", 16 )
			hMetaData22.add("PROVIDER_TYPE",253 )
			hMetaData22.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData22)
			Static hMetaData23 As new System.Collections.Hashtable
			hMetaData23.add("COLUMN_NAME", "MB_EDUTYPE" )
			hMetaData23.add("DB_TYPE", 16 )
			hMetaData23.add("PROVIDER_TYPE",253 )
			hMetaData23.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData23)
			Static hMetaData24 As new System.Collections.Hashtable
			hMetaData24.add("COLUMN_NAME", "MB_MONKTYPE" )
			hMetaData24.add("DB_TYPE", 16 )
			hMetaData24.add("PROVIDER_TYPE",253 )
			hMetaData24.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData24)
			Static hMetaData25 As new System.Collections.Hashtable
			hMetaData25.add("COLUMN_NAME", "MB_MONKPLACE" )
			hMetaData25.add("DB_TYPE", 16 )
			hMetaData25.add("PROVIDER_TYPE",253 )
			hMetaData25.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData25)
			Static hMetaData26 As new System.Collections.Hashtable
			hMetaData26.add("COLUMN_NAME", "MB_MONKDATE" )
			hMetaData26.add("DB_TYPE", 6 )
			hMetaData26.add("PROVIDER_TYPE",10 )
			hMetaData26.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.Date )
			arry.add(hMetaData26)
			Static hMetaData27 As new System.Collections.Hashtable
			hMetaData27.add("COLUMN_NAME", "MB_MONKPLACE1" )
			hMetaData27.add("DB_TYPE", 16 )
			hMetaData27.add("PROVIDER_TYPE",253 )
			hMetaData27.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData27)
			Static hMetaData28 As new System.Collections.Hashtable
			hMetaData28.add("COLUMN_NAME", "MB_LANG" )
			hMetaData28.add("DB_TYPE", 16 )
			hMetaData28.add("PROVIDER_TYPE",253 )
			hMetaData28.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData28)
			Static hMetaData29 As new System.Collections.Hashtable
			hMetaData29.add("COLUMN_NAME", "MB_OLANG" )
			hMetaData29.add("DB_TYPE", 16 )
			hMetaData29.add("PROVIDER_TYPE",253 )
			hMetaData29.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData29)
			Static hMetaData30 As new System.Collections.Hashtable
			hMetaData30.add("COLUMN_NAME", "MB_SPECIAL" )
			hMetaData30.add("DB_TYPE", 16 )
			hMetaData30.add("PROVIDER_TYPE",253 )
			hMetaData30.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData30)
			Static hMetaData31 As new System.Collections.Hashtable
			hMetaData31.add("COLUMN_NAME", "MB_PROFESSION" )
			hMetaData31.add("DB_TYPE", 16 )
			hMetaData31.add("PROVIDER_TYPE",253 )
			hMetaData31.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData31)
			Static hMetaData32 As new System.Collections.Hashtable
			hMetaData32.add("COLUMN_NAME", "MB_SICK" )
			hMetaData32.add("DB_TYPE", 16 )
			hMetaData32.add("PROVIDER_TYPE",253 )
			hMetaData32.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData32)
			Static hMetaData33 As new System.Collections.Hashtable
			hMetaData33.add("COLUMN_NAME", "MB_ALLERGY" )
			hMetaData33.add("DB_TYPE", 16 )
			hMetaData33.add("PROVIDER_TYPE",253 )
			hMetaData33.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData33)
			Static hMetaData34 As new System.Collections.Hashtable
			hMetaData34.add("COLUMN_NAME", "MB_OPERATE" )
			hMetaData34.add("DB_TYPE", 16 )
			hMetaData34.add("PROVIDER_TYPE",253 )
			hMetaData34.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData34)
			Static hMetaData35 As new System.Collections.Hashtable
			hMetaData35.add("COLUMN_NAME", "MB_OSICK" )
			hMetaData35.add("DB_TYPE", 16 )
			hMetaData35.add("PROVIDER_TYPE",253 )
			hMetaData35.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData35)
			Static hMetaData36 As new System.Collections.Hashtable
			hMetaData36.add("COLUMN_NAME", "MB_PIPOSHENA" )
			hMetaData36.add("DB_TYPE", 16 )
			hMetaData36.add("PROVIDER_TYPE",254 )
			hMetaData36.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData36)
			Static hMetaData37 As new System.Collections.Hashtable
			hMetaData37.add("COLUMN_NAME", "MB_TEACH" )
			hMetaData37.add("DB_TYPE", 16 )
			hMetaData37.add("PROVIDER_TYPE",253 )
			hMetaData37.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData37)
			Static hMetaData38 As new System.Collections.Hashtable
			hMetaData38.add("COLUMN_NAME", "MB_FAMENNIAN" )
			hMetaData38.add("DB_TYPE", 16 )
			hMetaData38.add("PROVIDER_TYPE",253 )
			hMetaData38.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData38)
			Static hMetaData39 As new System.Collections.Hashtable
			hMetaData39.add("COLUMN_NAME", "MB_OVER7DAY" )
			hMetaData39.add("DB_TYPE", 16 )
			hMetaData39.add("PROVIDER_TYPE",254 )
			hMetaData39.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData39)
			Static hMetaData40 As new System.Collections.Hashtable
			hMetaData40.add("COLUMN_NAME", "MB_PLACE" )
			hMetaData40.add("DB_TYPE", 16 )
			hMetaData40.add("PROVIDER_TYPE",253 )
			hMetaData40.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData40)
			Static hMetaData41 As new System.Collections.Hashtable
			hMetaData41.add("COLUMN_NAME", "MB_EMGCONT" )
			hMetaData41.add("DB_TYPE", 16 )
			hMetaData41.add("PROVIDER_TYPE",253 )
			hMetaData41.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData41)
			Static hMetaData42 As new System.Collections.Hashtable
			hMetaData42.add("COLUMN_NAME", "MB_CONTID" )
			hMetaData42.add("DB_TYPE", 16 )
			hMetaData42.add("PROVIDER_TYPE",253 )
			hMetaData42.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData42)
			Static hMetaData43 As new System.Collections.Hashtable
			hMetaData43.add("COLUMN_NAME", "MB_RELATE" )
			hMetaData43.add("DB_TYPE", 16 )
			hMetaData43.add("PROVIDER_TYPE",253 )
			hMetaData43.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData43)
			Static hMetaData44 As new System.Collections.Hashtable
			hMetaData44.add("COLUMN_NAME", "MB_CONTTEL_D" )
			hMetaData44.add("DB_TYPE", 16 )
			hMetaData44.add("PROVIDER_TYPE",253 )
			hMetaData44.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData44)
			Static hMetaData45 As new System.Collections.Hashtable
			hMetaData45.add("COLUMN_NAME", "MB_CONTTEL_N" )
			hMetaData45.add("DB_TYPE", 16 )
			hMetaData45.add("PROVIDER_TYPE",253 )
			hMetaData45.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData45)
			Static hMetaData46 As new System.Collections.Hashtable
			hMetaData46.add("COLUMN_NAME", "MB_CONTMOBIL" )
			hMetaData46.add("DB_TYPE", 16 )
			hMetaData46.add("PROVIDER_TYPE",253 )
			hMetaData46.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData46)
			Static hMetaData47 As new System.Collections.Hashtable
			hMetaData47.add("COLUMN_NAME", "MB_CONTFAX" )
			hMetaData47.add("DB_TYPE", 16 )
			hMetaData47.add("PROVIDER_TYPE",253 )
			hMetaData47.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData47)
			Static hMetaData48 As new System.Collections.Hashtable
			hMetaData48.add("COLUMN_NAME", "MB_CONT_CITY" )
			hMetaData48.add("DB_TYPE", 16 )
			hMetaData48.add("PROVIDER_TYPE",253 )
			hMetaData48.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData48)
			Static hMetaData49 As new System.Collections.Hashtable
			hMetaData49.add("COLUMN_NAME", "MB_CONT_VLG" )
			hMetaData49.add("DB_TYPE", 16 )
			hMetaData49.add("PROVIDER_TYPE",253 )
			hMetaData49.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData49)
			Static hMetaData50 As new System.Collections.Hashtable
			hMetaData50.add("COLUMN_NAME", "MB_CONT_ADDR" )
			hMetaData50.add("DB_TYPE", 16 )
			hMetaData50.add("PROVIDER_TYPE",253 )
			hMetaData50.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData50)
			Static hMetaData51 As new System.Collections.Hashtable
			hMetaData51.add("COLUMN_NAME", "MB_JOIN_ITEM" )
			hMetaData51.add("DB_TYPE", 16 )
			hMetaData51.add("PROVIDER_TYPE",253 )
			hMetaData51.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData51)
			Static hMetaData52 As new System.Collections.Hashtable
			hMetaData52.add("COLUMN_NAME", "MB_JOINOTH" )
			hMetaData52.add("DB_TYPE", 16 )
			hMetaData52.add("PROVIDER_TYPE",253 )
			hMetaData52.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData52)
			Static hMetaData53 As new System.Collections.Hashtable
			hMetaData53.add("COLUMN_NAME", "MB_FAMILY" )
			hMetaData53.add("DB_TYPE", 16 )
			hMetaData53.add("PROVIDER_TYPE",254 )
			hMetaData53.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData53)
			Static hMetaData54 As new System.Collections.Hashtable
			hMetaData54.add("COLUMN_NAME", "CHGUID" )
			hMetaData54.add("DB_TYPE", 16 )
			hMetaData54.add("PROVIDER_TYPE",253 )
			hMetaData54.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData54)
			Static hMetaData55 As new System.Collections.Hashtable
			hMetaData55.add("COLUMN_NAME", "CHGDATE" )
			hMetaData55.add("DB_TYPE", 6 )
			hMetaData55.add("PROVIDER_TYPE",12 )
			hMetaData55.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.DateTime )
			arry.add(hMetaData55)
			Static hMetaData56 As new System.Collections.Hashtable
			hMetaData56.add("COLUMN_NAME", "MB_SNORE" )
			hMetaData56.add("DB_TYPE", 16 )
			hMetaData56.add("PROVIDER_TYPE",254 )
			hMetaData56.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData56)
			Static hMetaData57 As new System.Collections.Hashtable
			hMetaData57.add("COLUMN_NAME", "MB_JOBTITLE" )
			hMetaData57.add("DB_TYPE", 16 )
			hMetaData57.add("PROVIDER_TYPE",253 )
			hMetaData57.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData57)
			Static hMetaData58 As new System.Collections.Hashtable
			hMetaData58.add("COLUMN_NAME", "MB_RELIGION" )
			hMetaData58.add("DB_TYPE", 16 )
			hMetaData58.add("PROVIDER_TYPE",253 )
			hMetaData58.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData58)
			Static hMetaData59 As new System.Collections.Hashtable
			hMetaData59.add("COLUMN_NAME", "CREDATE" )
			hMetaData59.add("DB_TYPE", 6 )
			hMetaData59.add("PROVIDER_TYPE",12 )
			hMetaData59.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.DateTime )
			arry.add(hMetaData59)
			Static hMetaData60 As new System.Collections.Hashtable
			hMetaData60.add("COLUMN_NAME", "MB_ZIP" )
			hMetaData60.add("DB_TYPE", 16 )
			hMetaData60.add("PROVIDER_TYPE",253 )
			hMetaData60.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData60)
			Static hMetaData61 As new System.Collections.Hashtable
			hMetaData61.add("COLUMN_NAME", "MB_ZIP1" )
			hMetaData61.add("DB_TYPE", 16 )
			hMetaData61.add("PROVIDER_TYPE",253 )
			hMetaData61.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData61)
			Static hMetaData62 As new System.Collections.Hashtable
			hMetaData62.add("COLUMN_NAME", "SCHOOL" )
			hMetaData62.add("DB_TYPE", 16 )
			hMetaData62.add("PROVIDER_TYPE",253 )
			hMetaData62.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData62)
			Static hMetaData63 As new System.Collections.Hashtable
			hMetaData63.add("COLUMN_NAME", "JOINMBSC" )
			hMetaData63.add("DB_TYPE", 16 )
			hMetaData63.add("PROVIDER_TYPE",254 )
			hMetaData63.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData63)
     End Sub
     Public Function getMetaArray() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getMetaArray
         Return arry
     End Function
     Public sub setPrimaryKeys() Implements com.Azion.NET.VB.DBMetaData.setPrimaryKeys
       m_arryPrimaryKeys.add("MB_MEMSEQ")
     End Sub
     Public Function getPrimaryKeys() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getPrimaryKeys
         Return m_arryPrimaryKeys
     End Function
End Class
