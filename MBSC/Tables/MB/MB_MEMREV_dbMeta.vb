Public Class MB_MEMREV_dbMeta
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
			hMetaData2.add("COLUMN_NAME", "MB_SEQNO" )
			hMetaData2.add("DB_TYPE", 7 )
			hMetaData2.add("PROVIDER_TYPE",246 )
			hMetaData2.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData2)
			Static hMetaData3 As new System.Collections.Hashtable
			hMetaData3.add("COLUMN_NAME", "MB_ITEMID" )
			hMetaData3.add("DB_TYPE", 16 )
			hMetaData3.add("PROVIDER_TYPE",254 )
			hMetaData3.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData3)
			Static hMetaData4 As new System.Collections.Hashtable
			hMetaData4.add("COLUMN_NAME", "MB_MEMFEE_SY" )
			hMetaData4.add("DB_TYPE", 7 )
			hMetaData4.add("PROVIDER_TYPE",246 )
			hMetaData4.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData4)
			Static hMetaData5 As new System.Collections.Hashtable
			hMetaData5.add("COLUMN_NAME", "MB_MEMFEE_SM" )
			hMetaData5.add("DB_TYPE", 7 )
			hMetaData5.add("PROVIDER_TYPE",246 )
			hMetaData5.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData5)
			Static hMetaData6 As new System.Collections.Hashtable
			hMetaData6.add("COLUMN_NAME", "MB_MEMFEE_EY" )
			hMetaData6.add("DB_TYPE", 7 )
			hMetaData6.add("PROVIDER_TYPE",246 )
			hMetaData6.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData6)
			Static hMetaData7 As new System.Collections.Hashtable
			hMetaData7.add("COLUMN_NAME", "MB_MEMFEE_EM" )
			hMetaData7.add("DB_TYPE", 7 )
			hMetaData7.add("PROVIDER_TYPE",246 )
			hMetaData7.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData7)
			Static hMetaData8 As new System.Collections.Hashtable
			hMetaData8.add("COLUMN_NAME", "MB_FEETYPE" )
			hMetaData8.add("DB_TYPE", 16 )
			hMetaData8.add("PROVIDER_TYPE",254 )
			hMetaData8.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData8)
			Static hMetaData9 As new System.Collections.Hashtable
			hMetaData9.add("COLUMN_NAME", "MB_TOTFEE" )
			hMetaData9.add("DB_TYPE", 7 )
			hMetaData9.add("PROVIDER_TYPE",246 )
			hMetaData9.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData9)
			Static hMetaData10 As new System.Collections.Hashtable
			hMetaData10.add("COLUMN_NAME", "MB_MEMTYP" )
			hMetaData10.add("DB_TYPE", 16 )
			hMetaData10.add("PROVIDER_TYPE",254 )
			hMetaData10.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData10)
			Static hMetaData11 As new System.Collections.Hashtable
			hMetaData11.add("COLUMN_NAME", "MB_RECNAME" )
			hMetaData11.add("DB_TYPE", 16 )
			hMetaData11.add("PROVIDER_TYPE",253 )
			hMetaData11.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData11)
			Static hMetaData12 As new System.Collections.Hashtable
			hMetaData12.add("COLUMN_NAME", "MB_DESC" )
			hMetaData12.add("DB_TYPE", 16 )
			hMetaData12.add("PROVIDER_TYPE",253 )
			hMetaData12.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData12)
			Static hMetaData13 As new System.Collections.Hashtable
			hMetaData13.add("COLUMN_NAME", "MB_PAY_TYPE" )
			hMetaData13.add("DB_TYPE", 16 )
			hMetaData13.add("PROVIDER_TYPE",254 )
			hMetaData13.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData13)
			Static hMetaData14 As new System.Collections.Hashtable
			hMetaData14.add("COLUMN_NAME", "MB_VOUCHER_Y" )
			hMetaData14.add("DB_TYPE", 7 )
			hMetaData14.add("PROVIDER_TYPE",246 )
			hMetaData14.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData14)
			Static hMetaData15 As new System.Collections.Hashtable
			hMetaData15.add("COLUMN_NAME", "MB_VOUCHER_N" )
			hMetaData15.add("DB_TYPE", 7 )
			hMetaData15.add("PROVIDER_TYPE",246 )
			hMetaData15.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData15)
			Static hMetaData16 As new System.Collections.Hashtable
			hMetaData16.add("COLUMN_NAME", "MB_PRINT" )
			hMetaData16.add("DB_TYPE", 16 )
			hMetaData16.add("PROVIDER_TYPE",254 )
			hMetaData16.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData16)
			Static hMetaData17 As new System.Collections.Hashtable
			hMetaData17.add("COLUMN_NAME", "MB_REISU" )
			hMetaData17.add("DB_TYPE", 7 )
			hMetaData17.add("PROVIDER_TYPE",246 )
			hMetaData17.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData17)
			Static hMetaData18 As new System.Collections.Hashtable
			hMetaData18.add("COLUMN_NAME", "MB_REBREV" )
			hMetaData18.add("DB_TYPE", 16 )
			hMetaData18.add("PROVIDER_TYPE",254 )
			hMetaData18.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData18)
			Static hMetaData19 As new System.Collections.Hashtable
			hMetaData19.add("COLUMN_NAME", "DELFLAG" )
			hMetaData19.add("DB_TYPE", 16 )
			hMetaData19.add("PROVIDER_TYPE",254 )
			hMetaData19.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData19)
			Static hMetaData20 As new System.Collections.Hashtable
			hMetaData20.add("COLUMN_NAME", "MB_TX_DATE" )
			hMetaData20.add("DB_TYPE", 7 )
			hMetaData20.add("PROVIDER_TYPE",246 )
			hMetaData20.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
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
			hMetaData23.add("COLUMN_NAME", "APVUID" )
			hMetaData23.add("DB_TYPE", 16 )
			hMetaData23.add("PROVIDER_TYPE",253 )
			hMetaData23.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData23)
			Static hMetaData24 As new System.Collections.Hashtable
			hMetaData24.add("COLUMN_NAME", "APVDATE" )
			hMetaData24.add("DB_TYPE", 6 )
			hMetaData24.add("PROVIDER_TYPE",12 )
			hMetaData24.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.DateTime )
			arry.add(hMetaData24)
			Static hMetaData25 As new System.Collections.Hashtable
			hMetaData25.add("COLUMN_NAME", "MB_CURCODE" )
			hMetaData25.add("DB_TYPE", 16 )
			hMetaData25.add("PROVIDER_TYPE",253 )
			hMetaData25.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData25)
			Static hMetaData26 As new System.Collections.Hashtable
			hMetaData26.add("COLUMN_NAME", "NOTE_DUE_DATE" )
			hMetaData26.add("DB_TYPE", 7 )
			hMetaData26.add("PROVIDER_TYPE",246 )
			hMetaData26.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData26)
			Static hMetaData27 As new System.Collections.Hashtable
			hMetaData27.add("COLUMN_NAME", "NOTE_NO" )
			hMetaData27.add("DB_TYPE", 16 )
			hMetaData27.add("PROVIDER_TYPE",253 )
			hMetaData27.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData27)
			Static hMetaData28 As new System.Collections.Hashtable
			hMetaData28.add("COLUMN_NAME", "NOTE_BANK" )
			hMetaData28.add("DB_TYPE", 16 )
			hMetaData28.add("PROVIDER_TYPE",253 )
			hMetaData28.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData28)
			Static hMetaData29 As new System.Collections.Hashtable
			hMetaData29.add("COLUMN_NAME", "NOTE_BR" )
			hMetaData29.add("DB_TYPE", 16 )
			hMetaData29.add("PROVIDER_TYPE",253 )
			hMetaData29.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData29)
			Static hMetaData30 As new System.Collections.Hashtable
			hMetaData30.add("COLUMN_NAME", "NOTE_HOLDER" )
			hMetaData30.add("DB_TYPE", 16 )
			hMetaData30.add("PROVIDER_TYPE",253 )
			hMetaData30.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData30)
			Static hMetaData31 As new System.Collections.Hashtable
			hMetaData31.add("COLUMN_NAME", "NOTE_AMT" )
			hMetaData31.add("DB_TYPE", 7 )
			hMetaData31.add("PROVIDER_TYPE",246 )
			hMetaData31.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData31)
			Static hMetaData32 As new System.Collections.Hashtable
			hMetaData32.add("COLUMN_NAME", "NOTE_CASH" )
			hMetaData32.add("DB_TYPE", 16 )
			hMetaData32.add("PROVIDER_TYPE",254 )
			hMetaData32.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData32)
			Static hMetaData33 As new System.Collections.Hashtable
			hMetaData33.add("COLUMN_NAME", "MB_RECV_Y" )
			hMetaData33.add("DB_TYPE", 7 )
			hMetaData33.add("PROVIDER_TYPE",246 )
			hMetaData33.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData33)
			Static hMetaData34 As new System.Collections.Hashtable
			hMetaData34.add("COLUMN_NAME", "MB_RECV_N" )
			hMetaData34.add("DB_TYPE", 7 )
			hMetaData34.add("PROVIDER_TYPE",246 )
			hMetaData34.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData34)
			Static hMetaData35 As new System.Collections.Hashtable
			hMetaData35.add("COLUMN_NAME", "MB_MONFEE" )
			hMetaData35.add("DB_TYPE", 7 )
			hMetaData35.add("PROVIDER_TYPE",246 )
			hMetaData35.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.NewDecimal )
			arry.add(hMetaData35)
			Static hMetaData36 As new System.Collections.Hashtable
			hMetaData36.add("COLUMN_NAME", "MB_SEND_PRINT" )
			hMetaData36.add("DB_TYPE", 16 )
			hMetaData36.add("PROVIDER_TYPE",254 )
			hMetaData36.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData36)
			Static hMetaData37 As new System.Collections.Hashtable
			hMetaData37.add("COLUMN_NAME", "VRYUID" )
			hMetaData37.add("DB_TYPE", 16 )
			hMetaData37.add("PROVIDER_TYPE",253 )
			hMetaData37.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.VarChar )
			arry.add(hMetaData37)
			Static hMetaData38 As new System.Collections.Hashtable
			hMetaData38.add("COLUMN_NAME", "VRYDATE" )
			hMetaData38.add("DB_TYPE", 6 )
			hMetaData38.add("PROVIDER_TYPE",12 )
			hMetaData38.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.DateTime )
			arry.add(hMetaData38)
			Static hMetaData39 As new System.Collections.Hashtable
			hMetaData39.add("COLUMN_NAME", "FUND1" )
			hMetaData39.add("DB_TYPE", 16 )
			hMetaData39.add("PROVIDER_TYPE",254 )
			hMetaData39.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData39)
			Static hMetaData40 As new System.Collections.Hashtable
			hMetaData40.add("COLUMN_NAME", "FUND2" )
			hMetaData40.add("DB_TYPE", 16 )
			hMetaData40.add("PROVIDER_TYPE",254 )
			hMetaData40.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData40)
			Static hMetaData41 As new System.Collections.Hashtable
			hMetaData41.add("COLUMN_NAME", "FUND3" )
			hMetaData41.add("DB_TYPE", 16 )
			hMetaData41.add("PROVIDER_TYPE",254 )
			hMetaData41.add("PROVIDER_TYPE_NAME", MySql.Data.MySqlClient.MySqlDbType.String )
			arry.add(hMetaData41)
     End Sub
     Public Function getMetaArray() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getMetaArray
         Return arry
     End Function
     Public sub setPrimaryKeys() Implements com.Azion.NET.VB.DBMetaData.setPrimaryKeys
       m_arryPrimaryKeys.add("MB_MEMSEQ")
       m_arryPrimaryKeys.add("MB_SEQNO")
     End Sub
     Public Function getPrimaryKeys() as System.Collections.ArrayList Implements com.Azion.NET.VB.DBMetaData.getPrimaryKeys
         Return m_arryPrimaryKeys
     End Function
End Class
