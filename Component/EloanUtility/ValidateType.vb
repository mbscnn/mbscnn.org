Option Explicit On
Option Strict On

Public Enum ValidateType

    ''' <summary>
    ''' 必填驗證
    ''' </summary>
    ''' <remarks></remarks>
    Required

    ''' <summary>
    ''' 驗證RadioButtonList必選
    ''' </summary>
    ''' <remarks></remarks>
    RequiredRadioList

    ''' <summary>
    ''' 英文數字驗證
    ''' </summary>
    ''' <remarks></remarks>
    LetterNumber

    ''' <summary>
    ''' 驗證網路地址
    ''' </summary>
    ''' <remarks></remarks>
    URL

    ''' <summary>
    ''' 驗證日期
    ''' </summary>
    ''' <remarks></remarks>
    [Date]

    ''' <summary>
    ''' 驗證數字
    ''' </summary>
    ''' <remarks></remarks>
    Number

    ''' <summary>
    ''' 驗證金額(最多15位整數，最多4位小數)
    ''' </summary>
    ''' <remarks></remarks>
    Currency

    ''' <summary>
    ''' 驗證金額(最多15位整數，最多2位小數)
    ''' </summary>
    ''' <remarks>Add by Avril 2012/03/15</remarks>
    Currency152

    ''' <summary>
    ''' 驗證金額(最多15位整數，最多0位小數)
    ''' </summary>
    ''' <remarks>Add by Avril 2012/03/15</remarks>
    Currency150

    ''' <summary>
    ''' 驗證金額(最多13位整數，最多0位小數)
    ''' </summary>
    ''' <remarks>Add by Avril 2012/03/23</remarks>
    Currency130

    ''' <summary>
    ''' 驗證金額(最多12位整數，最多0位小數)
    ''' </summary>
    ''' <remarks>Add by Avril 2012/04/09</remarks>
    Currency120

    ''' <summary>
    ''' 驗證金額(最多18位整數，最多3位小數)
    ''' </summary>
    ''' <remarks>Add by Avril 2012/04/18</remarks>
    Currency183

    ''' <summary>
    ''' 驗證金額(最多10位整數，最多0位小數)
    ''' </summary>
    ''' <remarks>Add by Avril 2012/05/30</remarks>
    Currency100

    ''' <summary>
    ''' 驗證持股比率(最多13位整數，最多1位小數)
    ''' </summary>
    ''' <remarks></remarks>
    Tax131

    ''' <summary>
    ''' 驗證不為0
    ''' </summary>
    ''' <remarks></remarks>
    NoZero

    ''' <summary>
    ''' 驗證稅率格式declmal(4,2)
    ''' </summary>
    ''' <remarks></remarks>
    TaxRate

    ''' <summary>
    ''' 手續費率numeric(5,2)
    ''' </summary>
    ''' <remarks>Add by Avril 2012/03/15</remarks>
    TaxRate52

    ''' <summary>
    ''' 驗證稅率格式declmal(3,2) 
    ''' </summary>
    ''' <remarks></remarks>
    ManPower

    ''' <summary>
    ''' 驗證稅率格式declmal(3,2) ,可輸入負數
    ''' </summary>
    ''' <remarks></remarks>
    NManPower

    ''' <summary>
    ''' 驗證稅率格式declmal(7,4)
    ''' </summary>
    ''' <remarks></remarks>
    Rate

    ''' <summary>
    ''' 驗證金額格式declmal(10,4)
    ''' </summary>
    ''' <remarks></remarks>
    Money

    ''' <summary>
    ''' 驗證數字格式declmal(10,3)
    ''' </summary>
    ''' <remarks></remarks>
    Qty

    ''' <summary>
    ''' 驗證業務人員的編號及英文名稱組成的字符串
    ''' </summary>
    ''' <remarks></remarks>
    SysUser

    ''' <summary>
    ''' 驗證declmal(17,2)
    ''' </summary>
    ''' <remarks></remarks>
    CurrencyCommon

    ''' <summary>
    '''  驗證declmal(19,4)
    ''' </summary>
    ''' <remarks></remarks>
    Tax

    ''' <summary>
    ''' 只能輸入數字(帶小數的)
    ''' </summary>
    ''' <remarks></remarks>
    Numbers

    ''' <summary>
    ''' 最多24位數字
    ''' </summary>
    ''' <remarks></remarks>
    Number24

    ''' <summary>
    ''' 超過指定的日期
    ''' </summary>
    ''' <remarks></remarks>
    UpDate

    ''' <summary>
    ''' 驗證不大於
    ''' </summary>
    ''' <remarks></remarks>
    NoUpValue

    ''' <summary>
    ''' (比較兩個欄位)驗證兩個欄位相等
    ''' </summary>
    ''' <remarks></remarks>
    CompareEqules

    ''' <summary>
    '''  (比較兩個欄位)驗證兩欄位不相等
    ''' </summary>
    ''' <remarks></remarks>
    CompareNoEqules

    ''' <summary>
    '''  (比較兩個欄位)驗證可空日期的比較,為空或者小於等於
    ''' </summary>
    ''' <remarks></remarks>
    EmptyOrLessDay

    ''' <summary>
    ''' (比較兩個欄位)驗證可空日期的比較,為空或者小於等於
    ''' </summary>
    ''' <remarks></remarks>
    EmptyOrLessMonth

    ''' <summary>
    ''' (比較兩個欄位)驗證可空日期的比較,為空或者大於等於
    ''' </summary>
    ''' <remarks></remarks>
    EmptyOrUperDay

    ''' <summary>
    ''' (比較兩個欄位)驗證開始日期全部輸入或全部為空 
    ''' </summary>
    ''' <remarks></remarks>
    BothBeginEnd

    ''' <summary>
    '''  (比較兩個欄位)驗證迄日期是否大於起日
    ''' </summary>
    ''' <remarks></remarks>
    CompareDate

    ''' <summary>
    '''  (比較兩個欄位)驗證可空日期的比較,為空或者大於
    ''' </summary>
    ''' <remarks></remarks>
    UperDay

    ''' <summary>
    '''  (比較兩個欄位)兩個按鈕中是否有一個選中項
    ''' </summary>
    ''' <remarks></remarks>
    IsUse

    ''' <summary>
    '''  驗證固定字節數
    ''' </summary>
    ''' <remarks></remarks>
    ByteRangeLength

    ''' <summary>
    ''' 最大長度
    ''' </summary>
    ''' <remarks></remarks>
    MaxLength

    ''' <summary>
    ''' 最小長度
    ''' </summary>
    ''' <remarks></remarks>
    MinLength

    ''' <summary>
    ''' (比較兩個欄位)兩個文本框欄位至少輸入一個
    ''' </summary>
    ''' <remarks></remarks>
    InputOne

    ''' <summary>
    ''' 大於等於某個數字
    ''' </summary>
    ''' <remarks></remarks>
    GreatEqual

    ''' <summary>
    ''' 大於某個數字
    ''' </summary>
    ''' <remarks></remarks>
    GreatThan

    ''' <summary>
    ''' 小於等於某個數字
    ''' </summary>
    ''' <remarks></remarks>
    LessEqual

    ''' <summary>
    ''' 小於某個數字
    ''' </summary>
    ''' <remarks></remarks>
    LessThan

    ''' <summary>
    ''' 驗證輸入內容是否為星號(*)
    ''' </summary>
    ''' <remarks></remarks>
    IsStar

    ''' <summary>
    ''' 驗證輸入內容是否為0或整數倍數
    ''' </summary>
    ''' <remarks></remarks>
    IsZero

    ''' <summary>
    ''' 驗證日期欄位小於特定的日期
    ''' </summary>
    ''' <remarks></remarks>
    LessThanSpecificDate

    ''' <summary>
    ''' 驗證日期欄位小於等於特定的日期
    ''' </summary>
    ''' <remarks></remarks>
    LessEqualSpecificDate

    ''' <summary>
    ''' 驗證日期欄位大於特定的日期
    ''' </summary>
    ''' <remarks></remarks>
    GreatThanSpecificDate

    ''' <summary>
    ''' 驗證日期欄位大於等於特定的日期
    ''' </summary>
    ''' <remarks></remarks>
    GreatEqualSpecificDate

    ''' <summary>
    '''  驗證兩個控件只能輸入一個
    ''' </summary>
    ''' <remarks></remarks>
    JustOneValue

    ''' <summary>
    ''' 驗證數字(最多6位整數，最多0位小數)
    ''' </summary>
    ''' <remarks>Add by Avril 2012/05/02</remarks>
    Currency60

    ''' <summary>
    ''' 比率驗證decimal(5,3)
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/08 Add</remarks>
    TaxRate53

    ''' <summary>
    ''' 比率驗證decimal(6,3)
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/08 Add</remarks>
    TaxRate63

    ''' <summary>
    ''' 手續費率numeric(6,2) 
    ''' </summary>
    ''' <remarks></remarks>
    TaxRate62
End Enum
