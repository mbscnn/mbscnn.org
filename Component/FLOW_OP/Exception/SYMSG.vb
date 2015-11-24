Public Enum SYMSG As Integer
    SUCCESS = &H0                ' 成功
    ALERT = &H1                  ' 警告
    ALERTWINDOW = &H2            ' 警告，只顯示在彈出視窗

    UNDEFINE = &H80000000               ' MESSAGE CODE未定義
    FATAL_ERROR = &H80000001            ' 嚴重錯誤


    SQL_EXCEPTION = &H80001000                          ' An exception is thrown by dbUtility.
    SQL_INSERT_EXCEPTION = &H80001001                   ' An exception is thrown by dbUtility.
    SQL_UPDATE_EXCEPTION = &H80001002                   ' An exception is thrown by dbUtility.
    SQL_INSERTUPDATE_EXCEPTION = &H80001003             ' An exception is thrown by dbUtility.

    BOSBASE_EXCEPTION = &H80002000

    ERROR_INVALID_PARAMETER = &H87    '錯誤參數

    SYFLOW_CASE_NOT_FOUND = &H10000000                        '找不到擁有者的案件
    SYFLOW_CASENOTFOUND_PROHIBITESAMEUSER = &H10000001      '找不到擁有者的案件，原因是有找到案件的擁有者，但因為與上一步擁的送出者相同所以被禁止擁有
    SYFLOW_CREATE_NEW_CASEID_FAILED = &H10000010              '取新案號失敗

    SYCASEID_ISCLOSE = &H10000100                            '案件的狀態已關閉，不可以再改變
    SYCASEID_CANNOT_CHANGE = &H10000101                        '案件的狀態無法改變，或找不到可以改變狀態的案件
    SYCASEID_CANNOT_CHANGECLIENT = &H10000200                  '案件的無法改分配給其它人
    SYCASEID_CANNOT_CHANGECLIENT_CSAEUNFOUND = &H10000201      '案件的無法改分配給其它人，找不到案件
    SYCASEID_CANNOT_CHANGECLIENT_AMBIGUOUSSTEP = &H10000202    '案件的無法改分配給其它人，案件有多個步驟可以改分派，但最多只能改分派一個步驟
    SYCASEID_BRID_NOT_FOUND = &H10000301                        '無法取得分行代碼
    SYCASEID_WRONG_FORMAT = &H10000302                          '取號的號碼格式錯誤
    SYCASEID_WRONG_DUPLICATION = &H10000303                     '取號的號碼重複而且重新取號100次仍然無法取得新案號

    SYCASEREVISION_CANNOT_INSERTREVISION = &H10010000    '修正補充無法寫入資料

    SYCONDITIONID_CANNOT_GETINFO = &H10020000              '無法取得流程步驟的條件內容

    SYFLOWSTEP_NEXTSTEP_NOUSER = &H10030000         '下一步驟沒有USER
    SYFLOWSTEP_NEXTSTEP_NOVALIDUSER = &H10030001    '下一步驟沒有USER，下一步驟旳USER不可以跟上一步驗相同
    SYFLOWSTEP_NEXTSTEP_NOSTEP = &H10030010         '沒有下一步驟
    SYFLOWSTEP_SUBFLOWSEQ_NOT_FOUND = &H10030020    '無法取得子流程編號
    SYFLOWSTEP_SUBFLOWSEQ_NOT_FOUND_AMBIGUOUS = &H10030021    '無法從CASEID及取得子流程編號(子流程編號數目超過1，子流程編號數目不能超過1)
    SYFLOWSTEP_FIRSTSTEP_NOT_FOUND = &H10030030     '無法取得最先一個步驟代碼
    SYFLOWSTEP_LASTSTEP_NOT_FOUND = &H10030031      '無法取得最先一個步驟代碼
    SYFLOWSTEP_BRID_NOT_FOUND = &H10030040          '無法找到使用者的分行代碼
    SYFLOWSTEP_SPLITTER_NOT_FOUND = &H10030041      '無法找到平行流程的分支點
    SYFLOWSTEP_INCONSISTENT_STEP = &H10030100            '案件現在的步驟不一致，案件可能重複送件



    SY_FLOWMAP_RECORD_NOT_FOUND = &H10040000        '流程步驟不存在
    SY_FLOWMAP_FIRSTSTEPNO_NOT_FOUND = &H10040001   '流程的第一個步驟不存在


    SYUSER_USER_NOT_FOUND = &H10050000              '使用者不存在

    SYFLOWDEF_DIRECTION_NOT_FOUND = &H10060000      '無法取得流程步驟的PN方向
    SYFLOWDEF_FIRSTSTEP_NOT_FOUND = &H10060001      '無法取得流程步驟的第一個步驟(START)
    SYFLOWDEF_BRADEPNO_NOT_FOUND = &H10060002       '無法取得分行代碼，最有可能的原因：流程步驟未設定關連的角色(未加入角色資料至SY_REL_ROLE_FLOWMAP)

    SYFLOWID_CANNOT_GET_SUBSYSID = &H10070000       '無法從取得SY_FLOW_ID的SUBSYSID內容
    SYFLOWID_CANNOT_GET_FLOWCNAME = &H10070001       '無法從取得SY_FLOW_ID的FLOWCNAME內容

    SYFLOWINCIDENT_CANNOT_GET_MAX_SUBFLOWSEQ = &H10080000   '無法取得SUBFLOW_SEQ的最大值
    SYFLOWINCIDENT_CANNOT_GET_MAX_SUBFLOWCOUNT = &H10080001 '無法取得SUBFLOW_COUNT的最大值
    SYFLOWINCIDENT_SET_WRONG_STATUS = &H10080002            '[SY_FLOWINCIDENT].[STATUS]只能設為1或3
    SYFLOWINCIDENT_CANNOT_GET_MAX_PARENTSEQ = &H10080003    '無法取得PARENT_SEQ的內容

    SYBRANCH_BRADEPNO_NOT_FOUND = &H10100000                '無法取得BRA_DEPNO的內容
    SYBRANCH_CHILDBRADEPNO_NOT_FOUND = &H10100001           '無法取得部門內的所有子部門

    SYRELBRANCHUSER_BRADEPNO_NOT_FOUND = &H10110000         '無法取得使用者的BRA_DEPNO

    SYCONFRATION_NAME_NOT_FOUND = &H10120000            '無法取得設定值

    FLOW_INVALID_USER = &H20000000              '使用者不存在
    FLOW_USER_LEFT = &H20000001                 '使用者已離職
    FLOW_USERINFO_NOT_FOUND = &H20000002        '使用者登入資訊遺失，可能是系統更新元件時重置不同元件內的物件參考或未設定登入使用者資訊
    FLOW_FLOWINFO_NOT_FOUND = &H20000010        '無法取得流程資訊
    FLOW_CASEID_CANNOT_WRITTEN = &H20000020     'caseid無法寫入至SY_CASEID
    FLOW_CASEINFO_NOT_FOUND = &H20000021        '無法取得案件資訊
    FLOW_CASE_CANNOT_RESET = &H20000022         '案件未結束，不可以重新開啟案件
    FLOW_TOO_MANY_EVENTS_FIRED = &H20000030     '流程事件觸發不超過10次
    FLOW_CASE_CANNOT_SEND_BECAUSE_CASE_CLOSED = &H20000031  '案件已結案，無法再送出案件
    FLOW_NEXTSTEP_NOT_FOUND = &H20000100        '無法取得流程步驟的下一步驟
    FLOW_NEXTSTEP_NOT_FOUND_AMBIGUOUS = &H20000101        '無法取得流程步驟的下一步驟，有兩個以上的下一步驟
    FLOW_CLIENT_NOT_FOUND_IN_CURRENT_STEP = &H20000102    '目前步驟沒有可以送件的使用者
    FLOW_CASEREVISION_NOT_FOUND = &H20000103    '無法從SY_CASEREVISION取得資料
    FLOW_NEXTSTEP_CLIENT_NOT_IN_HIS_BRANCH = &H20000103   '下一步驟的USER不在自己的分行內

    FLOW_CLIENT_NOT_FOUND_IN_NEXTSTEP_SAMEPROHIBITED = &H20000200  '無法取得下一步驟的使用者，因為下一步驟的唯一的使用者跟此步驟的使用者相同，所以無法案件無法送出
    FLOW_CLIENT_NOT_FOUND_IN_NEXTSTEP = &H20000201  '無法取得下一步驟的使用者

    FLOW_NO_RIGHTS_TO_PERFORM_NEXTSTEP = &H20000300 '使用者沒有權限執行下一個流程步驟



End Enum