VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "팝빌 홈택스 현금영수증 연계 API SDK Example"
   ClientHeight    =   11430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   11430
   ScaleWidth      =   12390
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame6 
      Caption         =   "홈택스 현금영수증 연계 관련 API"
      Height          =   7935
      Left            =   240
      TabIndex        =   23
      Top             =   3240
      Width           =   11775
      Begin VB.TextBox txtJobID 
         Height          =   330
         Left            =   1920
         TabIndex        =   38
         Top             =   2760
         Width           =   2175
      End
      Begin VB.ListBox cashbillList 
         Height          =   4020
         Left            =   240
         TabIndex        =   36
         Top             =   3480
         Width           =   10935
      End
      Begin VB.Frame Frame9 
         Caption         =   "부가기능"
         Height          =   2415
         Left            =   5280
         TabIndex        =   31
         Top             =   360
         Width           =   2775
         Begin VB.CommandButton btnGetFlatRatePopUpURL 
            Caption         =   "정액제 서비스 신청 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton btnGetFlatRateState 
            Caption         =   "정액제 서비스 상태 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton btnGetCertificatePopUpURL 
            Caption         =   "공인인증서 등록 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   33
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton btnGetCertificateExpireDate 
            Caption         =   "공인인증서 만료일자 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   32
            Top             =   1800
            Width           =   2535
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "매출/매입 수집결과 조회"
         Height          =   1935
         Left            =   2640
         TabIndex        =   28
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnSearch 
            Caption         =   "수집 결과 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnSummary 
            Caption         =   "수집 결과 요약정보 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "매출/매입 내역 수집"
         Height          =   1935
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnRequestJob 
            Caption         =   "수집 요청"
            Height          =   410
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton btnGetJobState 
            Caption         =   "수집 상태 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnListActiveJob 
            Caption         =   "수집 상태 목록 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   1935
         End
      End
      Begin VB.Label Label4 
         Caption         =   "(작업아이디는 '수집 요청' 호출시 생성됩니다.)"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   3120
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "작업아이디(jobID) :"
         Height          =   180
         Left            =   240
         TabIndex        =   37
         Top             =   2835
         Width           =   1695
      End
   End
   Begin VB.CommandButton btnPopbillURL_CHRG 
      Caption         =   " 포인트 충전 URL"
      Height          =   410
      Left            =   7320
      TabIndex        =   6
      Top             =   1755
      Width           =   2055
   End
   Begin VB.CommandButton btnUpdateCorpInfo 
      Caption         =   "회사정보 수정"
      Height          =   410
      Left            =   9720
      TabIndex        =   5
      Top             =   1755
      Width           =   1935
   End
   Begin VB.CommandButton btnUpdateContact 
      Caption         =   "담당자 정보 수정"
      Height          =   410
      Left            =   4920
      TabIndex        =   4
      Top             =   2235
      Width           =   2055
   End
   Begin VB.CommandButton btnListContact 
      Caption         =   "담당자 목록 조회"
      Height          =   410
      Left            =   4920
      TabIndex        =   3
      Top             =   1755
      Width           =   2055
   End
   Begin VB.CommandButton btnCheckID 
      Caption         =   "ID 중복 확인"
      Height          =   410
      Left            =   480
      TabIndex        =   2
      Top             =   1755
      Width           =   1455
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   6120
      TabIndex        =   1
      Text            =   "testkorea"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Text            =   "1234567890"
      Top             =   135
      Width           =   1935
   End
   Begin VB.Frame Frame15 
      Caption         =   "회사정보 관련"
      Height          =   1935
      Left            =   9600
      TabIndex        =   7
      Top             =   915
      Width           =   2240
      Begin VB.CommandButton btnGetCorpInfo 
         Caption         =   "회사정보 조회"
         Height          =   410
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2535
      Left            =   240
      TabIndex        =   9
      Top             =   555
      Width           =   11775
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련"
         Height          =   1935
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   2295
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "담당자 관련"
         Height          =   1935
         Left            =   4560
         TabIndex        =   12
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   410
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보"
         Height          =   1935
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   410
            Left            =   120
            TabIndex        =   19
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         Height          =   1935
         Left            =   6960
         TabIndex        =   10
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
            Caption         =   " 팝빌 로그인 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   2055
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   "팝빌회원 아이디 : "
      Height          =   180
      Left            =   4560
      TabIndex        =   22
      Top             =   195
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "팝빌회원 사업자번호 :"
      Height          =   180
      Left            =   360
      TabIndex        =   21
      Top             =   195
      Width           =   1920
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'링크아이디
Private Const linkID = "TESTER"
'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private htCashbillService As New PBHTCashbillService

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.Message)
End Sub

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.Message)
End Sub

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = htCashbillService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
End Sub

Private Sub btnGetCertificateExpireDate_Click()
    Dim expireDate As String
    
    expireDate = htCashbillService.GetCertificateExpireDate(txtCorpNum.Text)
    
    If expireDate = "" Then
        
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "인증서만료일 : " + expireDate
End Sub

Private Sub btnGetCertificatePopUpURL_Click()
    Dim url As String
    
    url = htCashbillService.GetCertificatePopUpURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    
    Set ChargeInfo = htCashbillService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "unitCost ([정액제]월정액요금) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    
    Set CorpInfo = htCashbillService.GetCorpInfo(txtCorpNum.Text, txtUserID.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ceoname : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetFlatRatePopUpURL_Click()
    Dim url As String
    
    url = htCashbillService.GetFlatRatePopUpURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetFlatRateState_Click()
    Dim flatRateInfo As PBHTCashbillFlatRate
    
    Set flatRateInfo = htCashbillService.GetFlatRateState(txtCorpNum.Text)
     
    If flatRateInfo Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "referencdeID (사업자번호) : " + flatRateInfo.referenceID + vbCrLf
    tmp = tmp + "contractDT (정액제 서비스 시작일시) : " + flatRateInfo.contractDT + vbCrLf
    tmp = tmp + "useEndDate (정액제 서비스 종료일) : " + flatRateInfo.useEndDate + vbCrLf
    tmp = tmp + "baseDate (자동연장 결제일) : " + CStr(flatRateInfo.baseDate) + vbCrLf
    tmp = tmp + "state (정액제 서비스 상태) : " + CStr(flatRateInfo.state) + vbCrLf
    tmp = tmp + "closeRequestYN (서비스 해지신청 여부) : " + CStr(flatRateInfo.closeRequestYN) + vbCrLf
    tmp = tmp + "useRestrictYN (서비스 사용제한 여부) : " + CStr(flatRateInfo.useRestrictYN) + vbCrLf
    tmp = tmp + "closeOnExpired (서비스만료시 해지여부 ) : " + CStr(flatRateInfo.closeOnExpired) + vbCrLf
    tmp = tmp + "unPaidYN (미수금 보유 여부) : " + CStr(flatRateInfo.unPaidYN) + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetJobState_Click()
    Dim jobInfo As PBHTCashbillJobState
    
    Set jobInfo = htCashbillService.GetJobState(txtCorpNum.Text, txtJobID.Text)
     
    If jobInfo Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "jobID (작업아이디) : " + jobInfo.jobID + vbCrLf
    tmp = tmp + "jobState (수집상태) : " + CStr(jobInfo.jobState) + vbCrLf
    tmp = tmp + "queryType (수집유형) : " + jobInfo.queryType + vbCrLf
    tmp = tmp + "queryDateType (일자유형) : " + jobInfo.queryDateType + vbCrLf
    tmp = tmp + "queryStDate (시작일자) : " + jobInfo.queryStDate + vbCrLf
    tmp = tmp + "queryEnDate (종료일자) : " + jobInfo.queryEnDate + vbCrLf
    tmp = tmp + "errorCode (오류코드) : " + CStr(jobInfo.errorCode) + vbCrLf
    tmp = tmp + "errorReason (오류메시지) : " + jobInfo.errorReason + vbCrLf
    tmp = tmp + "jobStartDT (작업 시작일시) : " + jobInfo.jobStartDT + vbCrLf
    tmp = tmp + "jobEndDT (작업 종료일시) : " + jobInfo.jobEndDT + vbCrLf
    tmp = tmp + "collectCount (수집개수) : " + CStr(jobInfo.collectCount) + vbCrLf
    tmp = tmp + "regDT (수집 요청일시) : " + jobInfo.regDT + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = htCashbillService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
End Sub

Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = htCashbillService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    joinData.linkID = linkID '링크 아이디
    joinData.CorpNum = "1231212312" '사업자번호 "-" 제외.
    joinData.ceoname = "대표자성명"
    joinData.corpName = "회원상호"
    joinData.addr = "주소"
    joinData.bizType = "업태"
    joinData.bizClass = "업종"
    joinData.id = "userid"      '6자 이상 20자 미만.
    joinData.pwd = "pwd_must_be_long_enough"    '6자 이상 20자 미만.
    joinData.ContactName = "담당자성명"
    joinData.ContactTEL = "02-999-9999"
    joinData.ContactHP = "010-1234-5678"
    joinData.ContactFAX = "02-999-9998"
    joinData.ContactEmail = "test@test.com"
    
    Set Response = htCashbillService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.Message)
End Sub

Private Sub btnListActiveJob_Click()
    Dim jobList As Collection
        
    Set jobList = htCashbillService.ListActiveJob(txtCorpNum.Text)
     
    If jobList Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "작업아이디(jobID)의 유효시간은 1시간입니다" + vbCrLf + vbCrLf
    tmp = tmp + "jobID | jobState | queryType | queryDateType | queryStDate | queryEnDate | errorCode | errorReason | jobStartDT | jobEndDT | collectCount | regDT " + vbCrLf
    
    Dim info As PBHTCashbillJobState
    
    For Each info In jobList
        tmp = tmp + CStr(info.jobID) + " | "
        tmp = tmp + CStr(info.jobState) + " | "
        tmp = tmp + info.queryType + " | "
        tmp = tmp + info.queryDateType + " | "
        tmp = tmp + info.queryStDate + " | "
        tmp = tmp + info.queryEnDate + " | "
        tmp = tmp + CStr(info.errorCode) + " | "
        tmp = tmp + info.errorReason + " | "
        tmp = tmp + info.jobStartDT + " | "
        tmp = tmp + info.jobEndDT + " | "
        tmp = tmp + CStr(info.collectCount) + " | "
        tmp = tmp + info.regDT
        tmp = tmp + vbCrLf
    Next
    
    MsgBox tmp
    
    If jobList.count > 0 Then
        txtJobID.Text = jobList.Item(1).jobID
    End If
End Sub

Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = htCashbillService.ListContact(txtCorpNum.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT " + vbCrLf
    
    Dim info As PBContactInfo
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnPopbillURL_CHRG_Click()
    Dim url As String
    
    url = htCashbillService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
         MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = "testkorea_20151007"
    
    '담당자 비밀번호
    joinData.pwd = "test@test.com"
    
    '담당자명
    joinData.personName = "담당자명"
    
    '연락처
    joinData.tel = "070-1234-1234"
    
    '휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '이메일 주소
    joinData.email = "test@test.com"
    
    '팩스번호
    joinData.fax = "070-1234-1234"
    
    '전체조회 여부, true-회사조회, false-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 권한여부
    joinData.mgrYN = False
        
    Set Response = htCashbillService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.Message)
End Sub

Private Sub btnRequestJob_Click()
    Dim jobID As String
    Dim SDate As String
    Dim EDate As String
    Dim cbType As KeyType
    
    '현금영수증 유형, SELL-매출, BUY-매입, TURSTEE-위수탁
    cbType = SELL
        
    '시작일자, 표시형식(yyyyMMdd)
    SDate = "20160501"
    
    '종료일자, 표시형식(yyyyMMdd)
    EDate = "20160701"
        
        
    '작업아이디(jobID)의 유효시간은 1시간입니다.
    jobID = htCashbillService.RequestJob(txtCorpNum.Text, cbType, SDate, EDate)
    
    If jobID = "" Then
         MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "jobID(작업아이디) : " + jobID + vbCrLf
    
    
    txtJobID.Text = jobID
End Sub

Private Sub btnSearch_Click()
    Dim SearchList As PBHTCashbillSearch
    Dim cbType As New Collection
    Dim tradeType As New Collection
    Dim tradeUsage As New Collection
    Dim page As Integer
    Dim perPage As Integer
    Dim order As String
    Dim tmp As String
    Dim rowTmp As String
    
    
    '현금영수증 형태 배열, N-일반현금영수증, C-취소현금영수증
    tradeType.Add "N"
    tradeType.Add "C"
        
    '거래용도 배열, P-소득공제용, C-지출증빙용
    tradeUsage.Add "P"
    tradeUsage.Add "C"
        
    '페이지 번호
    page = 1
    
    '페이지당 검색개수, 최대 1000건
    perPage = 10
    
    '정렬 방향, D-내림차순, A-오름차순
    order = "D"
        
        
    'Search 호출
    Set SearchList = htCashbillService.Search(txtCorpNum.Text, txtJobID.Text, tradeType, tradeUsage, page, perPage, order)
    
        
    If SearchList Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (응답코드) : " + CStr(SearchList.code) + vbCrLf
    tmp = tmp + "message (응답메시지) : " + SearchList.Message + vbCrLf
    tmp = tmp + "total (총 검색결과 건수) : " + CStr(SearchList.total) + vbCrLf
    tmp = tmp + "perPage (페이지당 검색개수) : " + CStr(SearchList.perPage) + vbCrLf
    tmp = tmp + "pageNum (페이지 번호) : " + CStr(SearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (페이지 개수) : " + CStr(SearchList.pageCount) + vbCrLf + vbCrLf
    
    cashbillList.Clear
    
    cashbillList.AddItem "구분 | 거래일시 | 식별번호 | 공급가액 | 세액 | 봉사료 | 거래금액 | 문서형태 | 국세청승인번호", 0
    
    Dim cbInfo As PBHTCashbill
           
    For Each cbInfo In SearchList.list
        ' 추가적인 현금영수증 항목은 [홈택스 현금영수증 연계 API 연동매뉴얼 > 4.1.응답전문 구성] 을 참조하시기 바랍니다.'
        rowTmp = ""
        rowTmp = cbInfo.tradeUsage + " | "
        rowTmp = rowTmp + cbInfo.tradeDT + " | "
        rowTmp = rowTmp + cbInfo.identityNum + " | "
        rowTmp = rowTmp + cbInfo.supplyCost + " | "
        rowTmp = rowTmp + cbInfo.tax + " | "
        rowTmp = rowTmp + cbInfo.serviceFee + " | "
        rowTmp = rowTmp + cbInfo.totalAmount + " | "
        rowTmp = rowTmp + cbInfo.tradeType + " | "
        rowTmp = rowTmp + cbInfo.ntsconfirmNum
        
        cashbillList.AddItem rowTmp, cashbillList.ListCount
        
    Next
    
    MsgBox (tmp)
End Sub

Private Sub btnSummary_Click()
    Dim summaryInfo As PBHTCashbillSummary
    Dim cbType As New Collection
    Dim tradeType As New Collection
    Dim tradeUsage As New Collection
    Dim page As Integer
    Dim perPage As Integer
    Dim order As String
    Dim tmp As String
    Dim rowTmp As String
    
    
    '현금영수증 형태 배열, N-일반현금영수증, C-취소현금영수증
    tradeType.Add "N"
    tradeType.Add "C"
        
    '거래용도 배열, P-소득공제용, C-지출증빙용
    tradeUsage.Add "P"
    tradeUsage.Add "C"
                
       
    'Summary 호출
    Set summaryInfo = htCashbillService.Summary(txtCorpNum.Text, txtJobID.Text, tradeType, tradeUsage)
    
        
    If summaryInfo Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "count (수집결과건수) : " + CStr(summaryInfo.count) + vbCrLf
    tmp = tmp + "supplyCostTotal (공급가액 합계) : " + CStr(summaryInfo.supplyCostTotal) + vbCrLf
    tmp = tmp + "taxTotal (세액 합계) : " + CStr(summaryInfo.taxTotal) + vbCrLf
    tmp = tmp + "serviceFeeTotal (봉사료 합계) : " + CStr(summaryInfo.serviceFeeTotal) + vbCrLf
    tmp = tmp + "amountTotal (합계 금액) : " + CStr(summaryInfo.amountTotal) + vbCrLf
           
            
    MsgBox (tmp)
End Sub

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자명
    joinData.personName = "담당자명_수정"
    
    '연락처
    joinData.tel = "070-1234-1234"
    
    '휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '이메일 주소
    joinData.email = "test@test.com"
    
    '팩스번호
    joinData.fax = "070-1234-1234"
    
    '전체조회여부, True-회사조회, False-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 권한여부
    joinData.mgrYN = False
                
    Set Response = htCashbillService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.Message)
End Sub

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자 성명
    CorpInfo.ceoname = "대표자"
    
    '상호명
    CorpInfo.corpName = "상호"
    
    '주소
    CorpInfo.addr = "서울특별시"
    
    '업태
    CorpInfo.bizType = "업태"
    
    '종목
    CorpInfo.bizClass = "종목"
    
    Set Response = htCashbillService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.Message)
End Sub

Private Sub Form_Load()
    '모듈 초기화
    htCashbillService.Initialize linkID, SecretKey
    
    '연동환경 설정값 True(테스트용), False(상업용)
    htCashbillService.IsTest = True
End Sub

