VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "팝빌 홈택스 현금영수증 매입매출 조회 API SDK Example"
   ClientHeight    =   12075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17355
   LinkTopic       =   "Form1"
   ScaleHeight     =   12075
   ScaleWidth      =   17355
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   12600
      TabIndex        =   53
      Top             =   120
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2970
      Left            =   120
      TabIndex        =   27
      Top             =   600
      Width           =   16935
      Begin VB.Frame Frame15 
         Caption         =   "회사정보 관련"
         Height          =   2415
         Left            =   14400
         TabIndex        =   46
         Top             =   360
         Width           =   2240
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         Height          =   2415
         Left            =   12000
         TabIndex        =   41
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   " 팝빌 로그인 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보"
         Height          =   2415
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   410
            Left            =   120
            TabIndex        =   40
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "담당자 관련"
         Height          =   2415
         Left            =   4680
         TabIndex        =   36
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetContactInfo 
            Caption         =   "담당자 정보 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   45
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   44
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   410
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련"
         Height          =   2415
         Left            =   2040
         TabIndex        =   34
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "연동과금 포인트 "
         Height          =   2415
         Left            =   7080
         TabIndex        =   31
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "포인트 사용내역 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   51
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "포인트 결제내역 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   50
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   " 포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "파트너과금 포인트 "
         Height          =   2415
         Left            =   9480
         TabIndex        =   28
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "홈택스 현금영수증 연계 관련 API"
      Height          =   7935
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   16935
      Begin VB.Frame Frame12 
         Caption         =   "홈택스 인증관련 기능"
         Height          =   2415
         Left            =   8400
         TabIndex        =   19
         Top             =   360
         Width           =   5895
         Begin VB.CommandButton btnDeleteDeptUser 
            Caption         =   "부서사용자 등록정보 삭제"
            Height          =   375
            Left            =   3000
            TabIndex        =   26
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton btnCheckLoginDeptUser 
            Caption         =   "부서사용자 로그인 테스트"
            Height          =   410
            Left            =   3000
            TabIndex        =   25
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton btnCheckDeptUser 
            Caption         =   "부서사용자 등록정보 확인"
            Height          =   410
            Left            =   3000
            TabIndex        =   24
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton btnRegistDeptUser 
            Caption         =   "부서사용자 계정등록"
            Height          =   375
            Left            =   3000
            TabIndex        =   23
            Top             =   1800
            Width           =   2535
         End
         Begin VB.CommandButton btnCheckCertValidation 
            Caption         =   "공인인증서 로그인 테스트"
            Height          =   375
            Left            =   240
            TabIndex        =   22
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton btnGetCertificateExpireDate 
            Caption         =   "공인인증서 만료일자 확인"
            Height          =   410
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton btnGetCertificatePopUpURL 
            Caption         =   "홈택스 인증관리 팝업 URL"
            Height          =   410
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.TextBox txtJobID 
         Height          =   330
         Left            =   1920
         TabIndex        =   17
         Top             =   2760
         Width           =   3135
      End
      Begin VB.ListBox cashbillList 
         Height          =   4020
         Left            =   240
         TabIndex        =   15
         Top             =   3480
         Width           =   16335
      End
      Begin VB.Frame Frame9 
         Caption         =   "부가기능"
         Height          =   1935
         Left            =   5280
         TabIndex        =   12
         Top             =   360
         Width           =   2895
         Begin VB.CommandButton btnGetFlatRatePopUpURL 
            Caption         =   "정액제 서비스 신청 URL"
            Height          =   410
            Left            =   160
            TabIndex        =   14
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton btnGetFlatRateState 
            Caption         =   "정액제 서비스 상태 확인"
            Height          =   410
            Left            =   160
            TabIndex        =   13
            Top             =   840
            Width           =   2535
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "매출/매입 수집결과 조회"
         Height          =   1935
         Left            =   2640
         TabIndex        =   9
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnSearch 
            Caption         =   "수집 결과 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnSummary 
            Caption         =   "수집 결과 요약정보 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "매출/매입 내역 수집"
         Height          =   1935
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnRequestJob 
            Caption         =   "수집 요청"
            Height          =   410
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton btnGetJobState 
            Caption         =   "수집 상태 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnListActiveJob 
            Caption         =   "수집 상태 목록 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   1935
         End
      End
      Begin VB.Label Label4 
         Caption         =   "(작업아이디는 '수집 요청' 호출시 생성됩니다.)"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "작업아이디(jobID) :"
         Height          =   180
         Left            =   240
         TabIndex        =   16
         Top             =   2835
         Width           =   1695
      End
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "URL : "
      Height          =   180
      Left            =   11880
      TabIndex        =   52
      Top             =   195
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "팝빌회원 아이디 : "
      Height          =   180
      Left            =   4560
      TabIndex        =   3
      Top             =   195
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "팝빌회원 사업자번호 :"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   195
      Width           =   1920
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' 팝빌 홈택스 현금영수증 조회 API VB 6.0 SDK Example
'
' - 업데이트 일자 : 2022-01-17
' - 연동 기술지원 연락처 : 1600-9854
' - 연동 기술지원 이메일 : code@linkhubcorp.com
' - VB6 SDK 적용방법 안내 : https://docs.popbill.com/htcashbill/tutorial/vb
'
' <테스트 연동개발 준비사항>
' 1) 31, 34번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
' 2) 홈택스 연동서비스를 이용하기 위해 팝빌에 인증정보를 등록 합니다. (인증방법은 부서사용자 인증 / 공동인증서 인증 방식이 있습니다.)
'    - 팝빌로그인 > [홈택스연동] > [환경설정] > [인증 관리] 메뉴에서 [홈택스 부서사용자 등록] 혹은
'      [홈택스 공동인증서 등록]을 통해 인증정보를 등록합니다.
'    - 홈택스연동 인증 관리 팝업 URL(GetCertificatePopUpURL API) 반환된 URL에 접속 하여
'      [홈택스 부서사용자 등록] 혹은 [홈택스 공동인증서 등록]을 통해 인증정보를 등록합니다.

'=========================================================================

Option Explicit

'=========================================================================
' - 인증정보(링크아이디, 비밀키)는 파트너의 연동회원을 식별하는
'   인증에 사용되는 정보로 유출되지 않도록 주의하시기 바랍니다.
' - 상업용 전환이후에도 인증정보(링크아이디, 비밀키)는 변경되지 않습니다.
'=========================================================================

'링크아이디
Private Const linkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'홈택스 현금영수증 연동 클래스 선언
Private htCashbillService As New PBHTCashbillService

'=========================================================================
' 사업자번호를 조회하여 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
' - https://docs.popbill.com/htcashbill/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 사용하고자 하는 아이디의 중복여부를 확인합니다.
' - https://docs.popbill.com/htcashbill/vb/api#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 사용자를 연동회원으로 가입처리합니다.
' - https://docs.popbill.com/htcashbill/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '아이디, 6자이상 50자 미만
    joinData.id = "userid"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf$%^123"
    
    '파트너링크 아이디
    joinData.linkID = linkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1234567890"
    
    '대표자성명, 최대 100자
    joinData.ceoname = "대표자성명"
    
    '상호명, 최대 200자
    joinData.corpName = "회원상호"
    
    '사업장 주소, 최대 300자
    joinData.addr = "주소"
    
    '업태, 최대 100자
    joinData.bizType = "업태"
    
    '종목, 최대 100자
    joinData.bizClass = "종목"

    '담당자 성명, 최대 100자
    joinData.ContactName = "담당자성명"
    
    '담당자 이메일, 최대 100자
    joinData.ContactEmail = "test@test.com"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    Set Response = htCashbillService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 팝빌 홈택스연동(현금) API 서비스 과금정보를 확인합니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = htCashbillService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost ([정액제]월정액요금) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 팝빌 사이트에 로그인 상태로 접근할 수 있는 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
           
    url = htCashbillService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 연동회원 사업자번호에 담당자(팝빌 로그인 계정)를 추가합니다.
' - https://docs.popbill.com/htcashbill/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 50자 미만
    joinData.id = "testkorea"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf$%^123"
    
    '담당자명, 최대 100자
    joinData.personName = "담당자명"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
    
    '담당자 팩스번,최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 메일주소, 최대 100자
    joinData.email = "test@test.com"
    
    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
        
    Set Response = htCashbillService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 확인합니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    '확인할 담당자 아이디
    ContactID = ""
    
    Set info = htCashbillService.GetContactInfo(txtCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 목록을 확인합니다.
' - https://docs.popbill.com/htcashbill/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim info As PBContactInfo
    Dim tmp As String
    
    Set resultList = htCashbillService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 수정합니다.
' - https://docs.popbill.com/htcashbill/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = txtUserID.Text
    
    '담당자 성명, 최대 100자
    joinData.personName = "담당자명_수정"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
        
    '담당자 팩스번호, 최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 이메일, 최대 100자
    joinData.email = "test@test.com"

    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
                
    Set Response = htCashbillService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = htCashbillService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (대표자명) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName (상호) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr (주소) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType (업태) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass (종목) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다
' - https://docs.popbill.com/htcashbill/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명, 최대 100자
    CorpInfo.ceoname = "대표자"
    
    '상호, 최대 200자
    CorpInfo.corpName = "상호"
    
    '주소, 최대 300자
    CorpInfo.addr = "서울특별시"
    
    '업태, 최대 100자
    CorpInfo.bizType = "업태"
    
    '종목, 최대 100자
    CorpInfo.bizClass = "종목"
    
    Set Response = htCashbillService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)를 통해 확인하시기 바랍니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = htCashbillService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "연동회원 잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 연동회원 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
           
    url = htCashbillService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 연동회원 포인트 결제내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim url As String
           
    url = htCashbillService.GetPaymentURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 연동회원 포인트 사용내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim url As String
           
    url = htCashbillService.GetUseHistoryURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를 이용하시기 바랍니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = htCashbillService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "파트너 잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 파트너 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
           
    url = htCashbillService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 홈택스에 신고된 현금영수증 매입/매출 내역 수집을 팝빌에 요청합니다. (조회기간 단위 : 최대 3개월)
' - https://docs.popbill.com/htcashbill/vb/api#RequestJob
'=========================================================================
Private Sub btnRequestJob_Click()
    Dim jobID As String
    Dim SDate As String
    Dim EDate As String
    Dim cbType As KeyType
    
    '현금영수증 유형, SELL-매출, BUY-매입
    cbType = BUY
        
    '시작일자, 표시형식(yyyyMMdd)
    SDate = "20220101"
    
    '종료일자, 표시형식(yyyyMMdd)
    EDate = "20220130"
    
    jobID = htCashbillService.RequestJob(txtCorpNum.Text, cbType, SDate, EDate)
    
    If jobID = "" Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "jobID(작업아이디) : " + jobID + vbCrLf
        
    txtJobID.Text = jobID
End Sub

'=========================================================================
' 함수 RequestJob(수집 요청)를 통해 반환 받은 작업 아이디의 상태를 확인합니다.
' - 거래 내역 조회(Search API) 함수 또는 거래 요약 정보 조회(Summary API) 함수전
'   수집 작업의 진행 상태, 수집 작업의 성공 여부를 확인해야 합니다.
' - 작업 상태(jobState) = 3(완료)이고 수집 결과 코드(errorCode) = 1(수집성공)이면
'   거래 내역 조회(Search) 또는 거래 요약 정보 조회(Summary) 를 해야합니다.
' - 작업 상태(jobState)가 3(완료)이지만 수집 결과 코드(errorCode)가 1(수집성공)이 아닌 경우에는
'   오류메시지(errorReason)로 수집 실패에 대한 원인을 파악할 수 있습니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetJobState
'=========================================================================
Private Sub btnGetJobState_Click()
    Dim jobInfo As PBHTCashbillJobState
    Dim tmp As String
    
    Set jobInfo = htCashbillService.GetJobState(txtCorpNum.Text, txtJobID.Text)
     
    If jobInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
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

'=========================================================================
' 현금영수증 매입/매출 내역 수집요청에 대한 상태 목록을 확인합니다.
' - 수집 요청 후 1시간이 경과한 수집 요청건은 상태정보가 반환되지 않습니다.
' - https://docs.popbill.com/htcashbill/vb/api#ListActiveJob
'=========================================================================
Private Sub btnListActiveJob_Click()
    Dim jobList As Collection
    Dim tmp As String
    Dim info As PBHTCashbillJobState
    
    Set jobList = htCashbillService.ListActiveJob(txtCorpNum.Text)
     
    If jobList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "작업아이디(jobID)의 유효시간은 1시간입니다" + vbCrLf + vbCrLf
    tmp = tmp + "jobID(작업아이디) | jobState(수집상태) | queryType(수집유형) | queryDateType(일자유형) | queryStDate(시작일자) | queryEnDate(종료일자) |" _
            + "errorCode(오류코드) | errorReason(오류메시지) | jobStartDT(작업 시작일시) | jobEndDT(작업 종료일시) | collectCount(수집개수) | regDT(수집 요청일시) " + vbCrLf
    
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

'=========================================================================
' 함수 GetJobState(수집 상태 확인)를 통해 상태 정보 확인된 작업아이디를 활용하여 현금영수증 매입/매출 내역을 조회합니다.
' - https://docs.popbill.com/htcashbill/vb/api#Search
'=========================================================================
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

    Set SearchList = htCashbillService.Search(txtCorpNum.Text, txtJobID.Text, tradeType, tradeUsage, page, perPage, order)
    
    If SearchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (응답코드) : " + CStr(SearchList.code) + vbCrLf
    tmp = tmp + "message (응답메시지) : " + SearchList.Message + vbCrLf
    tmp = tmp + "total (총 검색결과 건수) : " + CStr(SearchList.total) + vbCrLf
    tmp = tmp + "perPage (페이지당 검색개수) : " + CStr(SearchList.perPage) + vbCrLf
    tmp = tmp + "pageNum (페이지 번호) : " + CStr(SearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (페이지 개수) : " + CStr(SearchList.pageCount) + vbCrLf + vbCrLf
    
    cashbillList.Clear
    
    cashbillList.AddItem "ntsconfirmNum (국세청승인번호) | tradeDate (거래일자) | tradeDT (거래일시) | tradeType (문서형태) | tradeUsage (거래구분) | totalAmount (거래금액)", 0
    cashbillList.AddItem "supplyCost (공급가액) | tax (부가세) | serviceFee (봉사료) | invoiceType (매입/매출) | franchiseCorpNum (발행자 사업자번호) ", 1
    cashbillList.AddItem "franchiseCorpName (발행자 상호) | franchiseCorpType (발행자 사업자유형) | identityNum (거래처 식별번호) | identityNumType (식별번호유형)", 2
    cashbillList.AddItem "customerName (고객명) | cardOwnerName (카드소유자명) | deductionType (공제유형)", 3

    Dim cbInfo As PBHTCashbill
           
    For Each cbInfo In SearchList.list
        rowTmp = ""
        rowTmp = cbInfo.ntsconfirmNum + " | "
        rowTmp = rowTmp + cbInfo.tradeDate + " | "
        rowTmp = rowTmp + cbInfo.tradeDT + " | "
        rowTmp = rowTmp + cbInfo.tradeType + " | "
        rowTmp = rowTmp + cbInfo.tradeUsage + " | "
        rowTmp = rowTmp + cbInfo.totalAmount + " | "
        rowTmp = rowTmp + cbInfo.supplyCost + " | "
        rowTmp = rowTmp + cbInfo.tax + " | "
        rowTmp = rowTmp + cbInfo.serviceFee + " | "
        rowTmp = rowTmp + cbInfo.invoiceType + " | "
        rowTmp = rowTmp + cbInfo.franchiseCorpNum + " | "
        rowTmp = rowTmp + cbInfo.franchiseCorpName + " | "
        rowTmp = rowTmp + cbInfo.franchiseCorpType + " | "
        rowTmp = rowTmp + cbInfo.identityNum + " | "
        rowTmp = rowTmp + cbInfo.franchiseCorpName + " | "
        rowTmp = rowTmp + cbInfo.customerName + " | "
        rowTmp = rowTmp + cbInfo.cardOwnerName + " | "
        rowTmp = rowTmp + cbInfo.deductionType
        cashbillList.AddItem rowTmp, cashbillList.ListCount
    Next
               
    MsgBox (tmp)
End Sub

'=========================================================================
' 함수 GetJobState(수집 상태 확인)를 통해 상태 정보가 확인된 작업아이디를 활용하여 수집된 현금영수증 매입/매출 내역의 요약 정보를 조회합니다.
' - https://docs.popbill.com/htcashbill/vb/api#Summary
'=========================================================================
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
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "count (수집결과건수) : " + CStr(summaryInfo.count) + vbCrLf
    tmp = tmp + "supplyCostTotal (공급가액 합계) : " + CStr(summaryInfo.supplyCostTotal) + vbCrLf
    tmp = tmp + "taxTotal (세액 합계) : " + CStr(summaryInfo.taxTotal) + vbCrLf
    tmp = tmp + "serviceFeeTotal (봉사료 합계) : " + CStr(summaryInfo.serviceFeeTotal) + vbCrLf
    tmp = tmp + "amountTotal (합계 금액) : " + CStr(summaryInfo.amountTotal) + vbCrLf
           
    MsgBox (tmp)
End Sub

'=========================================================================
' 홈택스연동 정액제 서비스 신청 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetFlatRatePopUpURL
'=========================================================================
Private Sub btnGetFlatRatePopUpURL_Click()
    Dim url As String
    
    url = htCashbillService.GetFlatRatePopUpURL(txtCorpNum.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 홈택스연동 정액제 서비스 상태를 확인합니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetFlatRateState
'=========================================================================
Private Sub btnGetFlatRateState_Click()
    Dim flatRateInfo As PBHTCashbillFlatRate
    Dim tmp As String
    
    Set flatRateInfo = htCashbillService.GetFlatRateState(txtCorpNum.Text)
     
    If flatRateInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
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

'=========================================================================
' 홈택스연동 인증정보를 관리하는 페이지의 팝업 URL을 반환합니다.
' - 인증방식에는 부서사용자/공동인증서 인증 방식이 있습니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetCertificatePopUpURL
'=========================================================================
Private Sub btnGetCertificatePopUpURL_Click()
    Dim url As String
    
    url = htCashbillService.GetCertificatePopUpURL(txtCorpNum.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 홈택스연동 인증을 위해 팝빌에 등록된 인증서 만료일자를 확인합니다.
' - https://docs.popbill.com/htcashbill/vb/api#GetCertificateExpireDate
'=========================================================================
Private Sub btnGetCertificateExpireDate_Click()
    Dim expireDate As String
    
    expireDate = htCashbillService.GetCertificateExpireDate(txtCorpNum.Text)
    
    If expireDate = "" Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "인증서만료일 : " + expireDate
End Sub

'=========================================================================
' 팝빌에 등록된 인증서로 홈택스 로그인 가능 여부를 확인합니다.
' - https://docs.popbill.com/htcashbill/vb/api#CheckCertValidation
'=========================================================================
Private Sub btnCheckCertValidation_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.CheckCertValidation(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 홈택스연동 인증을 위해 팝빌에 등록된 현금영수증 자료조회 부서사용자 계정을 확인합니다.
' - https://docs.popbill.com/htcashbill/vb/api#CheckDeptUser
'=========================================================================
Private Sub btnCheckDeptUser_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.CheckDeptUser(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 팝빌에 등록된 현금영수증 자료조회 부서사용자 계정 정보로 홈택스 로그인 가능 여부를 확인합니다.
' - https://docs.popbill.com/htcashbill/vb/api#CheckLoginDeptUser
'=========================================================================
Private Sub btnCheckLoginDeptUser_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.CheckLoginDeptUser(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 팝빌에 등록된 홈택스 현금영수증 자료조회 부서사용자 계정을 삭제합니다.
' - https://docs.popbill.com/htcashbill/vb/api#DeleteDeptUser
'=========================================================================
Private Sub btnDeleteDeptUser_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.DeleteDeptUser(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

'=========================================================================
' 홈택스연동 인증을 위해 팝빌에 현금영수증 자료조회 부서사용자 계정을 등록합니다.
' - https://docs.popbill.com/htcashbill/vb/api#RegistDeptUser
'=========================================================================
Private Sub btnRegistDeptUser_Click()
    Dim Response As PBResponse
    Dim DeptUserID As String
    Dim DeptUserPWD As String
    
    '홈택스에서 생성한 현금영수증 부서사용자 아이디
    DeptUserID = "userid_test"
    
    '홈택스에서 생성한 현금영수증 부서사용자 비밀번호
    DeptUserPWD = "passwd_test"
    
    Set Response = htCashbillService.RegistDeptUser(txtCorpNum.Text, DeptUserID, DeptUserPWD)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "응답메시지 : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.Message)
End Sub

Private Sub Form_Load()

    '모듈 초기화
    htCashbillService.Initialize linkID, SecretKey
    
    '연동환경 설정값 True(테스트용), False(상업용)
    htCashbillService.IsTest = True
    
    '인증토큰 IP제한기능 사용여부, True-사용, False-미사용, 기본값(True)
    htCashbillService.IPRestrictOnOff = True
    
    '로컬시스템 시간 사용여부 True-사용, Fasle-미사용, 기본값(False)
    htCashbillService.UseLocalTimeYN = False
    
End Sub

