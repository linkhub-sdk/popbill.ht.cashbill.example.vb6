VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� Ȩ�ý� ���ݿ����� ���� API SDK Example"
   ClientHeight    =   11430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   ScaleHeight     =   11430
   ScaleWidth      =   12390
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame6 
      Caption         =   "Ȩ�ý� ���ݿ����� ���� ���� API"
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
         Caption         =   "�ΰ����"
         Height          =   2415
         Left            =   5280
         TabIndex        =   31
         Top             =   360
         Width           =   2775
         Begin VB.CommandButton btnGetFlatRatePopUpURL 
            Caption         =   "������ ���� ��û URL"
            Height          =   410
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton btnGetFlatRateState 
            Caption         =   "������ ���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   34
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton btnGetCertificatePopUpURL 
            Caption         =   "���������� ��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   33
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton btnGetCertificateExpireDate 
            Caption         =   "���������� �������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   32
            Top             =   1800
            Width           =   2535
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "����/���� ������� ��ȸ"
         Height          =   1935
         Left            =   2640
         TabIndex        =   28
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnSearch 
            Caption         =   "���� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnSummary 
            Caption         =   "���� ��� ������� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "����/���� ���� ����"
         Height          =   1935
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnRequestJob 
            Caption         =   "���� ��û"
            Height          =   410
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton btnGetJobState 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnListActiveJob 
            Caption         =   "���� ���� ��� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   1935
         End
      End
      Begin VB.Label Label4 
         Caption         =   "(�۾����̵�� '���� ��û' ȣ��� �����˴ϴ�.)"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   3120
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "�۾����̵�(jobID) :"
         Height          =   180
         Left            =   240
         TabIndex        =   37
         Top             =   2835
         Width           =   1695
      End
   End
   Begin VB.CommandButton btnPopbillURL_CHRG 
      Caption         =   " ����Ʈ ���� URL"
      Height          =   410
      Left            =   7320
      TabIndex        =   6
      Top             =   1755
      Width           =   2055
   End
   Begin VB.CommandButton btnUpdateCorpInfo 
      Caption         =   "ȸ������ ����"
      Height          =   410
      Left            =   9720
      TabIndex        =   5
      Top             =   1755
      Width           =   1935
   End
   Begin VB.CommandButton btnUpdateContact 
      Caption         =   "����� ���� ����"
      Height          =   410
      Left            =   4920
      TabIndex        =   4
      Top             =   2235
      Width           =   2055
   End
   Begin VB.CommandButton btnListContact 
      Caption         =   "����� ��� ��ȸ"
      Height          =   410
      Left            =   4920
      TabIndex        =   3
      Top             =   1755
      Width           =   2055
   End
   Begin VB.CommandButton btnCheckID 
      Caption         =   "ID �ߺ� Ȯ��"
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
      Caption         =   "ȸ������ ����"
      Height          =   1935
      Left            =   9600
      TabIndex        =   7
      Top             =   915
      Width           =   2240
      Begin VB.CommandButton btnGetCorpInfo 
         Caption         =   "ȸ������ ��ȸ"
         Height          =   410
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2535
      Left            =   240
      TabIndex        =   9
      Top             =   555
      Width           =   11775
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ����"
         Height          =   1935
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   2295
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "����� ����"
         Height          =   1935
         Left            =   4560
         TabIndex        =   12
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   1935
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   19
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         Height          =   1935
         Left            =   6960
         TabIndex        =   10
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
            Caption         =   " �˺� �α��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   2055
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4560
      TabIndex        =   22
      Top             =   195
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ :"
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

'��ũ���̵�
Private Const linkID = "TESTER"
'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
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
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

Private Sub btnGetCertificateExpireDate_Click()
    Dim expireDate As String
    
    expireDate = htCashbillService.GetCertificateExpireDate(txtCorpNum.Text)
    
    If expireDate = "" Then
        
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������������ : " + expireDate
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
    
    tmp = tmp + "unitCost ([������]�����׿��) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
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
    
    tmp = tmp + "referencdeID (����ڹ�ȣ) : " + flatRateInfo.referenceID + vbCrLf
    tmp = tmp + "contractDT (������ ���� �����Ͻ�) : " + flatRateInfo.contractDT + vbCrLf
    tmp = tmp + "useEndDate (������ ���� ������) : " + flatRateInfo.useEndDate + vbCrLf
    tmp = tmp + "baseDate (�ڵ����� ������) : " + CStr(flatRateInfo.baseDate) + vbCrLf
    tmp = tmp + "state (������ ���� ����) : " + CStr(flatRateInfo.state) + vbCrLf
    tmp = tmp + "closeRequestYN (���� ������û ����) : " + CStr(flatRateInfo.closeRequestYN) + vbCrLf
    tmp = tmp + "useRestrictYN (���� ������� ����) : " + CStr(flatRateInfo.useRestrictYN) + vbCrLf
    tmp = tmp + "closeOnExpired (���񽺸���� �������� ) : " + CStr(flatRateInfo.closeOnExpired) + vbCrLf
    tmp = tmp + "unPaidYN (�̼��� ���� ����) : " + CStr(flatRateInfo.unPaidYN) + vbCrLf
    
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
    
    tmp = tmp + "jobID (�۾����̵�) : " + jobInfo.jobID + vbCrLf
    tmp = tmp + "jobState (��������) : " + CStr(jobInfo.jobState) + vbCrLf
    tmp = tmp + "queryType (��������) : " + jobInfo.queryType + vbCrLf
    tmp = tmp + "queryDateType (��������) : " + jobInfo.queryDateType + vbCrLf
    tmp = tmp + "queryStDate (��������) : " + jobInfo.queryStDate + vbCrLf
    tmp = tmp + "queryEnDate (��������) : " + jobInfo.queryEnDate + vbCrLf
    tmp = tmp + "errorCode (�����ڵ�) : " + CStr(jobInfo.errorCode) + vbCrLf
    tmp = tmp + "errorReason (�����޽���) : " + jobInfo.errorReason + vbCrLf
    tmp = tmp + "jobStartDT (�۾� �����Ͻ�) : " + jobInfo.jobStartDT + vbCrLf
    tmp = tmp + "jobEndDT (�۾� �����Ͻ�) : " + jobInfo.jobEndDT + vbCrLf
    tmp = tmp + "collectCount (��������) : " + CStr(jobInfo.collectCount) + vbCrLf
    tmp = tmp + "regDT (���� ��û�Ͻ�) : " + jobInfo.regDT + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = htCashbillService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
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
    
    joinData.linkID = linkID '��ũ ���̵�
    joinData.CorpNum = "1231212312" '����ڹ�ȣ "-" ����.
    joinData.ceoname = "��ǥ�ڼ���"
    joinData.corpName = "ȸ����ȣ"
    joinData.addr = "�ּ�"
    joinData.bizType = "����"
    joinData.bizClass = "����"
    joinData.id = "userid"      '6�� �̻� 20�� �̸�.
    joinData.pwd = "pwd_must_be_long_enough"    '6�� �̻� 20�� �̸�.
    joinData.ContactName = "����ڼ���"
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
    
    tmp = tmp + "�۾����̵�(jobID)�� ��ȿ�ð��� 1�ð��Դϴ�" + vbCrLf + vbCrLf
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
    
    '����� ���̵�
    joinData.id = "testkorea_20151007"
    
    '����� ��й�ȣ
    joinData.pwd = "test@test.com"
    
    '����ڸ�
    joinData.personName = "����ڸ�"
    
    '����ó
    joinData.tel = "070-1234-1234"
    
    '�޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '�̸��� �ּ�
    joinData.email = "test@test.com"
    
    '�ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    '��ü��ȸ ����, true-ȸ����ȸ, false-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
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
    
    '���ݿ����� ����, SELL-����, BUY-����, TURSTEE-����Ź
    cbType = SELL
        
    '��������, ǥ������(yyyyMMdd)
    SDate = "20160501"
    
    '��������, ǥ������(yyyyMMdd)
    EDate = "20160701"
        
        
    '�۾����̵�(jobID)�� ��ȿ�ð��� 1�ð��Դϴ�.
    jobID = htCashbillService.RequestJob(txtCorpNum.Text, cbType, SDate, EDate)
    
    If jobID = "" Then
         MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "jobID(�۾����̵�) : " + jobID + vbCrLf
    
    
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
    
    
    '���ݿ����� ���� �迭, N-�Ϲ����ݿ�����, C-������ݿ�����
    tradeType.Add "N"
    tradeType.Add "C"
        
    '�ŷ��뵵 �迭, P-�ҵ������, C-����������
    tradeUsage.Add "P"
    tradeUsage.Add "C"
        
    '������ ��ȣ
    page = 1
    
    '�������� �˻�����, �ִ� 1000��
    perPage = 10
    
    '���� ����, D-��������, A-��������
    order = "D"
        
        
    'Search ȣ��
    Set SearchList = htCashbillService.Search(txtCorpNum.Text, txtJobID.Text, tradeType, tradeUsage, page, perPage, order)
    
        
    If SearchList Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (�����ڵ�) : " + CStr(SearchList.code) + vbCrLf
    tmp = tmp + "message (����޽���) : " + SearchList.Message + vbCrLf
    tmp = tmp + "total (�� �˻���� �Ǽ�) : " + CStr(SearchList.total) + vbCrLf
    tmp = tmp + "perPage (�������� �˻�����) : " + CStr(SearchList.perPage) + vbCrLf
    tmp = tmp + "pageNum (������ ��ȣ) : " + CStr(SearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (������ ����) : " + CStr(SearchList.pageCount) + vbCrLf + vbCrLf
    
    cashbillList.Clear
    
    cashbillList.AddItem "���� | �ŷ��Ͻ� | �ĺ���ȣ | ���ް��� | ���� | ����� | �ŷ��ݾ� | �������� | ����û���ι�ȣ", 0
    
    Dim cbInfo As PBHTCashbill
           
    For Each cbInfo In SearchList.list
        ' �߰����� ���ݿ����� �׸��� [Ȩ�ý� ���ݿ����� ���� API �����Ŵ��� > 4.1.�������� ����] �� �����Ͻñ� �ٶ��ϴ�.'
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
    
    
    '���ݿ����� ���� �迭, N-�Ϲ����ݿ�����, C-������ݿ�����
    tradeType.Add "N"
    tradeType.Add "C"
        
    '�ŷ��뵵 �迭, P-�ҵ������, C-����������
    tradeUsage.Add "P"
    tradeUsage.Add "C"
                
       
    'Summary ȣ��
    Set summaryInfo = htCashbillService.Summary(txtCorpNum.Text, txtJobID.Text, tradeType, tradeUsage)
    
        
    If summaryInfo Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "count (��������Ǽ�) : " + CStr(summaryInfo.count) + vbCrLf
    tmp = tmp + "supplyCostTotal (���ް��� �հ�) : " + CStr(summaryInfo.supplyCostTotal) + vbCrLf
    tmp = tmp + "taxTotal (���� �հ�) : " + CStr(summaryInfo.taxTotal) + vbCrLf
    tmp = tmp + "serviceFeeTotal (����� �հ�) : " + CStr(summaryInfo.serviceFeeTotal) + vbCrLf
    tmp = tmp + "amountTotal (�հ� �ݾ�) : " + CStr(summaryInfo.amountTotal) + vbCrLf
           
            
    MsgBox (tmp)
End Sub

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����ڸ�
    joinData.personName = "����ڸ�_����"
    
    '����ó
    joinData.tel = "070-1234-1234"
    
    '�޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '�̸��� �ּ�
    joinData.email = "test@test.com"
    
    '�ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    '��ü��ȸ����, True-ȸ����ȸ, False-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
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
    
    '��ǥ�� ����
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ��
    CorpInfo.corpName = "��ȣ"
    
    '�ּ�
    CorpInfo.addr = "����Ư����"
    
    '����
    CorpInfo.bizType = "����"
    
    '����
    CorpInfo.bizClass = "����"
    
    Set Response = htCashbillService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(htCashbillService.LastErrCode) + "] " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.Message)
End Sub

Private Sub Form_Load()
    '��� �ʱ�ȭ
    htCashbillService.Initialize linkID, SecretKey
    
    '����ȯ�� ������ True(�׽�Ʈ��), False(�����)
    htCashbillService.IsTest = True
End Sub

