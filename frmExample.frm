VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� Ȩ�ý� ���ݿ����� ���Ը��� ��ȸ API SDK Example"
   ClientHeight    =   12075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17355
   LinkTopic       =   "Form1"
   ScaleHeight     =   12075
   ScaleWidth      =   17355
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   12600
      TabIndex        =   53
      Top             =   120
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2970
      Left            =   120
      TabIndex        =   27
      Top             =   600
      Width           =   16935
      Begin VB.Frame Frame15 
         Caption         =   "ȸ������ ����"
         Height          =   2415
         Left            =   14400
         TabIndex        =   46
         Top             =   360
         Width           =   2240
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "ȸ������ ����"
            Height          =   410
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "ȸ������ ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         Height          =   2415
         Left            =   12000
         TabIndex        =   41
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   " �˺� �α��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   2415
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID �ߺ� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   40
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "����� ����"
         Height          =   2415
         Left            =   4680
         TabIndex        =   36
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetContactInfo 
            Caption         =   "����� ���� Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "����� ���� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   45
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "����� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   44
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ����"
         Height          =   2415
         Left            =   2040
         TabIndex        =   34
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "�������� ����Ʈ "
         Height          =   2415
         Left            =   7080
         TabIndex        =   31
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "����Ʈ ��볻�� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   51
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "����Ʈ �������� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   50
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   " ����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ "
         Height          =   2415
         Left            =   9480
         TabIndex        =   28
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   2055
         End
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   2055
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Ȩ�ý� ���ݿ����� ���� ���� API"
      Height          =   7935
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   16935
      Begin VB.Frame Frame12 
         Caption         =   "Ȩ�ý� �������� ���"
         Height          =   2415
         Left            =   8400
         TabIndex        =   19
         Top             =   360
         Width           =   5895
         Begin VB.CommandButton btnDeleteDeptUser 
            Caption         =   "�μ������ ������� ����"
            Height          =   375
            Left            =   3000
            TabIndex        =   26
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton btnCheckLoginDeptUser 
            Caption         =   "�μ������ �α��� �׽�Ʈ"
            Height          =   410
            Left            =   3000
            TabIndex        =   25
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton btnCheckDeptUser 
            Caption         =   "�μ������ ������� Ȯ��"
            Height          =   410
            Left            =   3000
            TabIndex        =   24
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton btnRegistDeptUser 
            Caption         =   "�μ������ �������"
            Height          =   375
            Left            =   3000
            TabIndex        =   23
            Top             =   1800
            Width           =   2535
         End
         Begin VB.CommandButton btnCheckCertValidation 
            Caption         =   "���������� �α��� �׽�Ʈ"
            Height          =   375
            Left            =   240
            TabIndex        =   22
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton btnGetCertificateExpireDate 
            Caption         =   "���������� �������� Ȯ��"
            Height          =   410
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton btnGetCertificatePopUpURL 
            Caption         =   "Ȩ�ý� �������� �˾� URL"
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
         Caption         =   "�ΰ����"
         Height          =   1935
         Left            =   5280
         TabIndex        =   12
         Top             =   360
         Width           =   2895
         Begin VB.CommandButton btnGetFlatRatePopUpURL 
            Caption         =   "������ ���� ��û URL"
            Height          =   410
            Left            =   160
            TabIndex        =   14
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton btnGetFlatRateState 
            Caption         =   "������ ���� ���� Ȯ��"
            Height          =   410
            Left            =   160
            TabIndex        =   13
            Top             =   840
            Width           =   2535
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "����/���� ������� ��ȸ"
         Height          =   1935
         Left            =   2640
         TabIndex        =   9
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnSearch 
            Caption         =   "���� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnSummary 
            Caption         =   "���� ��� ������� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "����/���� ���� ����"
         Height          =   1935
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnRequestJob 
            Caption         =   "���� ��û"
            Height          =   410
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton btnGetJobState 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnListActiveJob 
            Caption         =   "���� ���� ��� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   1935
         End
      End
      Begin VB.Label Label4 
         Caption         =   "(�۾����̵�� '���� ��û' ȣ��� �����˴ϴ�.)"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "�۾����̵�(jobID) :"
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
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4560
      TabIndex        =   3
      Top             =   195
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ :"
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
' �˺� Ȩ�ý� ���ݿ����� ��ȸ API VB 6.0 SDK Example
'
' - ������Ʈ ���� : 2022-01-17
' - ���� ������� ����ó : 1600-9854
' - ���� ������� �̸��� : code@linkhubcorp.com
' - VB6 SDK ������ �ȳ� : https://docs.popbill.com/htcashbill/tutorial/vb
'
' <�׽�Ʈ �������� �غ����>
' 1) 31, 34�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
' 2) Ȩ�ý� �������񽺸� �̿��ϱ� ���� �˺��� ���������� ��� �մϴ�. (��������� �μ������ ���� / ���������� ���� ����� �ֽ��ϴ�.)
'    - �˺��α��� > [Ȩ�ý�����] > [ȯ�漳��] > [���� ����] �޴����� [Ȩ�ý� �μ������ ���] Ȥ��
'      [Ȩ�ý� ���������� ���]�� ���� ���������� ����մϴ�.
'    - Ȩ�ý����� ���� ���� �˾� URL(GetCertificatePopUpURL API) ��ȯ�� URL�� ���� �Ͽ�
'      [Ȩ�ý� �μ������ ���] Ȥ�� [Ȩ�ý� ���������� ���]�� ���� ���������� ����մϴ�.

'=========================================================================

Option Explicit

'=========================================================================
' - ��������(��ũ���̵�, ���Ű)�� ��Ʈ���� ����ȸ���� �ĺ��ϴ�
'   ������ ���Ǵ� ������ ������� �ʵ��� �����Ͻñ� �ٶ��ϴ�.
' - ����� ��ȯ���Ŀ��� ��������(��ũ���̵�, ���Ű)�� ������� �ʽ��ϴ�.
'=========================================================================

'��ũ���̵�
Private Const linkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'Ȩ�ý� ���ݿ����� ���� Ŭ���� ����
Private htCashbillService As New PBHTCashbillService

'=========================================================================
' ����ڹ�ȣ�� ��ȸ�Ͽ� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ϰ��� �ϴ� ���̵��� �ߺ����θ� Ȯ���մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ڸ� ����ȸ������ ����ó���մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '���̵�, 6���̻� 50�� �̸�
    joinData.id = "userid"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf$%^123"
    
    '��Ʈ�ʸ�ũ ���̵�
    joinData.linkID = linkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1234567890"
    
    '��ǥ�ڼ���, �ִ� 100��
    joinData.ceoname = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 200��
    joinData.corpName = "ȸ����ȣ"
    
    '����� �ּ�, �ִ� 300��
    joinData.addr = "�ּ�"
    
    '����, �ִ� 100��
    joinData.bizType = "����"
    
    '����, �ִ� 100��
    joinData.bizClass = "����"

    '����� ����, �ִ� 100��
    joinData.ContactName = "����ڼ���"
    
    '����� �̸���, �ִ� 100��
    joinData.ContactEmail = "test@test.com"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    Set Response = htCashbillService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' �˺� Ȩ�ý�����(����) API ���� ���������� Ȯ���մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = htCashbillService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost ([������]�����׿��) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �˺� ����Ʈ�� �α��� ���·� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
           
    url = htCashbillService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� �����(�˺� �α��� ����)�� �߰��մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 50�� �̸�
    joinData.id = "testkorea"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf$%^123"
    
    '����ڸ�, �ִ� 100��
    joinData.personName = "����ڸ�"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
    
    '����� �ѽ���,�ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �����ּ�, �ִ� 100��
    joinData.email = "test@test.com"
    
    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
        
    Set Response = htCashbillService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ Ȯ���մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    'Ȯ���� ����� ���̵�
    ContactID = ""
    
    Set info = htCashbillService.GetContactInfo(txtCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | hp(�޴�����ȣ) |  fax(�ѽ���ȣ) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ����� Ȯ���մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim info As PBContactInfo
    Dim tmp As String
    
    Set resultList = htCashbillService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | hp(�޴�����ȣ) |  fax(�ѽ���ȣ) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ �����մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
    
    '����� ����, �ִ� 100��
    joinData.personName = "����ڸ�_����"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
        
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �̸���, �ִ� 100��
    joinData.email = "test@test.com"

    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
                
    Set Response = htCashbillService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = htCashbillService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (��ǥ�ڸ�) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName (��ȣ) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr (�ּ�) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType (����) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass (����) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
' - https://docs.popbill.com/htcashbill/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�, �ִ� 100��
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ, �ִ� 200��
    CorpInfo.corpName = "��ȣ"
    
    '�ּ�, �ִ� 300��
    CorpInfo.addr = "����Ư����"
    
    '����, �ִ� 100��
    CorpInfo.bizType = "����"
    
    '����, �ִ� 100��
    CorpInfo.bizClass = "����"
    
    Set Response = htCashbillService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)�� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = htCashbillService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "����ȸ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
           
    url = htCashbillService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ �������� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim url As String
           
    url = htCashbillService.GetPaymentURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ��볻�� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim url As String
           
    url = htCashbillService.GetUseHistoryURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)�� �̿��Ͻñ� �ٶ��ϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = htCashbillService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��Ʈ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
           
    url = htCashbillService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' Ȩ�ý��� �Ű�� ���ݿ����� ����/���� ���� ������ �˺��� ��û�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 3����)
' - https://docs.popbill.com/htcashbill/vb/api#RequestJob
'=========================================================================
Private Sub btnRequestJob_Click()
    Dim jobID As String
    Dim SDate As String
    Dim EDate As String
    Dim cbType As KeyType
    
    '���ݿ����� ����, SELL-����, BUY-����
    cbType = BUY
        
    '��������, ǥ������(yyyyMMdd)
    SDate = "20220101"
    
    '��������, ǥ������(yyyyMMdd)
    EDate = "20220130"
    
    jobID = htCashbillService.RequestJob(txtCorpNum.Text, cbType, SDate, EDate)
    
    If jobID = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "jobID(�۾����̵�) : " + jobID + vbCrLf
        
    txtJobID.Text = jobID
End Sub

'=========================================================================
' �Լ� RequestJob(���� ��û)�� ���� ��ȯ ���� �۾� ���̵��� ���¸� Ȯ���մϴ�.
' - �ŷ� ���� ��ȸ(Search API) �Լ� �Ǵ� �ŷ� ��� ���� ��ȸ(Summary API) �Լ���
'   ���� �۾��� ���� ����, ���� �۾��� ���� ���θ� Ȯ���ؾ� �մϴ�.
' - �۾� ����(jobState) = 3(�Ϸ�)�̰� ���� ��� �ڵ�(errorCode) = 1(��������)�̸�
'   �ŷ� ���� ��ȸ(Search) �Ǵ� �ŷ� ��� ���� ��ȸ(Summary) �� �ؾ��մϴ�.
' - �۾� ����(jobState)�� 3(�Ϸ�)������ ���� ��� �ڵ�(errorCode)�� 1(��������)�� �ƴ� ��쿡��
'   �����޽���(errorReason)�� ���� ���п� ���� ������ �ľ��� �� �ֽ��ϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetJobState
'=========================================================================
Private Sub btnGetJobState_Click()
    Dim jobInfo As PBHTCashbillJobState
    Dim tmp As String
    
    Set jobInfo = htCashbillService.GetJobState(txtCorpNum.Text, txtJobID.Text)
     
    If jobInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
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

'=========================================================================
' ���ݿ����� ����/���� ���� ������û�� ���� ���� ����� Ȯ���մϴ�.
' - ���� ��û �� 1�ð��� ����� ���� ��û���� ���������� ��ȯ���� �ʽ��ϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#ListActiveJob
'=========================================================================
Private Sub btnListActiveJob_Click()
    Dim jobList As Collection
    Dim tmp As String
    Dim info As PBHTCashbillJobState
    
    Set jobList = htCashbillService.ListActiveJob(txtCorpNum.Text)
     
    If jobList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "�۾����̵�(jobID)�� ��ȿ�ð��� 1�ð��Դϴ�" + vbCrLf + vbCrLf
    tmp = tmp + "jobID(�۾����̵�) | jobState(��������) | queryType(��������) | queryDateType(��������) | queryStDate(��������) | queryEnDate(��������) |" _
            + "errorCode(�����ڵ�) | errorReason(�����޽���) | jobStartDT(�۾� �����Ͻ�) | jobEndDT(�۾� �����Ͻ�) | collectCount(��������) | regDT(���� ��û�Ͻ�) " + vbCrLf
    
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
' �Լ� GetJobState(���� ���� Ȯ��)�� ���� ���� ���� Ȯ�ε� �۾����̵� Ȱ���Ͽ� ���ݿ����� ����/���� ������ ��ȸ�մϴ�.
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

    Set SearchList = htCashbillService.Search(txtCorpNum.Text, txtJobID.Text, tradeType, tradeUsage, page, perPage, order)
    
    If SearchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (�����ڵ�) : " + CStr(SearchList.code) + vbCrLf
    tmp = tmp + "message (����޽���) : " + SearchList.Message + vbCrLf
    tmp = tmp + "total (�� �˻���� �Ǽ�) : " + CStr(SearchList.total) + vbCrLf
    tmp = tmp + "perPage (�������� �˻�����) : " + CStr(SearchList.perPage) + vbCrLf
    tmp = tmp + "pageNum (������ ��ȣ) : " + CStr(SearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (������ ����) : " + CStr(SearchList.pageCount) + vbCrLf + vbCrLf
    
    cashbillList.Clear
    
    cashbillList.AddItem "ntsconfirmNum (����û���ι�ȣ) | tradeDate (�ŷ�����) | tradeDT (�ŷ��Ͻ�) | tradeType (��������) | tradeUsage (�ŷ�����) | totalAmount (�ŷ��ݾ�)", 0
    cashbillList.AddItem "supplyCost (���ް���) | tax (�ΰ���) | serviceFee (�����) | invoiceType (����/����) | franchiseCorpNum (������ ����ڹ�ȣ) ", 1
    cashbillList.AddItem "franchiseCorpName (������ ��ȣ) | franchiseCorpType (������ ���������) | identityNum (�ŷ�ó �ĺ���ȣ) | identityNumType (�ĺ���ȣ����)", 2
    cashbillList.AddItem "customerName (����) | cardOwnerName (ī������ڸ�) | deductionType (��������)", 3

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
' �Լ� GetJobState(���� ���� Ȯ��)�� ���� ���� ������ Ȯ�ε� �۾����̵� Ȱ���Ͽ� ������ ���ݿ����� ����/���� ������ ��� ������ ��ȸ�մϴ�.
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
    
    '���ݿ����� ���� �迭, N-�Ϲ����ݿ�����, C-������ݿ�����
    tradeType.Add "N"
    tradeType.Add "C"
        
    '�ŷ��뵵 �迭, P-�ҵ������, C-����������
    tradeUsage.Add "P"
    tradeUsage.Add "C"
       
    'Summary ȣ��
    Set summaryInfo = htCashbillService.Summary(txtCorpNum.Text, txtJobID.Text, tradeType, tradeUsage)
    
    If summaryInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "count (��������Ǽ�) : " + CStr(summaryInfo.count) + vbCrLf
    tmp = tmp + "supplyCostTotal (���ް��� �հ�) : " + CStr(summaryInfo.supplyCostTotal) + vbCrLf
    tmp = tmp + "taxTotal (���� �հ�) : " + CStr(summaryInfo.taxTotal) + vbCrLf
    tmp = tmp + "serviceFeeTotal (����� �հ�) : " + CStr(summaryInfo.serviceFeeTotal) + vbCrLf
    tmp = tmp + "amountTotal (�հ� �ݾ�) : " + CStr(summaryInfo.amountTotal) + vbCrLf
           
    MsgBox (tmp)
End Sub

'=========================================================================
' Ȩ�ý����� ������ ���� ��û �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetFlatRatePopUpURL
'=========================================================================
Private Sub btnGetFlatRatePopUpURL_Click()
    Dim url As String
    
    url = htCashbillService.GetFlatRatePopUpURL(txtCorpNum.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' Ȩ�ý����� ������ ���� ���¸� Ȯ���մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetFlatRateState
'=========================================================================
Private Sub btnGetFlatRateState_Click()
    Dim flatRateInfo As PBHTCashbillFlatRate
    Dim tmp As String
    
    Set flatRateInfo = htCashbillService.GetFlatRateState(txtCorpNum.Text)
     
    If flatRateInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
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

'=========================================================================
' Ȩ�ý����� ���������� �����ϴ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ������Ŀ��� �μ������/���������� ���� ����� �ֽ��ϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetCertificatePopUpURL
'=========================================================================
Private Sub btnGetCertificatePopUpURL_Click()
    Dim url As String
    
    url = htCashbillService.GetCertificatePopUpURL(txtCorpNum.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' Ȩ�ý����� ������ ���� �˺��� ��ϵ� ������ �������ڸ� Ȯ���մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#GetCertificateExpireDate
'=========================================================================
Private Sub btnGetCertificateExpireDate_Click()
    Dim expireDate As String
    
    expireDate = htCashbillService.GetCertificateExpireDate(txtCorpNum.Text)
    
    If expireDate = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������������ : " + expireDate
End Sub

'=========================================================================
' �˺��� ��ϵ� �������� Ȩ�ý� �α��� ���� ���θ� Ȯ���մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#CheckCertValidation
'=========================================================================
Private Sub btnCheckCertValidation_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.CheckCertValidation(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' Ȩ�ý����� ������ ���� �˺��� ��ϵ� ���ݿ����� �ڷ���ȸ �μ������ ������ Ȯ���մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#CheckDeptUser
'=========================================================================
Private Sub btnCheckDeptUser_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.CheckDeptUser(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' �˺��� ��ϵ� ���ݿ����� �ڷ���ȸ �μ������ ���� ������ Ȩ�ý� �α��� ���� ���θ� Ȯ���մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#CheckLoginDeptUser
'=========================================================================
Private Sub btnCheckLoginDeptUser_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.CheckLoginDeptUser(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' �˺��� ��ϵ� Ȩ�ý� ���ݿ����� �ڷ���ȸ �μ������ ������ �����մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#DeleteDeptUser
'=========================================================================
Private Sub btnDeleteDeptUser_Click()
    Dim Response As PBResponse
    
    Set Response = htCashbillService.DeleteDeptUser(txtCorpNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' Ȩ�ý����� ������ ���� �˺��� ���ݿ����� �ڷ���ȸ �μ������ ������ ����մϴ�.
' - https://docs.popbill.com/htcashbill/vb/api#RegistDeptUser
'=========================================================================
Private Sub btnRegistDeptUser_Click()
    Dim Response As PBResponse
    Dim DeptUserID As String
    Dim DeptUserPWD As String
    
    'Ȩ�ý����� ������ ���ݿ����� �μ������ ���̵�
    DeptUserID = "userid_test"
    
    'Ȩ�ý����� ������ ���ݿ����� �μ������ ��й�ȣ
    DeptUserPWD = "passwd_test"
    
    Set Response = htCashbillService.RegistDeptUser(txtCorpNum.Text, DeptUserID, DeptUserPWD)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

Private Sub Form_Load()

    '��� �ʱ�ȭ
    htCashbillService.Initialize linkID, SecretKey
    
    '����ȯ�� ������ True(�׽�Ʈ��), False(�����)
    htCashbillService.IsTest = True
    
    '������ū IP���ѱ�� ��뿩��, True-���, False-�̻��, �⺻��(True)
    htCashbillService.IPRestrictOnOff = True
    
    '���ýý��� �ð� ��뿩�� True-���, Fasle-�̻��, �⺻��(False)
    htCashbillService.UseLocalTimeYN = False
    
End Sub

