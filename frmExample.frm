VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� Ȩ�ý� ���ݿ����� ���Ը��� ��ȸ API SDK Example"
   ClientHeight    =   11430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17355
   LinkTopic       =   "Form1"
   ScaleHeight     =   11430
   ScaleWidth      =   17355
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.Frame Frame6 
      Caption         =   "Ȩ�ý� ���ݿ����� ���� ���� API"
      Height          =   7935
      Left            =   240
      TabIndex        =   20
      Top             =   3240
      Width           =   16935
      Begin VB.Frame Frame12 
         Caption         =   "Ȩ�ý� �������� ���"
         Height          =   2415
         Left            =   8280
         TabIndex        =   41
         Top             =   360
         Width           =   5895
         Begin VB.CommandButton btnDeleteDeptUser 
            Caption         =   "�μ������ ������� ����"
            Height          =   375
            Left            =   3000
            TabIndex        =   48
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton btnCheckLoginDeptUser 
            Caption         =   "�μ������ �α��� �׽�Ʈ"
            Height          =   375
            Left            =   3000
            TabIndex        =   47
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton btnCheckDeptUser 
            Caption         =   "�μ������ ������� Ȯ��"
            Height          =   375
            Left            =   3000
            TabIndex        =   46
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton btnRegistDeptUser 
            Caption         =   "�μ������ �������"
            Height          =   375
            Left            =   240
            TabIndex        =   45
            Top             =   1800
            Width           =   2535
         End
         Begin VB.CommandButton btnCheckCertValidation 
            Caption         =   "���������� �α��� �׽�Ʈ"
            Height          =   375
            Left            =   240
            TabIndex        =   44
            Top             =   1320
            Width           =   2535
         End
         Begin VB.CommandButton btnGetCertificateExpireDate 
            Caption         =   "���������� �������� Ȯ��"
            Height          =   410
            Left            =   240
            TabIndex        =   43
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton btnGetCertificatePopUpURL 
            Caption         =   "Ȩ�ý� �������� �˾� URL"
            Height          =   410
            Left            =   240
            TabIndex        =   42
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.TextBox txtJobID 
         Height          =   330
         Left            =   1920
         TabIndex        =   33
         Top             =   2760
         Width           =   3135
      End
      Begin VB.ListBox cashbillList 
         Height          =   4020
         Left            =   240
         TabIndex        =   31
         Top             =   3480
         Width           =   10935
      End
      Begin VB.Frame Frame9 
         Caption         =   "�ΰ����"
         Height          =   2415
         Left            =   5280
         TabIndex        =   28
         Top             =   360
         Width           =   2775
         Begin VB.CommandButton btnGetFlatRatePopUpURL 
            Caption         =   "������ ���� ��û URL"
            Height          =   410
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   2535
         End
         Begin VB.CommandButton btnGetFlatRateState 
            Caption         =   "������ ���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   2535
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "����/���� ������� ��ȸ"
         Height          =   1935
         Left            =   2640
         TabIndex        =   25
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton btnSearch 
            Caption         =   "���� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnSummary 
            Caption         =   "���� ��� ������� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   26
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "����/���� ���� ����"
         Height          =   1935
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnRequestJob 
            Caption         =   "���� ��û"
            Height          =   410
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton btnGetJobState 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnListActiveJob 
            Caption         =   "���� ���� ��� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   22
            Top             =   1320
            Width           =   1935
         End
      End
      Begin VB.Label Label4 
         Caption         =   "(�۾����̵�� '���� ��û' ȣ��� �����˴ϴ�.)"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   3120
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "�۾����̵�(jobID) :"
         Height          =   180
         Left            =   240
         TabIndex        =   32
         Top             =   2835
         Width           =   1695
      End
   End
   Begin VB.CommandButton btnUpdateCorpInfo 
      Caption         =   "ȸ������ ����"
      Height          =   410
      Left            =   14640
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton btnUpdateContact 
      Caption         =   "����� ���� ����"
      Height          =   410
      Left            =   5040
      TabIndex        =   4
      Top             =   2235
      Width           =   2055
   End
   Begin VB.CommandButton btnListContact 
      Caption         =   "����� ��� ��ȸ"
      Height          =   410
      Left            =   5040
      TabIndex        =   3
      Top             =   1755
      Width           =   2055
   End
   Begin VB.CommandButton btnCheckID 
      Caption         =   "ID �ߺ� Ȯ��"
      Height          =   410
      Left            =   600
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
      Left            =   14520
      TabIndex        =   6
      Top             =   960
      Width           =   2240
      Begin VB.CommandButton btnGetCorpInfo 
         Caption         =   "ȸ������ ��ȸ"
         Height          =   410
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2490
      Left            =   240
      TabIndex        =   8
      Top             =   570
      Width           =   16935
      Begin VB.Frame Frame11 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ "
         Height          =   1935
         Left            =   9360
         TabIndex        =   36
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "�������� ����Ʈ "
         Height          =   1935
         Left            =   7080
         TabIndex        =   35
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton btnPopbillURL_CHRG 
            Caption         =   " ����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ����"
         Height          =   1935
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "����� ����"
         Height          =   1935
         Left            =   4680
         TabIndex        =   11
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   1935
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   16
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         Height          =   1935
         Left            =   11880
         TabIndex        =   9
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
            Caption         =   " �˺� �α��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   2055
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4560
      TabIndex        =   19
      Top             =   195
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ :"
      Height          =   180
      Left            =   360
      TabIndex        =   18
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
' �˺� Ȩ�ý� ���ݿ����� ���Ը��� API VB 6.0 SDK Example
'
' - VB6 SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/569/
' - ������Ʈ ���� : 2018-10-04
' - ���� ������� ����ó : 1600-8536 / 070-4304-2991
' - ���� ������� �̸��� : code@linkhub.co.kr
'
' <�׽�Ʈ �������� �غ����>
' 1) 29, 32�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
' 3) Ȩ�ý����� �̿밡���� ������������ ����մϴ�.
'    - �˺��α��� > [Ȩ�ý�����] > [ȯ�漳��] > [���������� ����] �޴�
'    - ���������� ���(GetCertificatePopUpURL API) ��ȯ�� URL�� �̿��Ͽ�
'      �˾� ���������� ���������� ���
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

Private htCashbillService As New PBHTCashbillService

'=========================================================================
' �˺��� ��ϵ� ������������ Ȩ�ý� �α����� �׽�Ʈ�Ѵ�.
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
' �˺��� ��ϵ� ���ݿ����� �μ������ ���̵� Ȯ���մϴ�.
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
' �˺� ȸ�����̵� �ߺ����θ� Ȯ���մϴ�.
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
' �ش� ������� ��Ʈ�� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
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
' �˺��� ��ϵ� ���ݿ����� �μ������ ���������� �̿��Ͽ� Ȩ�ý� �α����� �׽�Ʈ�մϴ�.
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
'  �˺��� ��ϵ� ���ݿ����� �μ������ ���������� �����մϴ�.
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
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)
'   �� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = htCashbillService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��ϵ� Ȩ�ý� ������������ �������ڸ� Ȯ���մϴ�.
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
' Ȩ�ý� ���������� ��� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�� URL�� ������å�� ���� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetCertificatePopUpURL_Click()
    Dim url As String
    
    url = htCashbillService.GetCertificatePopUpURL(txtCorpNum.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ���� Ȩ�ý� ���ݿ����� ���� API ���� ���������� Ȯ���մϴ�.
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
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
'=========================================================================

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = htCashbillService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname(��ǥ�ڼ���) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName(��ȣ��) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr(�ּ�) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType(����) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass(����) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
    
End Sub

'=========================================================================
' ������ ��û �˾� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetFlatRatePopUpURL_Click()
    Dim url As String
    
    url = htCashbillService.GetFlatRatePopUpURL(txtCorpNum.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ���� ������ ���� �̿���¸� Ȯ���մϴ�.
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
' ���� ��û ���¸� Ȯ���մϴ�.
' - �����׸� ���� ������ "[Ȩ�ý� ���ݿ����� ���� API �����Ŵ���
'   > 3.2.2. GetJobState (���� ���� Ȯ��)" �� �����Ͻñ� �ٶ��ϴ� .
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
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)��
'   �̿��Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = htCashbillService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = htCashbillService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺�(www.popbill.com)�� �α��ε� �˺� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = htCashbillService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺� ����ȸ�� ������ ��û�մϴ�.
'=========================================================================

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '��ũ ���̵�
    joinData.linkID = linkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1231212312"
    
    '��ǥ�ڼ���, �ִ� 30��
    joinData.ceoname = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 70��
    joinData.corpName = "ȸ����ȣ"
    
    '�ּ�, �ִ� 300��
    joinData.addr = "�ּ�"
    
    '����, �ִ� 40��
    joinData.bizType = "����"
    
    '����, �ִ� 40��
    joinData.bizClass = "����"
    
    '���̵�, 6���̻� 20�� �̸�
    joinData.id = "userid"
    
    '��й�ȣ, 6���̻� 20�� �̸�
    joinData.pwd = "pwd_must_be_long_enough"
    
    '����ڸ�, �ִ� 30��
    joinData.ContactName = "����ڼ���"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    '����� ����, �ִ� 70��
    joinData.ContactEmail = "test@test.com"
    
    Set Response = htCashbillService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ���� ��û�ǵ鿡 ���� ���� ����� Ȯ���մϴ�.
' - ���� ��û �۾����̵�(JobID)�� ��ȿ�ð��� 1�ð� �Դϴ�.
' - �����׸� ���� ������ "[Ȩ�ý� ���ݿ����� ���� API �����Ŵ���]
'   > 3.2.3. ListActiveJob (���� ���� ��� Ȯ��)" �� �����Ͻñ� �ٶ��ϴ�.
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
    tmp = tmp + "jobID | jobState | queryType | queryDateType | queryStDate | queryEnDate | errorCode | errorReason | jobStartDT | jobEndDT | collectCount | regDT " + vbCrLf
    
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
' ����ȸ���� ����� ����� Ȯ���մϴ�.
'=========================================================================

Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = htCashbillService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT | state" + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnPopbillURL_CHRG_Click()
    Dim url As String
    
    url = htCashbillService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ���� ����ڸ� �űԷ� ����մϴ�.
'=========================================================================

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 20�� �̸�
    joinData.id = "testkorea_20161011"
    
    '��й�ȣ, 6�� �̻� 20�� �̸�
    joinData.pwd = "test@test.com"
    
    '����ڸ�, �ִ� 30��
    joinData.personName = "����ڸ�"
    
    '����� ����ó
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '����� �����ּ�
    joinData.email = "test@test.com"
    
    '����� �ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    'ȸ����ȸ ���ѿ���, true-ȸ����ȸ / false-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
        
    Set Response = htCashbillService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' Ȩ�ý� ���ݿ����� �μ������ ������ ����մϴ�.
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

'=========================================================================
' ���ݿ����� ����/���� ���� ������ ��û�մϴ�
' - ����/���� ���� ���μ����� "[Ȩ�ý� ���ݿ����� ���� API �����Ŵ���]
'   > 1.2. ���μ��� �帧��" �� �����Ͻñ� �ٶ��ϴ�.
' - ���� ��û�� ��ȯ���� �۾����̵�(JobID)�� ��ȿ�ð��� 1�ð� �Դϴ�.
'=========================================================================

Private Sub btnRequestJob_Click()
    Dim jobID As String
    Dim SDate As String
    Dim EDate As String
    Dim cbType As KeyType
    
    '���ݿ����� ����, SELL-����, BUY-����, TURSTEE-����Ź
    cbType = SELL
        
    '��������, ǥ������(yyyyMMdd)
    SDate = "20160901"
    
    '��������, ǥ������(yyyyMMdd)
    EDate = "20161031"
        
    jobID = htCashbillService.RequestJob(txtCorpNum.Text, cbType, SDate, EDate)
    
    If jobID = "" Then
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "jobID(�۾����̵�) : " + jobID + vbCrLf
    
    txtJobID.Text = jobID
End Sub

'=========================================================================
' �˻������� ����Ͽ� ��������� ��ȸ�մϴ�.
' - �����׸� ���� ������ "[Ȩ�ý� ���ݿ����� ���� API �����Ŵ���]
'   > 3.3.1. Search (���� ��� ��ȸ)" �� �����Ͻñ� �ٶ��ϴ�.
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
    Dim cbInfo As PBHTCashbill
    
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
    
    cashbillList.AddItem "���� | ����/���� | �ŷ��Ͻ� | �ĺ���ȣ | ���ް��� | ���� | ����� | �ŷ��ݾ� | �������� | ����û���ι�ȣ", 0
    
    For Each cbInfo In SearchList.list
        ' �߰����� ���ݿ����� �׸��� [Ȩ�ý� ���ݿ����� ���� API �����Ŵ��� > 4.1.�������� ����] �� �����Ͻñ� �ٶ��ϴ�.'
        rowTmp = ""
        rowTmp = cbInfo.tradeUsage + " | "
        rowTmp = rowTmp + cbInfo.invoiceType + " | "
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

'=========================================================================
' �˻������� ����Ͽ� ���� ��� ��������� ��ȸ�մϴ�.
' - �����׸� ���� ������ "[Ȩ�ý� ���ݿ����� ���� API �����Ŵ���]
'   > 3.3.2. Summary (���� ��� ������� ��ȸ)" �� �����Ͻñ� �ٶ��ϴ�.
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
' ����ȸ���� ����� ������ �����մϴ�.
'=========================================================================

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
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
        MsgBox ("�����ڵ� : " + CStr(htCashbillService.LastErrCode) + vbCrLf + "����޽��� : " + htCashbillService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.Message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
'=========================================================================

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ
    CorpInfo.corpName = "��ȣ"
    
    '�ּ�
    CorpInfo.addr = "����Ư����"
    
    '����
    CorpInfo.bizType = "����"
    
    '����
    CorpInfo.bizClass = "����"
    
    Set Response = htCashbillService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
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
End Sub

