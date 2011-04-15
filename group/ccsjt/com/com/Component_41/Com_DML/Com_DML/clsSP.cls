VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "clsSP1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Conn As ADODB.Connection
Private cmdSp As New ADODB.Command  '�����������ʵ��

Public Enum eType
    adArray             ' ����������һ������߼� OR ��ָʾ���������������͵İ�ȫ���� (DBTYPE_ARRAY)��
    adBigInt            ' 8 - �ֽڴ����ŵ����� (DBTYPE_I8)��
    adBinary            ' ������ֵ (DBTYPE_BYTES)��
    adBoolean           ' ������ֵ (DBTYPE_BOOL)��
    adByRef             '����������һ������߼� OR ��ָʾ�������������������ݵ�ָ�� (DBTYPE_BYREF)��
    adBSTR              ' �Կս�β���ַ��� (Unicode) (DBTYPE_BSTR)��
    adChar              ' �ַ���ֵ (DBTYPE_STR)��
    adCurrency          ' ����ֵ (DBTYPE_CY)���������ֵ�С����λ�ù̶���С�����Ҳ�����λ���֡���ֵ����Ϊ 8 �ֽڵķ�ΧΪ 10,000 �Ĵ���������ֵ��
    adDate              ' ����ֵ (DBTYPE_DATE)�����ڰ�˫��������ֵ�����棬����ȫ����ʾ�� 1899 �� 12 �� 30 ��ʼ����������С��������һ�쵱�е�Ƭ��ʱ�䡣
    adDBDate            ' ����ֵ (yyyymmdd) (DBTYPE_DBDATE)��
    adDBTime            ' ʱ��ֵ (hhmmss) (DBTYPE_DBTIME)��
    adDBTimeStamp       ' ʱ�����yyyymmddhhmmss �� 10 �ڷ�֮һ��С����(DBTYPE_DBTIMESTAMP).
    adDecimal           ' ���й̶����Ⱥͷ�Χ�ľ�ȷ����ֵ (DBTYPE_DECIMAL)��
    adDouble            ' ˫���ȸ���ֵ (DBTYPE_R8)��
    adEmpty             ' δָ��ֵ (DBTYPE_EMPTY)��
    adError             ' 32 λ������� (DBTYPE_ERROR)��
    adGUID              ' ȫ��Ψһ�ı�ʶ�� (GUID) (DBTYPE_GUID)��
    adIDispatch         ' OLE ������ Idispatch �ӿڵ�ָ�� (DBTYPE_IDISPATCH)��
    adInteger           ' 4 �ֽڵĴ��������� (DBTYPE_I4)��
    adIUnknown          ' OLE ������ IUnknown �ӿڵ�ָ�� (DBTYPE_IUNKNOWN)��
    adLongVarBinary     ' ��������ֵ�������� Parameter ���󣩡�
    adLongVarChar       ' ���ַ���ֵ�������� Parameter ���󣩡�
    adLongVarWChar      ' �Կս�β�ĳ��ַ���ֵ�������� Parameter ���󣩡�
    adNumeric           ' ���й̶����Ⱥͷ�Χ�ľ�ȷ����ֵ (DBTYPE_NUMERIC)��
    adSingle            ' �����ȸ���ֵ (DBTYPE_R4)��
    adSmallInt          ' 2 �ֽڴ��������� (DBTYPE_I2)��
    adTinyInt           ' 1 �ֽڴ��������� (DBTYPE_I1)��
    adUnsignedBigInt    ' 8 �ֽڲ����������� (DBTYPE_UI8)��
    adUnsignedInt       ' 4 �ֽڲ����������� (DBTYPE_UI4)��
    adUnsignedSmallInt  ' 2 �ֽڲ����������� (DBTYPE_UI2)��
    adUnsignedTinyInt   ' 1 �ֽڲ����������� (DBTYPE_UI1)��
    adUserDefined       ' �û�����ı��� (DBTYPE_UDT)��
    adVarBinary         ' ������ֵ���� Parameter ���󣩡�
    adVarChar           ' �ַ���ֵ���� Parameter ���󣩡�
    adVariant           ' �Զ������� (DBTYPE_VARIANT)��
    adVector            ' ����������һ������߼� OR �У�ָʾ������ DBVECTOR �ṹ���� OLE DB ���壩���ýṹ����Ԫ�صļ������������� (DBTYPE_VECTOR) ���ݵ�ָ�롣
    adVarWChar          ' �Կս�β�� Unicode �ַ������� Parameter ���󣩡�
    adWChar             ' �Կս�β�� Unicode �ַ��� (DBTYPE_WSTR)��
End Enum

Public Enum eDirection
    adParamUnknown       ' ָʾ��������Ϊδ֪.
    adParamInput         ' Ĭ��ֵ.ָʾ�������.
    adParamOutput        ' ָʾΪ�������.
    adParamInputOutput   ' ָʾΪ����������������.
    adParamReturnValue   ' ָʾΪ����ֵ.
End Enum

'���󷽷���ִ�д洢����
'������ SpName:�洢������(�������),
'       DbClass:���ݿ����Ӵ� (�������)
'����ֵ���洢���������һ��Select���ļ�¼��
Public Function exeSp(ByVal SpName As String, Optional ByVal DbClass As DBClassType = 1) As Variant
    Dim strConn As String
    Dim clsObj As clsDML
    Dim ErrorNo As Long
    
    ErrorNo = 0
    
    'get DB connect string
    Set clsObj = New clsDML
    strConn = clsObj.GetConnStr(DbClass)
    
    On Error GoTo ErrorHandlerKnow
    'open database connect
    ErrorNo = DBConnect(strConn)
    On Error GoTo ErrorHandlerUnknow
    '���������
    cmdSp.CommandText = SpName
    cmdSp.CommandType = adCmdStoredProc
    'ִ������  ��  ͨ��ִ�������ü�¼��
    Set cmdSp.ActiveConnection = Conn
    Set exeSp = cmdSp.Execute
    Set clsObj = Nothing
    Exit Function

ErrorHandlerUnknow:
    ErrorNo = clsObj.GetDBError(Conn.Errors(0).NativeError)
    Err.Number = ErrorNo
ErrorHandlerKnow:
    ErrorNo = Err.Number
    Set clsObj = Nothing
    If ErrorNo = 10009 Then
        ErrorNo = 10021
        Err.Raise ErrorNo, "Com_DML.clsSP.DBConnect", "execute SP unknow error"
    Else
        Err.Raise ErrorNo, Err.Source, Err.Description
    End If

End Function

'���ܣ���database����
'������strConn��ADO Connect �����ַ���(�������)
'����ֵ����������(ErrorNo)
Private Function DBConnect(ByVal strConn As String) As Long
    Dim clsObj As clsDML
    Dim ErrorNo As Long
    ErrorNo = 0
    
    On Error GoTo ErrorHandler
    Set Conn = New ADODB.Connection
    Conn.Open strConn
    Conn.CursorLocation = adUseClient
    Exit Function
    
ErrorHandler:
    Set clsObj = New clsDML
    ErrorNo = clsObj.GetDBError(Conn.Errors(0).NativeError)
    Set clsObj = Nothing
    If ErrorNo = 10009 Then
        ErrorNo = 10019
        Err.Raise ErrorNo, "Com_DML.clsSP.DBConnect", "open database unknow error"
    Else
        Err.Raise ErrorNo, Err.Source, Err.Description
    End If
    DBConnect = ErrorNo     'unsuccessful
End Function

'���󷽷�:��sp��������������һ����������
'������ Name:��ѡ���ַ��������� Parameter �������ơ�
'        Types:��ѡ��������ֵ��ָ�� Parameter ������������,
'        Direction:��ѡ��������ֵ��ָ�� Parameter ��������,
'        Size: ��ѡ��������ֵ��ָ������ֵ��󳤶ȣ����ַ����ֽ���Ϊ��λ����
'        Value: ��ѡ�������ͣ�ָ�� Parameter ����ֵ.
'����ֵ:��ǰ���ӵ�parameter����
Public Function ParamAdd(Optional ByVal mName As String, _
                    Optional ByVal mTypes As eType, _
                    Optional ByVal mDirection As eDirection = adParamInput, _
                    Optional ByVal mSize As Long, _
                    Optional ByRef mValue As Variant) As ADODB.Parameter
    Dim Param As New ADODB.Parameter
    Dim ErrorNo As Long
    ErrorNo = 0
    On Error GoTo ErrorHandler
    Set Param = cmdSp.CreateParameter(mName, mTypes, mDirection, mSize)
    cmdSp.Parameters.Append Param
    If Not IsMissing(mValue) Then
        Param.Value = mValue
    End If

    'Ϊ�µ� Parameter ����һ�����á�
    Set ParamAdd = Param
    Set Param = Nothing
    Exit Function
    
ErrorHandler:
    Set Param = Nothing
    ErrorNo = 10020
    Err.Raise ErrorNo, "Com_DML.clsSp.ParamAdd", "Append Parameter object error"
End Function

'����:ȡ���������еĲ���������
'����ֵ�������еĶ�����
Public Function ParamCount() As Long
    On Error GoTo ErrorHandler
    ParamCount = cmdSp.Parameters.Count
    Exit Function
ErrorHandler:
    Err.Raise 10022, "Com_DML.clsSP.ParamCount", "fail in getting count of  parameter in ADOBO.Command.Parameters"
End Function

'���󷽷�:ɾ�������е�һ����������
'������Index��Parameter��������(���������
'����ֵ����
Public Sub ParamDel(ByVal Index As Integer)
    On Error GoTo ErrorHandler
    cmdSp.Parameters.Delete (Index)
    Exit Sub
ErrorHandler:
    Err.Raise 10023, "Com_DML.clsSP.ParamDel", "fail in deleting parameter in ADOBO.Command.Parameters"
End Sub

'���󷽷�:ȡ�����е�һ����������
'������Index��Parameter��������(���������
'����ֵ�����������ΪIndex�Ĳ�������
Public Function ParamItem(ByVal Index As Integer) As ADODB.Parameter
    On Error GoTo ErrorHandler
    Set ParamItem = cmdSp.Parameters.Item(Index)
    Exit Function
ErrorHandler:
    Err.Raise 10024, "Com_DML.clsSP.ParamItem", "fail in getting a parameter object in ADOBO.Command.Parameters"
End Function

'������Parameters collect ����ˢ��
'����ֵ����
Public Sub ParamRefresh()
    On Error GoTo ErrorHandler
    cmdSp.Parameters.Refresh
    Exit Sub
ErrorHandler:
    Err.Raise 10035, "Com_DML.clsSP.ParamRefresh", "fail in refreshing stored procedures parameters collect"
End Sub
