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
Private cmdSp As New ADODB.Command  '创建命令对象实例

Public Enum eType
    adArray             ' 与其他类型一起加入逻辑 OR 以指示该数据是哪种类型的安全数组 (DBTYPE_ARRAY)。
    adBigInt            ' 8 - 字节带符号的整数 (DBTYPE_I8)。
    adBinary            ' 二进制值 (DBTYPE_BYTES)。
    adBoolean           ' 布尔型值 (DBTYPE_BOOL)。
    adByRef             '与其他类型一起加入逻辑 OR 以指示该数据是其他类型数据的指针 (DBTYPE_BYREF)。
    adBSTR              ' 以空结尾的字符串 (Unicode) (DBTYPE_BSTR)。
    adChar              ' 字符串值 (DBTYPE_STR)。
    adCurrency          ' 货币值 (DBTYPE_CY)。货币数字的小数点位置固定、小数点右侧有四位数字。该值保存为 8 字节的范围为 10,000 的带符号整型值。
    adDate              ' 日期值 (DBTYPE_DATE)。日期按双精度型数值来保存，数字全部表示从 1899 年 12 月 30 开始的日期数。小数部分是一天当中的片段时间。
    adDBDate            ' 日期值 (yyyymmdd) (DBTYPE_DBDATE)。
    adDBTime            ' 时间值 (hhmmss) (DBTYPE_DBTIME)。
    adDBTimeStamp       ' 时间戳（yyyymmddhhmmss 加 10 亿分之一的小数）(DBTYPE_DBTIMESTAMP).
    adDecimal           ' 具有固定精度和范围的精确数字值 (DBTYPE_DECIMAL)。
    adDouble            ' 双精度浮点值 (DBTYPE_R8)。
    adEmpty             ' 未指定值 (DBTYPE_EMPTY)。
    adError             ' 32 位错误代码 (DBTYPE_ERROR)。
    adGUID              ' 全局唯一的标识符 (GUID) (DBTYPE_GUID)。
    adIDispatch         ' OLE 对象上 Idispatch 接口的指针 (DBTYPE_IDISPATCH)。
    adInteger           ' 4 字节的带符号整型 (DBTYPE_I4)。
    adIUnknown          ' OLE 对象上 IUnknown 接口的指针 (DBTYPE_IUNKNOWN)。
    adLongVarBinary     ' 长二进制值（仅用于 Parameter 对象）。
    adLongVarChar       ' 长字符串值（仅用于 Parameter 对象）。
    adLongVarWChar      ' 以空结尾的长字符串值（仅用于 Parameter 对象）。
    adNumeric           ' 具有固定精度和范围的精确数字值 (DBTYPE_NUMERIC)。
    adSingle            ' 单精度浮点值 (DBTYPE_R4)。
    adSmallInt          ' 2 字节带符号整型 (DBTYPE_I2)。
    adTinyInt           ' 1 字节带符号整型 (DBTYPE_I1)。
    adUnsignedBigInt    ' 8 字节不带符号整型 (DBTYPE_UI8)。
    adUnsignedInt       ' 4 字节不带符号整型 (DBTYPE_UI4)。
    adUnsignedSmallInt  ' 2 字节不带符号整型 (DBTYPE_UI2)。
    adUnsignedTinyInt   ' 1 字节不带符号整型 (DBTYPE_UI1)。
    adUserDefined       ' 用户定义的变量 (DBTYPE_UDT)。
    adVarBinary         ' 二进制值（仅 Parameter 对象）。
    adVarChar           ' 字符串值（仅 Parameter 对象）。
    adVariant           ' 自动变体型 (DBTYPE_VARIANT)。
    adVector            ' 与其他类型一起加入逻辑 OR 中，指示数据是 DBVECTOR 结构（由 OLE DB 定义）。该结构含有元素的计数和其他类型 (DBTYPE_VECTOR) 数据的指针。
    adVarWChar          ' 以空结尾的 Unicode 字符串（仅 Parameter 对象）。
    adWChar             ' 以空结尾的 Unicode 字符串 (DBTYPE_WSTR)。
End Enum

Public Enum eDirection
    adParamUnknown       ' 指示参数方向为未知.
    adParamInput         ' 默认值.指示输入参数.
    adParamOutput        ' 指示为输出参数.
    adParamInputOutput   ' 指示为输入参数和输出参数.
    adParamReturnValue   ' 指示为返回值.
End Enum

'对象方法：执行存储过程
'参数： SpName:存储过程名(传入参数),
'       DbClass:数据库连接串 (传入参数)
'返回值：存储过程中最后一条Select语句的纪录集
Public Function exeSp(ByVal SpName As String, Optional ByVal DbClass As DBClassType = 1) As Variant
    Dim strConn As String
    Dim clsObj As clsDML1
    Dim ErrorNo As Long
    
    ErrorNo = 0
    
    'get DB connect string
    Set clsObj = New clsDML1
    strConn = clsObj.GetConnStr(DbClass)
    
    On Error GoTo ErrorHandlerKnow
    'open database connect
    ErrorNo = DBConnect(strConn)
    On Error GoTo ErrorHandlerUnknow
    '打开命令对象
    cmdSp.CommandText = SpName
    cmdSp.CommandType = adCmdStoredProc
    '执行命令  或  通过执行命令获得记录集
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

'功能：打开database连接
'参数：strConn：ADO Connect 连接字符串(传入参数)
'返回值：出错代码(ErrorNo)
Private Function DBConnect(ByVal strConn As String) As Long
    Dim clsObj As clsDML1
    Dim ErrorNo As Long
    ErrorNo = 0
    
    On Error GoTo ErrorHandler
    Set Conn = New ADODB.Connection
    Conn.Open strConn
    Conn.CursorLocation = adUseClient
    Exit Function
    
ErrorHandler:
    Set clsObj = New clsDML1
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

'对象方法:向sp参数集合中添加一个参数对象
'参数： Name:可选，字符串，代表 Parameter 对象名称。
'        Types:可选，长整型值，指定 Parameter 对象数据类型,
'        Direction:可选，长整型值，指定 Parameter 对象类型,
'        Size: 可选，长整型值，指定参数值最大长度（以字符或字节数为单位）。
'        Value: 可选，变体型，指定 Parameter 对象值.
'返回值:当前添加的parameter对象
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

    '为新的 Parameter 返回一个引用。
    Set ParamAdd = Param
    Set Param = Nothing
    Exit Function
    
ErrorHandler:
    Set Param = Nothing
    ErrorNo = 10020
    Err.Raise ErrorNo, "Com_DML.clsSp.ParamAdd", "Append Parameter object error"
End Function

'方法:取参数集合中的参数对象数
'返回值：集合中的对象数
Public Function ParamCount() As Long
    On Error GoTo ErrorHandler
    ParamCount = cmdSp.Parameters.Count
    Exit Function
ErrorHandler:
    Err.Raise 10022, "Com_DML.clsSP.ParamCount", "fail in getting count of  parameter in ADOBO.Command.Parameters"
End Function

'对象方法:删除集合中的一个参数对象
'参数：Index：Parameter对象的序号(传入参数）
'返回值：无
Public Sub ParamDel(ByVal Index As Integer)
    On Error GoTo ErrorHandler
    cmdSp.Parameters.Delete (Index)
    Exit Sub
ErrorHandler:
    Err.Raise 10023, "Com_DML.clsSP.ParamDel", "fail in deleting parameter in ADOBO.Command.Parameters"
End Sub

'对象方法:取集合中的一个参数对象
'参数：Index：Parameter对象的序号(传入参数）
'返回值：集合中序号为Index的参数对象
Public Function ParamItem(ByVal Index As Integer) As ADODB.Parameter
    On Error GoTo ErrorHandler
    Set ParamItem = cmdSp.Parameters.Item(Index)
    Exit Function
ErrorHandler:
    Err.Raise 10024, "Com_DML.clsSP.ParamItem", "fail in getting a parameter object in ADOBO.Command.Parameters"
End Function

'方法：Parameters collect 集合刷新
'返回值：无
Public Sub ParamRefresh()
    On Error GoTo ErrorHandler
    cmdSp.Parameters.Refresh
    Exit Sub
ErrorHandler:
    Err.Raise 10035, "Com_DML.clsSP.ParamRefresh", "fail in refreshing stored procedures parameters collect"
End Sub

