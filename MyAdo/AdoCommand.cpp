#include "stdafx.h"
#include "resource.h"
#include "AdoCommand.h"


#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

CAdoCommand::CAdoCommand()
{
	///创建 Connection 对象---------------------------
	m_pCommand.CreateInstance(__uuidof(Command));
	#ifdef _DEBUG
	if (m_pCommand == NULL)
	{
		AfxMessageBox("Command 对象创建失败! 请确认是否初始化了Com环境.");
	}
	#endif
	assert(m_pCommand != NULL);
}

CAdoCommand::CAdoCommand(CAdoConnection* pAdoConnection, CString strCommandText, CommandTypeEnum CommandType)
{
	///创建 Connection 对象---------------------------
	m_pCommand.CreateInstance(__uuidof(Command)); //"ADODB.Command"
	#ifdef _DEBUG
	if (m_pCommand == NULL)
	{
		AfxMessageBox("Command 对象创建失败! 请确认是否初始化了Com环境.");
	}
	#endif
	assert(m_pCommand != NULL);
	assert(pAdoConnection != NULL);
	SetConnection(pAdoConnection);
	if (strCommandText != _T(""))
	{
		SetCommandText(LPCTSTR(strCommandText));
	}
	SetCommandType(CommandType);
}

CAdoCommand::~CAdoCommand()
{
	Release();
}

void CAdoCommand::Release()
{
	try
	{
		m_pCommand.Release();
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: Release方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
	} 
}

_RecordsetPtr CAdoCommand::Execute(long Options)
{
	assert(m_pCommand != NULL);
	try
	{
		return m_pCommand->Execute(NULL, NULL, Options);
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: Execute 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return NULL;
	} 
}

bool CAdoCommand::Cancel()
{
	assert(m_pCommand != NULL);
	
	try
	{
		return (m_pCommand->Cancel() == S_OK);
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: Cancel 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	} 
}

_ParameterPtr CAdoCommand::CreateParameter(LPCTSTR lpstrName, 
								  DataTypeEnum Type, 
								  ParameterDirectionEnum Direction, 
								  long Size, 
								  _variant_t Value)
{

	assert(m_pCommand != NULL);
	try
	{
		return m_pCommand->CreateParameter(_bstr_t(lpstrName), Type, Direction, Size, Value);
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: CreateParameter 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return NULL;
	} 
}

bool CAdoCommand::SetCommandText(LPCTSTR lpstrCommand)
{
	assert(m_pCommand != NULL);
	try
	{
		m_pCommand->PutCommandText(_bstr_t(lpstrCommand));
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: PutCommandText 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	} 
}

bool CAdoCommand::SetConnection(CAdoConnection *pConnect)
{
	assert(pConnect != NULL);
	assert(pConnect->GetConnection() != NULL);
	assert(m_pCommand != NULL);
	
	try
	{
		m_pCommand->PutActiveConnection(_variant_t((IDispatch*)pConnect->GetConnection(), true));
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetConnection 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	} 
}
/*========================================================================
	Name:		指示 Command 对象的类型。
    ----------------------------------------------------------
	returns:	返回以下某个 CommandTypeEnum 的值.
		[常量]				 [说明] 
		----------------------------------
		adCmdText			指示strSQL为命令文本, 即普通的SQL语句. 
		adCmdTable			指示ADO生成SQL查询返回以 strSQL 命名的表中的
						所有行. 
		adCmdTableDirect	指示所作的更改在strSQL中命名的表中返回所有行. 
		adCmdStoredProc		指示strSQL为存储过程. 
		adCmdUnknown		指示strSQL参数中的命令类型为未知. 
		adCmdFile			指示应从在strSQL中命名的文件中恢复保留(保存的)
						Recordset. 
		adAsyncExecute		指示应异步执行strSQL. 
		adAsyncFetch		指示在提取 Initial Fetch Size 属性中指定的初始
						数量后, 应该异步提取所有剩余的行. 如果所需的行尚未
						提取, 主要的线程将被堵塞直到行重新可用. 
		adAsyncFetchNonBlocking 指示主要线程在提取期间从未堵塞. 如果所请求
						的行尚未提取, 当前行自动移到文件末尾. 
   ----------------------------------------------------------
	Remarks: 使用 CommandType 属性可优化 CommandText 属性的计算。
		如果 CommandType 属性的值等于 adCmdUnknown(默认值), 系统的性能将会
	降低, 因为 ADO 必须调用提供者以确定 CommandText 属性是 SQL 语句、还是存
	储过程或表格名称。如果知道正在使用的命令的类型, 可通过设置 CommandType 
	属性指令 ADO 直接转到相关代码。如果 CommandType 属性与 CommandText 属
	性中的命令类型不匹配, 调用 Execute 方法时将产生错误。
==========================================================================*/
bool CAdoCommand::SetCommandType(CommandTypeEnum CommandType)
{
	assert(m_pCommand != NULL);
	
	try
	{
		m_pCommand->PutCommandType(CommandType);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: PutCommandType 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	} 
}

long CAdoCommand::GetState()
{
	assert(m_pCommand != NULL);
	
	try
	{
		return m_pCommand->GetState();
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetState 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return -1;
	} 
}

bool CAdoCommand::SetCommandTimeOut(long lTime)
{
	assert(m_pCommand != NULL);
	
	try
	{
		m_pCommand->PutCommandTimeout(lTime);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetCommandTimeOut 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	} 
}

ParametersPtr CAdoCommand::GetParameters()
{
	assert(m_pCommand != NULL);
	
	try
	{
		return m_pCommand->GetParameters();
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetCommandTimeOut 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return NULL;
	} 
}

bool CAdoCommand::Append(_ParameterPtr param)
{
	assert(m_pCommand != NULL);
	
	try
	{
		return m_pCommand->GetParameters()->Append((IDispatch*)param);
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: Append 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	} 
}


_ParameterPtr CAdoCommand::GetParamter(LPCTSTR lpstrName)
{
	assert(m_pCommand != NULL);
	
	try
	{
		return m_pCommand->GetParameters()->GetItem(_variant_t(lpstrName));
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetParamter 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return NULL;
	} 
}

_ParameterPtr CAdoCommand::GetParameter(long index)
{
	assert(m_pCommand != NULL);
	
	try
	{
		return m_pCommand->GetParameters()->GetItem(_variant_t(index));
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetParamter 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return NULL;
	} 
}

_variant_t CAdoCommand::GetValue(long index)
{
	assert(m_pCommand != NULL);
	
	try
	{
		return m_pCommand->GetParameters()->GetItem(_variant_t(index))->Value;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		_variant_t vt;
		vt.vt = VT_NULL;
		return vt;
	} 
}

_variant_t CAdoCommand::GetValue(LPCTSTR lpstrName)
{
	assert(m_pCommand != NULL);
	
	try
	{
		return m_pCommand->GetParameters()->GetItem(_variant_t(lpstrName))->Value;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		_variant_t vt;
		vt.vt = VT_NULL;
		return vt;
	} 
}

_CommandPtr& CAdoCommand::GetCommand()
{
	return m_pCommand;
}

CAdoParameter CAdoCommand::operator [](int index)
{
	CAdoParameter pParameter;
	assert(m_pCommand != NULL);
	try
	{
		pParameter = m_pCommand->GetParameters()->GetItem(_variant_t(long(index)));
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: operator [] 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
	}
	return pParameter;
}

CAdoParameter CAdoCommand::operator [](LPCTSTR lpszParamName)
{
	CAdoParameter pParameter;
	assert(m_pCommand != NULL);
	try
	{
		pParameter = m_pCommand->GetParameters()->GetItem(_variant_t(lpszParamName));
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: operator [] 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
	}
	return pParameter;
}


/*########################################################################
			  ------------------------------------------------
						     CAdoParameter class
			  ------------------------------------------------
  ########################################################################*/
CAdoParameter::CAdoParameter()
{
	m_pParameter = NULL;
	m_pParameter.CreateInstance(__uuidof(Parameter));
	#ifdef _DEBUG
	if (m_pParameter == NULL)
	{
		AfxMessageBox("Parameter 对象创建失败! 请确认是否初始化了Com环境.");
	}
	#endif
	assert(m_pParameter != NULL);
	m_strName = _T("");
}

CAdoParameter::CAdoParameter(DataTypeEnum DataType, long lSize, ParameterDirectionEnum Direction, CString strName)
{
	m_pParameter = NULL;
	m_pParameter.CreateInstance(__uuidof(Parameter));
	#ifdef _DEBUG
	if (m_pParameter == NULL)
	{
		AfxMessageBox("Parameter 对象创建失败! 请确认是否初始化了Com环境.");
	}
	#endif
	assert(m_pParameter != NULL);

	m_pParameter->Direction = Direction;
	m_strName = strName;
	m_pParameter->Name = m_strName.AllocSysString();
	m_pParameter->Type = DataType;
	m_pParameter->Size = lSize;
}

_ParameterPtr& CAdoParameter::operator =(_ParameterPtr& pParameter)
{
	if (pParameter != NULL)
	{
		m_pParameter = pParameter;
	}
	else
	{
		return pParameter;
	}
	return m_pParameter;
}

CAdoParameter::~CAdoParameter()
{
	m_pParameter.Release();
	m_pParameter = NULL;
	m_strName = _T("");
}

/*========================================================================
	Name:		指示在 Parameter 对象中数字值或数字 Field 对象的精度。
    ----------------------------------------------------------
	Params:		设置或返回 Byte 值，用来表示值的最大位数。该值在 Parameter
			对象上为读/写，而在 Field 对象上为只读。
    ----------------------------------------------------------
	Remarks:	使用 Precision 属性可确定表示数字 Parameter 或 Field 对象值
			的最大位数
==========================================================================*/
bool CAdoParameter::SetPrecision(char nPrecision)
{
	assert(m_pParameter != NULL);
	try
	{
		m_pParameter->PutPrecision(nPrecision);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetPrecision 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

/*========================================================================
	Name:		指出 Parameter 或 Field 对象中数字值的范围。
    ----------------------------------------------------------
	Params:		设置或返回字节值，指示数字值所精确到的小数点位数。
    ----------------------------------------------------------
	Remarks:	使用 NumericScale 属性可确定用于表明数字型 Parameter 或 Field 
		对象的值的小数位数。
		对于 Parameter 对象，NumericScale 属性为读/写。对于 Field 对象，
	NumericScale 属性为只读。

==========================================================================*/
bool CAdoParameter::SetNumericScale(int nScale)
{
	assert(m_pParameter != NULL);
	try
	{
		m_pParameter->PutNumericScale(nScale);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetPrecision 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}	
}

/*========================================================================
	Name:		指示 Parameter 对象的数据类型.
    ----------------------------------------------------------
	Params:		[DataType]: DataTypeEnum 类型值, 请参考 CRecordSet 类相关
			函数.
==========================================================================*/
bool CAdoParameter::SetType(DataTypeEnum DataType)
{
	assert(m_pParameter != NULL);
	try
	{
		m_pParameter->PutType(DataType);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetType 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}	
}

DataTypeEnum CAdoParameter::GetType()
{
	assert(m_pParameter != NULL);
	try
	{
		return m_pParameter->GetType();
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetDirection 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return adEmpty;
	}
}

/*========================================================================
	Name:		表示 Parameter 对象的最大大小（按字节或字符）。
    ----------------------------------------------------------
	Params:		[size]: 表示 Parameter 对象的最大大小（按字节或字符）的长
			整型值。
==========================================================================*/
bool CAdoParameter::SetSize(int size)
{
	assert(m_pParameter != NULL);
	try
	{
		m_pParameter->PutSize(long(size));
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetSize 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}	
}

int CAdoParameter::GetSize()
{
	assert(m_pParameter != NULL);
	try
	{
		return (int)m_pParameter->GetSize();
		
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetDirection 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return -1;
	}
}

/*========================================================================
	Name:		指示对象的名称。
==========================================================================*/
bool CAdoParameter::SetName(CString strName)
{
	assert(m_pParameter != NULL);
	try
	{
		m_pParameter->PutName(_bstr_t(LPCTSTR(strName)));
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetName 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}	
}

CString CAdoParameter::GetName()
{
	assert(m_pParameter != NULL);
	try
	{
		return CString(LPCTSTR(m_pParameter->GetName()));
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetName 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return CString(_T(""));
	}
}

/*========================================================================
	Name:		指示 Parameter 所标明的是输入参数、输出参数还是既是输出又
		是输入参数，或该参数是否为存储过程返回的值。
    ----------------------------------------------------------
	Params:		[Direction]: 设置以下某个 ParameterDirectionEnum 值。
		[常量]				[说明] 
		-------------------------------------------
		AdParamUnknown		指示参数方向未知。 
		AdParamInput		默认值。指示输入参数。 
		AdParamOutput		指示输出参数。 
		AdParamInputOutput	同时指示输入参数和输出参数。 
		AdParamReturnValue	指示返回值。 
==========================================================================*/
bool CAdoParameter::SetDirection(ParameterDirectionEnum Direction)
{
	assert(m_pParameter != NULL);
	try
	{
		m_pParameter->PutDirection(Direction);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetDirection 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}	
}

ParameterDirectionEnum CAdoParameter::GetDirection()
{
	assert(m_pParameter != NULL);
	try
	{
		return m_pParameter->GetDirection();
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetDirection 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return adParamUnknown;
	}	
}

/*########################################################################
			  ------------------------------------------------
						
			  ------------------------------------------------
  ########################################################################*/

bool CAdoParameter::SetValue(const  _variant_t &value)
{
	assert(m_pParameter != NULL);

	try
	{
		if (m_pParameter->Size == 0)
		{
			m_pParameter->Size = sizeof(VARIANT);
		}
		m_pParameter->Value = value;
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::SetValue(const bool &value)
{
	try
	{
		if (m_pParameter->Size == 0)
		{
			m_pParameter->Size = sizeof(short);
		}
		m_pParameter->Value = _variant_t(value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::SetValue(const int &value)
{
	try
	{
		if (m_pParameter->Size == 0)
		{
			m_pParameter->Size = sizeof(int);
		}
		m_pParameter->Value = _variant_t(long(value));
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::SetValue(const long &value)
{
	try
	{
		if (m_pParameter->Size == 0)
		{
			m_pParameter->Size = sizeof(long);
		}
		m_pParameter->Value = _variant_t(value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::SetValue(const double &value)
{
	try
	{
		if (m_pParameter->Size == 0)
		{
			m_pParameter->Size = sizeof(double);
		}
		m_pParameter->Value = _variant_t(value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::SetValue(const CString &value)
{
	_variant_t var;
	var.vt = value.IsEmpty() ? VT_NULL : VT_BSTR;
	var.bstrVal = value.AllocSysString();

	try
	{
		if (m_pParameter->Size == 0)
		{
			m_pParameter->Size = value.GetLength();
		}
		m_pParameter->Value = var;
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::SetValue(const COleDateTime &value)
{
	_variant_t var;
	var.vt = VT_DATE;
	var.date = value;
	
	try
	{
		if (m_pParameter->Size == 0)
		{
			m_pParameter->Size = sizeof(DATE);
		}
		m_pParameter->Value = var;
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
	return true;
}

bool CAdoParameter::SetValue(const BYTE &value)
{
	try
	{
		if (m_pParameter->Size == 0)
		{
			m_pParameter->Size = sizeof(BYTE);
		}
		m_pParameter->Value = _variant_t(value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::SetValue(const short &value)
{
	try
	{
		if (m_pParameter->Size == 0)
		{
			m_pParameter->Size = sizeof(short);
		}
		m_pParameter->Value = _variant_t(value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::SetValue(const float &value)
{
	try
	{
		if (m_pParameter->Size == 0)
		{
			m_pParameter->Size = sizeof(float);
		}
		m_pParameter->Value = _variant_t(value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: SetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}


bool CAdoParameter::GetValue(bool &value)
{
	try
	{
		value = Var2bool(m_pParameter->Value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::GetValue(BYTE &value)
{
	try
	{
		value = Var2Byte(m_pParameter->Value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::GetValue(short &value)
{
	try
	{
		value = Var2short(m_pParameter->Value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::GetValue(int &value)
{
	try
	{
		value = (int)Var2long(m_pParameter->Value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::GetValue(long &value)
{
	try
	{
		value = Var2long(m_pParameter->Value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::GetValue(double &value)
{
	try
	{
		value = Var2double(m_pParameter->Value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::GetValue(CString &value)
{
	try
	{
		value = Var2CString(m_pParameter->Value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

bool CAdoParameter::GetValue(COleDateTime &value)
{
	try
	{
		value = Var2DateTime(m_pParameter->Value);
		return true;
	}
	catch (_com_error e)
	{
		TRACE(_T("Warning: GetValue 方法发生异常. 错误信息: %s; 文件: %s; 行: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}
