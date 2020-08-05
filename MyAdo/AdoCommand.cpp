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
	///���� Connection ����---------------------------
	m_pCommand.CreateInstance(__uuidof(Command));
	#ifdef _DEBUG
	if (m_pCommand == NULL)
	{
		AfxMessageBox("Command ���󴴽�ʧ��! ��ȷ���Ƿ��ʼ����Com����.");
	}
	#endif
	assert(m_pCommand != NULL);
}

CAdoCommand::CAdoCommand(CAdoConnection* pAdoConnection, CString strCommandText, CommandTypeEnum CommandType)
{
	///���� Connection ����---------------------------
	m_pCommand.CreateInstance(__uuidof(Command)); //"ADODB.Command"
	#ifdef _DEBUG
	if (m_pCommand == NULL)
	{
		AfxMessageBox("Command ���󴴽�ʧ��! ��ȷ���Ƿ��ʼ����Com����.");
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
		TRACE(_T("Warning: Release���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: Execute ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: Cancel ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: CreateParameter ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: PutCommandText ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetConnection ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	} 
}
/*========================================================================
	Name:		ָʾ Command ��������͡�
    ----------------------------------------------------------
	returns:	��������ĳ�� CommandTypeEnum ��ֵ.
		[����]				 [˵��] 
		----------------------------------
		adCmdText			ָʾstrSQLΪ�����ı�, ����ͨ��SQL���. 
		adCmdTable			ָʾADO����SQL��ѯ������ strSQL �����ı��е�
						������. 
		adCmdTableDirect	ָʾ�����ĸ�����strSQL�������ı��з���������. 
		adCmdStoredProc		ָʾstrSQLΪ�洢����. 
		adCmdUnknown		ָʾstrSQL�����е���������Ϊδ֪. 
		adCmdFile			ָʾӦ����strSQL���������ļ��лָ�����(�����)
						Recordset. 
		adAsyncExecute		ָʾӦ�첽ִ��strSQL. 
		adAsyncFetch		ָʾ����ȡ Initial Fetch Size ������ָ���ĳ�ʼ
						������, Ӧ���첽��ȡ����ʣ�����. ������������δ
						��ȡ, ��Ҫ���߳̽�������ֱ�������¿���. 
		adAsyncFetchNonBlocking ָʾ��Ҫ�߳�����ȡ�ڼ��δ����. ���������
						������δ��ȡ, ��ǰ���Զ��Ƶ��ļ�ĩβ. 
   ----------------------------------------------------------
	Remarks: ʹ�� CommandType ���Կ��Ż� CommandText ���Եļ��㡣
		��� CommandType ���Ե�ֵ���� adCmdUnknown(Ĭ��ֵ), ϵͳ�����ܽ���
	����, ��Ϊ ADO ��������ṩ����ȷ�� CommandText ������ SQL ��䡢���Ǵ�
	�����̻������ơ����֪������ʹ�õ����������, ��ͨ������ CommandType 
	����ָ�� ADO ֱ��ת����ش��롣��� CommandType ������ CommandText ��
	���е��������Ͳ�ƥ��, ���� Execute ����ʱ����������
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
		TRACE(_T("Warning: PutCommandType ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetState ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetCommandTimeOut ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetCommandTimeOut ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: Append ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetParamter ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetParamter ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: operator [] ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: operator [] ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		AfxMessageBox("Parameter ���󴴽�ʧ��! ��ȷ���Ƿ��ʼ����Com����.");
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
		AfxMessageBox("Parameter ���󴴽�ʧ��! ��ȷ���Ƿ��ʼ����Com����.");
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
	Name:		ָʾ�� Parameter ����������ֵ������ Field ����ľ��ȡ�
    ----------------------------------------------------------
	Params:		���û򷵻� Byte ֵ��������ʾֵ�����λ������ֵ�� Parameter
			������Ϊ��/д������ Field ������Ϊֻ����
    ----------------------------------------------------------
	Remarks:	ʹ�� Precision ���Կ�ȷ����ʾ���� Parameter �� Field ����ֵ
			�����λ��
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
		TRACE(_T("Warning: SetPrecision ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}

/*========================================================================
	Name:		ָ�� Parameter �� Field ����������ֵ�ķ�Χ��
    ----------------------------------------------------------
	Params:		���û򷵻��ֽ�ֵ��ָʾ����ֵ����ȷ����С����λ����
    ----------------------------------------------------------
	Remarks:	ʹ�� NumericScale ���Կ�ȷ�����ڱ��������� Parameter �� Field 
		�����ֵ��С��λ����
		���� Parameter ����NumericScale ����Ϊ��/д������ Field ����
	NumericScale ����Ϊֻ����

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
		TRACE(_T("Warning: SetPrecision ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}	
}

/*========================================================================
	Name:		ָʾ Parameter �������������.
    ----------------------------------------------------------
	Params:		[DataType]: DataTypeEnum ����ֵ, ��ο� CRecordSet �����
			����.
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
		TRACE(_T("Warning: SetType ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetDirection ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return adEmpty;
	}
}

/*========================================================================
	Name:		��ʾ Parameter ���������С�����ֽڻ��ַ�����
    ----------------------------------------------------------
	Params:		[size]: ��ʾ Parameter ���������С�����ֽڻ��ַ����ĳ�
			����ֵ��
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
		TRACE(_T("Warning: SetSize ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetDirection ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return -1;
	}
}

/*========================================================================
	Name:		ָʾ��������ơ�
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
		TRACE(_T("Warning: SetName ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetName ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return CString(_T(""));
	}
}

/*========================================================================
	Name:		ָʾ Parameter �����������������������������Ǽ��������
		�������������ò����Ƿ�Ϊ�洢���̷��ص�ֵ��
    ----------------------------------------------------------
	Params:		[Direction]: ��������ĳ�� ParameterDirectionEnum ֵ��
		[����]				[˵��] 
		-------------------------------------------
		AdParamUnknown		ָʾ��������δ֪�� 
		AdParamInput		Ĭ��ֵ��ָʾ��������� 
		AdParamOutput		ָʾ��������� 
		AdParamInputOutput	ͬʱָʾ������������������ 
		AdParamReturnValue	ָʾ����ֵ�� 
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
		TRACE(_T("Warning: SetDirection ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetDirection ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: SetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
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
		TRACE(_T("Warning: GetValue ���������쳣. ������Ϣ: %s; �ļ�: %s; ��: %d\n"), e.ErrorMessage(), __FILE__, __LINE__);
		return false;
	}
}
