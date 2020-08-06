#include "ADOConn.h"

using namespace std;

#pragma warning( disable : 4996 )

ADOConn::ADOConn()
{
	::CoInitialize(NULL);
}

ADOConn::~ADOConn()
{
	::CoUninitialize();
}

bool ADOConn::OnInitADOConn(const std::string& ConnStr)
{
	try
	{
		m_pConnection.CreateInstance("ADODB.Connection");
		_bstr_t strConnect = _bstr_t(ConnStr.c_str());
		return m_pConnection->Open(strConnect,"","",adModeUnknown) == S_OK ? true : false;
	}
	catch(_com_error e)
	{
		return false;
	}

	return true;
}

bool ADOConn::ExecuteSQL(const std::string& lpszSQL)
{
	try
	{
		m_pConnection->Execute(_bstr_t(lpszSQL.c_str()), NULL, adCmdText);
		return true;
	}
	catch(_com_error e)
	{
		return false;
	}

	return true;
}

_RecordsetPtr& ADOConn::GetRecordSet(const std::string& lpszSQL)
{
	try
	{
		m_pRecordset.CreateInstance(__uuidof(Recordset));
		m_pRecordset->Open(_bstr_t(lpszSQL.c_str()), m_pConnection.GetInterfacePtr(), adOpenForwardOnly, adLockReadOnly, adCmdText);
	}
	catch(_com_error e)
	{
	}

	return m_pRecordset;
}

bool ADOConn::GetCollect(const std::string& Name,std::string &lpDest)
{
	VARIANT  vt;
	try
	{
		vt = m_pRecordset->GetCollect(Name.c_str());
		if(vt.vt!=VT_NULL)
			lpDest=(LPCSTR)_bstr_t(vt);
		else
			lpDest="";
	}
	catch (_com_error e)
	{
		return false;
	}

	return true;
}

bool ADOConn::CloseADOConnection()
{
	try
	{
		m_pConnection->Close();
	}
	catch (_com_error e)
	{
		return false;
	}

	return true;
}

bool ADOConn::CloseTable()
{
	try
	{
		m_pRecordset->Close();
	}
	catch (_com_error e)
	{
		return false;
	}

	return true;
}

bool ADOConn::MoveNext()
{
	try
	{
		m_pRecordset->MoveNext();
	}
	catch (_com_error e)
	{
		return false;
	}

	return true;
}

bool ADOConn::adoEOF()
{
	try
	{
		return m_pRecordset->adoEOF? true : false;
	}
	catch (_com_error e)
	{
		return true;
	}
}

bool ADOConn::BeginTrans()
{
	try
	{
		m_pConnection->BeginTrans();
	}
	catch (_com_error e) 
	{
		return false;
	} 

	return true;
}

bool ADOConn::CommitTrans()
{
	try
	{
		m_pConnection->CommitTrans();
	}
	catch (_com_error e)
	{
		return false;
	}

	return true;
}

bool ADOConn::RollbackTrans()
{
	try
	{
		m_pConnection->RollbackTrans();
	}
	catch (_com_error e)
	{
		return false;
	}

	return true;
}

int ADOConn::GetRecordCount()
{
	DWORD nRows = 0; 

	try
	{
		nRows = (DWORD)(m_pRecordset->GetRecordCount()); 

		if(nRows == -1) 
		{ 
			nRows = 0; 
			if(m_pRecordset->adoEOF != VARIANT_TRUE) 
				m_pRecordset->MoveFirst(); 

			while(m_pRecordset->adoEOF != VARIANT_TRUE) 
			{ 
				nRows++; 
				m_pRecordset->MoveNext(); 
			} 
			if(nRows > 0) 
				m_pRecordset->MoveFirst(); 
		} 
	}
	catch (_com_error e)
	{
	}

	return nRows; 
}
