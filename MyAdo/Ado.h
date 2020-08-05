#if !defined(_ANYOU_COOL_ADO_H)
#define _ANYOU_COOL_ADO_H

#ifdef _WINDOWS
#include <afx.h>	 //use afxdisp.h should include this file first
#include <afxdisp.h> //COleDateTime
#endif
#include <icrsint.h> //ADO C/C++ Record Binding Definitions
#include <tchar.h>
#include <assert.h>

#pragma warning (disable: 4146 4800)
#import "msado15.dll" rename("EOF","adoEOF"), rename("BOF","adoBOF")
using namespace ADODB;
#pragma warning (default: 4146)

class CAdoConnection;
#include "AdoRecordSet.h"
#include "AdoCommand.h"

// ��ֵ����ת��(�ڲ�ʹ��) -----------------------------------
COleDateTime	Var2DateTime(const _variant_t& var);
COleCurrency	Var2Currency(const _variant_t& var);
bool			Var2bool(const _variant_t& var);
BYTE			Var2Byte(const _variant_t& var);
short			Var2short(const _variant_t& var);
long			Var2long(const _variant_t& var);
double			Var2double(const _variant_t& var);
CString			Var2CString(const _variant_t& var);

class CAdoConnection
{
public:
	CAdoConnection()
	{
		::CoInitialize(NULL);		
		m_pConnection.CreateInstance(__uuidof(Connection));
		assert(m_pConnection != NULL);
	};
	virtual ~CAdoConnection()
	{
		Close();
		m_pConnection.Release();
		m_pConnection = NULL;
		::CoUninitialize();
	};

public:
	// ���Ӷ��� ----------------------------------
	_ConnectionPtr& GetConnection()
	{
		return m_pConnection;
	};

	// ������Ϣ ----------------------------------
	CString	GetProviderName();
	CString	GetVersion();
	CString	GetDefaultDatabase();

	// �쳣��Ϣ ----------------------------------
	CString	GetErrorInfo();

	// �����ִ� ----------------------------------
	CString GetConnectionString()
	{
		return m_strConnect;
	}

	// ����״̬ ----------------------------------
	// 	adStateClosed = 0, 	adStateOpen = 1,
	// 	adStateConnecting = 2, 	adStateExecuting = 4, 	adStateFetching = 8
	bool			IsOpen();
	long			GetState();

	// ����ģʽ ----------------------------------
	ConnectModeEnum	GetMode();
	bool			SetMode(ConnectModeEnum mode);

	// ����ʱ�� ----------------------------------
	long			GetConnectTimeOut();
	bool			SetConnectTimeOut(long lTime = 5);

	// ����Դ��Ϣ -------------------------------
	_RecordsetPtr	OpenSchema(SchemaEnum QueryType);

	// ���� -----------------------------------------------
public:
	// ���ݿ����� --------------------------------
	bool			Open(LPCTSTR lpszConnect, long lOptions = adConnectUnspecified);
	bool			ConnectSQLServer(CString dbsrc, CString dbname, CString user, CString pass, long lOptions = adConnectUnspecified);
	bool			ConnectAccess(CString dbpath, CString pass = _T(""), long lOptions = adConnectUnspecified);
	bool			OpenUDLFile(LPCTSTR strFileName, long lOptions = adConnectUnspecified);
	void			Close();

	// ���� -----------------------------------------------
public:
	// ������ ----------------------------------
	long			BeginTrans();
	bool			RollbackTrans();
	bool			CommitTrans();

	// ִ�� SQL ��� ------------------------------
	int				Execute(LPCTSTR strSQL, long lOptions = adExecuteNoRecords);
	bool			Cancel();

	// ���� -----------------------------------------------
protected:
	CString			m_strConnect;
	_ConnectionPtr	m_pConnection;
};

#endif // !defined(_ANYOU_COOL_ADO_H)










































