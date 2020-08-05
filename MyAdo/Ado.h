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

// 数值类型转换(内部使用) -----------------------------------
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
	// 连接对象 ----------------------------------
	_ConnectionPtr& GetConnection()
	{
		return m_pConnection;
	};

	// 连接信息 ----------------------------------
	CString	GetProviderName();
	CString	GetVersion();
	CString	GetDefaultDatabase();

	// 异常信息 ----------------------------------
	CString	GetErrorInfo();

	// 连接字串 ----------------------------------
	CString GetConnectionString()
	{
		return m_strConnect;
	}

	// 连接状态 ----------------------------------
	// 	adStateClosed = 0, 	adStateOpen = 1,
	// 	adStateConnecting = 2, 	adStateExecuting = 4, 	adStateFetching = 8
	bool			IsOpen();
	long			GetState();

	// 连接模式 ----------------------------------
	ConnectModeEnum	GetMode();
	bool			SetMode(ConnectModeEnum mode);

	// 连接时间 ----------------------------------
	long			GetConnectTimeOut();
	bool			SetConnectTimeOut(long lTime = 5);

	// 数据源信息 -------------------------------
	_RecordsetPtr	OpenSchema(SchemaEnum QueryType);

	// 操作 -----------------------------------------------
public:
	// 数据库连接 --------------------------------
	bool			Open(LPCTSTR lpszConnect, long lOptions = adConnectUnspecified);
	bool			ConnectSQLServer(CString dbsrc, CString dbname, CString user, CString pass, long lOptions = adConnectUnspecified);
	bool			ConnectAccess(CString dbpath, CString pass = _T(""), long lOptions = adConnectUnspecified);
	bool			OpenUDLFile(LPCTSTR strFileName, long lOptions = adConnectUnspecified);
	void			Close();

	// 处理 -----------------------------------------------
public:
	// 事务处理 ----------------------------------
	long			BeginTrans();
	bool			RollbackTrans();
	bool			CommitTrans();

	// 执行 SQL 语句 ------------------------------
	int				Execute(LPCTSTR strSQL, long lOptions = adExecuteNoRecords);
	bool			Cancel();

	// 数据 -----------------------------------------------
protected:
	CString			m_strConnect;
	_ConnectionPtr	m_pConnection;
};

#endif // !defined(_ANYOU_COOL_ADO_H)










































