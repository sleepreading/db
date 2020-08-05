#if !defined(_ANYOU_COOL_ADOCOMMAND_H)
#define _ANYOU_COOL_ADOCOMMAND_H

#include "Ado.h"

class CAdoParameter
{
public:
	CAdoParameter();
	CAdoParameter(DataTypeEnum DataType, long lSize, ParameterDirectionEnum Direction, CString strName);
	virtual ~CAdoParameter();

public:
	_ParameterPtr GetParameter() {return m_pParameter;}

	// 参数精度 ---------------------------
	bool SetPrecision(char nPrecision);

	// 参数小数位数 -----------------------
	bool SetNumericScale(int nScale);

	// 参数类型 ---------------------------
	ParameterDirectionEnum GetDirection();
	bool SetDirection(ParameterDirectionEnum Direction);
	
	// 参数名称 ---------------------------
	CString	GetName();
	bool	SetName(CString strName);

	// 参数长度 ---------------------------
	int		GetSize();
	bool	SetSize(int size);

	// 参数据类型 -------------------------
	DataTypeEnum GetType();
	bool SetType(DataTypeEnum DataType);

// 方法 ------------------------------------------------
public:	
	bool GetValue(COleDateTime &value);
	bool GetValue(CString &value);
	bool GetValue(double &value);
	bool GetValue(long &value);
	bool GetValue(int &value);
	bool GetValue(short &value);
	bool GetValue(BYTE &value);
	bool GetValue(bool &value);

	bool SetValue(const float &value);
	bool SetValue(const short &value);
	bool SetValue(const BYTE &value);
	bool SetValue(const COleDateTime &value);
	bool SetValue(const CString &value);
	bool SetValue(const double &value);
	bool SetValue(const long &value);
	bool SetValue(const int &value);
	bool SetValue(const bool &value);
	bool SetValue(const _variant_t &value);

//其他方法 ------------------------------
public:
	_ParameterPtr& operator =(_ParameterPtr& pParameter);

// 数据 ------------------------------------------------
protected:
	_ParameterPtr m_pParameter;
	CString m_strName;
	DataTypeEnum m_nType;
};


class CAdoCommand
{
public:
	CAdoCommand();
	CAdoCommand(CAdoConnection* pAdoConnection, CString strCommandText = _T(""), CommandTypeEnum CommandType = adCmdStoredProc);
	virtual ~CAdoCommand();

public:
	_variant_t GetValue(LPCTSTR lpstrName);
	_variant_t GetValue(long index);

	_ParameterPtr GetParameter(long index);
	_ParameterPtr GetParamter(LPCTSTR lpstrName);

	bool Append(_ParameterPtr param);
	ParametersPtr GetParameters();
	bool SetCommandTimeOut(long lTime);
	long GetState();

	bool SetCommandType(CommandTypeEnum CommandType);
	bool SetCommandText(LPCTSTR lpstrCommand);

	CAdoParameter operator [](int index);
	CAdoParameter operator [](LPCTSTR lpszParamName);

// 实现 ------------------------------------------------
public:
	_ParameterPtr CreateParameter(LPCTSTR lpstrName, DataTypeEnum Type, ParameterDirectionEnum Direction, 
								  long Size, _variant_t Value);
	_RecordsetPtr CAdoCommand::Execute(long Options = adCmdStoredProc);
	bool Cancel();
	
// 其他方法 --------------------------------------------
public:
	_CommandPtr& GetCommand();
	bool SetConnection(CAdoConnection *pConnect);

// 数据 ------------------------------------------------
protected:
	void Release();
	_CommandPtr		m_pCommand;
	CString			m_strSQL;
};

#endif // !defined(_ANYOU_COOL_ADOCOMMAND_H)
