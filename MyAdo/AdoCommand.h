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

	// �������� ---------------------------
	bool SetPrecision(char nPrecision);

	// ����С��λ�� -----------------------
	bool SetNumericScale(int nScale);

	// �������� ---------------------------
	ParameterDirectionEnum GetDirection();
	bool SetDirection(ParameterDirectionEnum Direction);
	
	// �������� ---------------------------
	CString	GetName();
	bool	SetName(CString strName);

	// �������� ---------------------------
	int		GetSize();
	bool	SetSize(int size);

	// ���������� -------------------------
	DataTypeEnum GetType();
	bool SetType(DataTypeEnum DataType);

// ���� ------------------------------------------------
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

//�������� ------------------------------
public:
	_ParameterPtr& operator =(_ParameterPtr& pParameter);

// ���� ------------------------------------------------
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

// ʵ�� ------------------------------------------------
public:
	_ParameterPtr CreateParameter(LPCTSTR lpstrName, DataTypeEnum Type, ParameterDirectionEnum Direction, 
								  long Size, _variant_t Value);
	_RecordsetPtr CAdoCommand::Execute(long Options = adCmdStoredProc);
	bool Cancel();
	
// �������� --------------------------------------------
public:
	_CommandPtr& GetCommand();
	bool SetConnection(CAdoConnection *pConnect);

// ���� ------------------------------------------------
protected:
	void Release();
	_CommandPtr		m_pCommand;
	CString			m_strSQL;
};

#endif // !defined(_ANYOU_COOL_ADOCOMMAND_H)
