#if !defined(_ANYOU_COOL_ADORECORDSET_H)
#define _ANYOU_COOL_ADORECORDSET_H
#include "Ado.h"

class CAdoRecordSet 
{
public:
	CAdoRecordSet();
	CAdoRecordSet(CAdoConnection *pConnection);
	virtual ~CAdoRecordSet() {
		if (IsOpen()) Close();
		m_pRecordset.Release();
		m_pRecordset = NULL;
	};
	
public:
	// ��ǰ�༭״̬ ----------------------------
	EditModeEnum GetEditMode();
	
	// ��ǰ״̬ --------------------------------
	bool IsEOF();
	bool IsBOF();
	bool IsOpen();
	long GetState();
	long GetStatus();

	// ��¼�����ֶ��� ----------------------
	long GetRecordCount();
	long GetFieldsCount();

	// �����ص�����¼�� --------------------
	long GetMaxRecordCount();
	bool SetMaxRecordCount(long count);

	// ҳ�� --------------------------------
	long GetPageCount();
	// ÿҳ�ļ�¼�� ------------------------
	long GetPageSize();
	bool SetCacheSize(const long& lCacheSize);	
	long GetAbsolutePage();
	bool SetAbsolutePage(int nPage);
	
	// ���λ�� --------------------------------
	CursorLocationEnum GetCursorLocation();
	bool SetCursorLocation(CursorLocationEnum CursorLocation = adUseClient);
	
	// ������� --------------------------------
	CursorTypeEnum GetCursorType();
	bool SetCursorType(CursorTypeEnum CursorType = adOpenStatic);
	
	// ��ǩ --------------------------------
	_variant_t GetBookmark();
	bool SetBookmark(_variant_t varBookMark = _variant_t((long)adBookmarkFirst));
	
	// ��ǰ��¼λ�� ------------------------
	long GetAbsolutePosition();
	bool SetAbsolutePosition(int nPosition);
	
	// ��ѯ�ַ��� --------------------------
	CString	GetQueryString() {return m_strSQL;}
	void	SetQueryString(LPCTSTR strSQL) {m_strSQL = strSQL;}
	
	// ���Ӷ��� ----------------------------
	CAdoConnection* GetConnection() {return m_pConnection;}
	void SetAdoConnection(CAdoConnection *pConnection);

	// ��¼������ --------------------------
	_RecordsetPtr& GetRecordset();

	CString GetErrorInfo();

// �ֶ����� ----------------------------------------------
public:
	// �ֶμ� -------------------------------
	FieldsPtr GetFields();

	// �ֶζ��� -----------------------------
	FieldPtr  GetField(long lIndex);
	FieldPtr  GetField(LPCTSTR lpszFieldName);
	
	// �ֶ��� -------------------------------
	CString	  GetFieldName(long lIndex);
	
	// �ֶ��������� -------------------------
	DataTypeEnum GetFieldType(long lIndex);
	DataTypeEnum GetFieldType(LPCTSTR lpszFieldName);

	// �ֶ����� -----------------------------
	long  GetFieldAttributes(long lIndex);
	long  GetFieldAttributes(LPCTSTR lpszFieldName);

	// �ֶζ��峤�� -------------------------
	long  GetFieldDefineSize(long lIndex);
	long  GetFieldDefineSize(LPCTSTR lpszFieldName);

	// �ֶ�ʵ�ʳ��� -------------------------
	long  GetFieldActualSize(long lIndex);
	long  GetFieldActualSize(LPCTSTR lpszFieldName);

	// �ֶ��Ƿ�ΪNULL -----------------------
	bool  IsFieldNull(long index);
	bool  IsFieldNull(LPCTSTR lpFieldName);

// ��¼���� --------------------------------------------
public:
	bool AddNew();
	bool Update();
	bool UpdateBatch(AffectEnum AffectRecords = adAffectAll); 
	bool CancelUpdate();
	bool CancelBatch(AffectEnum AffectRecords = adAffectAll);
	bool Delete(AffectEnum AffectRecords = adAffectCurrent);
	
	// ˢ�¼�¼���е����� ------------------
	bool Requery(long Options = adConnectUnspecified);
	bool Resync(AffectEnum AffectRecords = adAffectAll, ResyncEnum ResyncValues = adResyncAllValues);   

	bool RecordBinding(CADORecordBinding &pAdoRecordBinding);
	bool AddNew(CADORecordBinding &pAdoRecordBinding);
	
// ��¼���������� --------------------------------------
public:
	bool MoveFirst();
	bool MovePrevious();
	bool MoveNext();
	bool MoveLast();
	bool Move(long lRecords, _variant_t Start = _variant_t((long)adBookmarkFirst));
	
	// ����ָ���ļ�¼ ----------------------
	bool Find(LPCTSTR lpszFind, SearchDirectionEnum SearchDirection = adSearchForward);
	bool FindNext();

// ��ѯ ------------------------------------------------
public:
	bool Open(LPCTSTR strSQL, long lOption = adCmdText, CursorTypeEnum CursorType = adOpenStatic, LockTypeEnum LockType = adLockOptimistic);
	bool Cancel();
	void Close();

	// ����/����־����ļ� -----------------
	bool Save(LPCTSTR strFileName = _T(""), PersistFormatEnum PersistFormat = adPersistXML);
	bool Load(LPCTSTR strFileName);
	
// �ֶδ�ȡ --------------------------------------------
public:
	bool PutCollect(long index, const _variant_t &value);
	bool PutCollect(long index, const CString &value);
	bool PutCollect(long index, const double &value);
	bool PutCollect(long index, const float  &value);
	bool PutCollect(long index, const long   &value);
	bool PutCollect(long index, const DWORD  &value);
	bool PutCollect(long index, const int    &value);
	bool PutCollect(long index, const short  &value);
	bool PutCollect(long index, const BYTE   &value);
	bool PutCollect(long index, const bool   &value);
	bool PutCollect(long index, const COleDateTime &value);
	bool PutCollect(long index, const COleCurrency &value);

	bool PutCollect(LPCTSTR strFieldName, const _variant_t &value);
	bool PutCollect(LPCTSTR strFieldName, const CString &value);
	bool PutCollect(LPCTSTR strFieldName, const double &value);
	bool PutCollect(LPCTSTR strFieldName, const float  &value);
	bool PutCollect(LPCTSTR strFieldName, const long   &value);
	bool PutCollect(LPCTSTR strFieldName, const DWORD  &value);
	bool PutCollect(LPCTSTR strFieldName, const int    &value);
	bool PutCollect(LPCTSTR strFieldName, const short  &value);
	bool PutCollect(LPCTSTR strFieldName, const BYTE   &value);
	bool PutCollect(LPCTSTR strFieldName, const bool   &value);
	bool PutCollect(LPCTSTR strFieldName, const COleDateTime &value);
	bool PutCollect(LPCTSTR strFieldName, const COleCurrency &value);

	// ---------------------------------------------------------

	bool GetCollect(long index, CString &value);
	bool GetCollect(long index, double  &value);
	bool GetCollect(long index, float   &value);
	bool GetCollect(long index, long    &value);
	bool GetCollect(long index, DWORD   &value);
	bool GetCollect(long index, int     &value);
	bool GetCollect(long index, short   &value);
	bool GetCollect(long index, BYTE    &value);
	bool GetCollect(long index, bool   &value);
	bool GetCollect(long index, COleDateTime &value);
	bool GetCollect(long index, COleCurrency &value);

	bool GetCollect(LPCSTR strFieldName, CString &strValue);
	bool GetCollect(LPCSTR strFieldName, double &value);
	bool GetCollect(LPCSTR strFieldName, float  &value);
	bool GetCollect(LPCSTR strFieldName, long   &value);
	bool GetCollect(LPCSTR strFieldName, DWORD  &value);
	bool GetCollect(LPCSTR strFieldName, int    &value);
	bool GetCollect(LPCSTR strFieldName, short  &value);
	bool GetCollect(LPCSTR strFieldName, BYTE   &value);
	bool GetCollect(LPCSTR strFieldName, bool   &value);
	bool GetCollect(LPCSTR strFieldName, COleDateTime &value);
	bool GetCollect(LPCSTR strFieldName, COleCurrency &value);

	// BLOB ���ݴ�ȡ ------------------------------------------
	bool AppendChunk(FieldPtr pField, LPVOID lpData, UINT nBytes);
	bool AppendChunk(long index, LPVOID lpData, UINT nBytes);
	bool AppendChunk(LPCSTR strFieldName, LPVOID lpData, UINT nBytes);
	bool AppendChunk(long index, LPCTSTR lpszFileName);
	bool AppendChunk(LPCSTR strFieldName, LPCTSTR lpszFileName);

	bool GetChunk(FieldPtr pField, LPVOID lpData);
	bool GetChunk(long index, LPVOID lpData);
	bool GetChunk(LPCSTR strFieldName, LPVOID lpData);
	bool GetChunk(long index, CBitmap &bitmap);
	bool GetChunk(LPCSTR strFieldName, CBitmap &bitmap);

// �������� --------------------------------------------
public:
	// ���˼�¼ ---------------------------------
	bool SetFilter(LPCTSTR lpszFilter);

	// ���� -------------------------------------
	bool SetSort(LPCTSTR lpszCriteria);

	// �����Ƿ�֧��ĳ���� -----------------------
	bool Supports(CursorOptionEnum CursorOptions = adAddNew);

	// ��¡ -------------------------------------
	bool Clone(CAdoRecordSet &pRecordSet);
	_RecordsetPtr operator = (_RecordsetPtr &pRecordSet);
	
	// ��ʽ�� _variant_t ����ֵ -----------------
	
//��Ա���� --------------------------------------------
protected:
	CAdoConnection     *m_pConnection;
	_RecordsetPtr		m_pRecordset;
	CString				m_strSQL;
	CString				m_strFind;
	CString				m_strFileName;
	IADORecordBinding	*m_pAdoRecordBinding;
	SearchDirectionEnum m_SearchDirection;
public:
	_variant_t			m_varBookmark;
};
//________________________________________________________________________

#endif //_ANYOU_COOL_ADORECORDSET_H