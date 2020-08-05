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
	// 当前编辑状态 ----------------------------
	EditModeEnum GetEditMode();
	
	// 当前状态 --------------------------------
	bool IsEOF();
	bool IsBOF();
	bool IsOpen();
	long GetState();
	long GetStatus();

	// 记录数及字段数 ----------------------
	long GetRecordCount();
	long GetFieldsCount();

	// 充许返回的最大记录数 --------------------
	long GetMaxRecordCount();
	bool SetMaxRecordCount(long count);

	// 页数 --------------------------------
	long GetPageCount();
	// 每页的记录数 ------------------------
	long GetPageSize();
	bool SetCacheSize(const long& lCacheSize);	
	long GetAbsolutePage();
	bool SetAbsolutePage(int nPage);
	
	// 光标位置 --------------------------------
	CursorLocationEnum GetCursorLocation();
	bool SetCursorLocation(CursorLocationEnum CursorLocation = adUseClient);
	
	// 光标类型 --------------------------------
	CursorTypeEnum GetCursorType();
	bool SetCursorType(CursorTypeEnum CursorType = adOpenStatic);
	
	// 书签 --------------------------------
	_variant_t GetBookmark();
	bool SetBookmark(_variant_t varBookMark = _variant_t((long)adBookmarkFirst));
	
	// 当前记录位置 ------------------------
	long GetAbsolutePosition();
	bool SetAbsolutePosition(int nPosition);
	
	// 查询字符串 --------------------------
	CString	GetQueryString() {return m_strSQL;}
	void	SetQueryString(LPCTSTR strSQL) {m_strSQL = strSQL;}
	
	// 连接对象 ----------------------------
	CAdoConnection* GetConnection() {return m_pConnection;}
	void SetAdoConnection(CAdoConnection *pConnection);

	// 记录集对象 --------------------------
	_RecordsetPtr& GetRecordset();

	CString GetErrorInfo();

// 字段属性 ----------------------------------------------
public:
	// 字段集 -------------------------------
	FieldsPtr GetFields();

	// 字段对象 -----------------------------
	FieldPtr  GetField(long lIndex);
	FieldPtr  GetField(LPCTSTR lpszFieldName);
	
	// 字段名 -------------------------------
	CString	  GetFieldName(long lIndex);
	
	// 字段数据类型 -------------------------
	DataTypeEnum GetFieldType(long lIndex);
	DataTypeEnum GetFieldType(LPCTSTR lpszFieldName);

	// 字段属性 -----------------------------
	long  GetFieldAttributes(long lIndex);
	long  GetFieldAttributes(LPCTSTR lpszFieldName);

	// 字段定义长度 -------------------------
	long  GetFieldDefineSize(long lIndex);
	long  GetFieldDefineSize(LPCTSTR lpszFieldName);

	// 字段实际长度 -------------------------
	long  GetFieldActualSize(long lIndex);
	long  GetFieldActualSize(LPCTSTR lpszFieldName);

	// 字段是否为NULL -----------------------
	bool  IsFieldNull(long index);
	bool  IsFieldNull(LPCTSTR lpFieldName);

// 记录更改 --------------------------------------------
public:
	bool AddNew();
	bool Update();
	bool UpdateBatch(AffectEnum AffectRecords = adAffectAll); 
	bool CancelUpdate();
	bool CancelBatch(AffectEnum AffectRecords = adAffectAll);
	bool Delete(AffectEnum AffectRecords = adAffectCurrent);
	
	// 刷新记录集中的数据 ------------------
	bool Requery(long Options = adConnectUnspecified);
	bool Resync(AffectEnum AffectRecords = adAffectAll, ResyncEnum ResyncValues = adResyncAllValues);   

	bool RecordBinding(CADORecordBinding &pAdoRecordBinding);
	bool AddNew(CADORecordBinding &pAdoRecordBinding);
	
// 记录集导航操作 --------------------------------------
public:
	bool MoveFirst();
	bool MovePrevious();
	bool MoveNext();
	bool MoveLast();
	bool Move(long lRecords, _variant_t Start = _variant_t((long)adBookmarkFirst));
	
	// 查找指定的记录 ----------------------
	bool Find(LPCTSTR lpszFind, SearchDirectionEnum SearchDirection = adSearchForward);
	bool FindNext();

// 查询 ------------------------------------------------
public:
	bool Open(LPCTSTR strSQL, long lOption = adCmdText, CursorTypeEnum CursorType = adOpenStatic, LockTypeEnum LockType = adLockOptimistic);
	bool Cancel();
	void Close();

	// 保存/载入持久性文件 -----------------
	bool Save(LPCTSTR strFileName = _T(""), PersistFormatEnum PersistFormat = adPersistXML);
	bool Load(LPCTSTR strFileName);
	
// 字段存取 --------------------------------------------
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

	// BLOB 数据存取 ------------------------------------------
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

// 其他方法 --------------------------------------------
public:
	// 过滤记录 ---------------------------------
	bool SetFilter(LPCTSTR lpszFilter);

	// 排序 -------------------------------------
	bool SetSort(LPCTSTR lpszCriteria);

	// 测试是否支持某方法 -----------------------
	bool Supports(CursorOptionEnum CursorOptions = adAddNew);

	// 克隆 -------------------------------------
	bool Clone(CAdoRecordSet &pRecordSet);
	_RecordsetPtr operator = (_RecordsetPtr &pRecordSet);
	
	// 格式化 _variant_t 类型值 -----------------
	
//成员变量 --------------------------------------------
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