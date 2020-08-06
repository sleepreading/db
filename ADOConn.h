#pragma once

#include <string>

#import "C:/Program Files/Common Files/System/ado/msado15.dll" no_namespace rename("EOF","adoEOF") rename("BOF","adoBOF")

class ADOConn
{
public:
  ADOConn();
  virtual ~ADOConn();

public:
  _ConnectionPtr& GetConnPtr() { return m_pConnection; }
  _RecordsetPtr& GetRecoPtr() { return m_pRecordset; }

public:
  bool RollbackTrans();
  bool CommitTrans();
  bool BeginTrans();

  bool adoEOF();
  bool MoveNext();

  bool CloseTable();
  bool CloseADOConnection();

  bool GetCollect(const std::string& Name,std::string &lpDest);
  bool ExecuteSQL(const std::string &lpszSQL); // update delete insert
  bool OnInitADOConn(const std::string& ConnStr);
  _RecordsetPtr& GetRecordSet(const std::string& lpszSQL);  // select
  int GetRecordCount();

public:
  _ConnectionPtr m_pConnection;
  _RecordsetPtr m_pRecordset;
};

