#include "stdafx.h"
#include "adodata.h"
//#include "String"//Add

ADODATA::ADODATA()
{
	::CoInitialize(NULL);    //init
	m_pConnection.CreateInstance(_uuidof(Connection));
	m_pCommand.CreateInstance(_uuidof(Command));
	m_pRecordset.CreateInstance(_uuidof(Recordset));
	uninit_flag   = false;
	m_bstrConnStr = NULL;
}

ADODATA::ADODATA(char *serveName, char *databaseName, char *user, char *passWord)
{
	::CoInitialize(NULL);    //init
	m_pConnection.CreateInstance(_uuidof(Connection));
	m_pCommand.CreateInstance(_uuidof(Command));
	m_pRecordset.CreateInstance(_uuidof(Recordset));
	m_bstrConnStr = NULL;
	SetConnection(serveName,databaseName,user,passWord);
}

ADODATA::~ADODATA()
{
	if (m_pRecordset->State == 1)
	{
		m_pRecordset->Close();
	}
	if (m_pConnection->State == 1)
	{
		m_pConnection->Close();
	}

	m_pConnection.Release();
	m_pCommand.Release();
	m_pRecordset.Release();    

	SysFreeString(m_bstrConnStr);
	m_bstrConnStr = NULL;

	if(uninit_flag)
	{
		::CoUninitialize();
	}
}

bool ADODATA::SetConnection(char *serveName, char *databaseName, char *user, char *passWord)
{
	m_serveName    = serveName;
	m_databaseName = databaseName;
	m_user         = user;
	m_passWord     = passWord;
	SetConnStr();    

	if(m_pConnection->State == 1)
	{
		m_pConnection->Close(); 
	}

	try
	{
		m_pConnection->Open(m_bstrConnStr,user,passWord,0);//"UFCL2SERVER","neuralneuralneu",0);
	}
	catch(_com_error e)
	{
       #ifdef _DEBUG
		       CString errorMessage;
		       errorMessage.Format("%s", (LPTSTR)e.Description());
		//  AfxMessageBox(errorMessage);
       #endif
		return false;
	}
	uninit_flag = true;
	return true;
}    

void ADODATA::SetConnStr(void)
{
	CString strTemp;
	//SQL server
	//strTemp = TEXT("Provider = SQLOLEDB.1;");
	//ORACLE
	strTemp = TEXT("provider =OraOLEDB.Oracle;data source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST =")+ m_serveName;


	//if(m_passWord == "")
	//{
	//	strTemp += TEXT("Persist Security Info=False;User ID=") + m_user + TEXT(";Initial Catalog=");
	//}
	//else
	//{
	//	strTemp += TEXT("Password=") + m_passWord + TEXT(";Persist Security Info=True;User ID=") + m_user ;
	//}

	strTemp +=  TEXT(")(PORT = 1521))(CONNECT_DATA =(SERVER = DEDICATED)(SERVICE_NAME = ")+ m_databaseName;
	strTemp +=  TEXT(")));");
	m_bstrConnStr = strTemp.AllocSysString(); 



	//strTemp = TEXT("provider = OraOLEDB.Oracle;data source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = nano)(PORT = 1521))(CONNECT_DATA =(SERVER = DEDICATED)(SERVICE_NAME = RZFM)));");
	//m_bstrConnStr = strTemp.AllocSysString(); 

}
bool ADODATA::GetRecordset(char *sqlstr)
{
	CString m_sqlStr;
	BSTR    m_bstrSqlStr;

	m_sqlStr     = sqlstr;
	m_bstrSqlStr = m_sqlStr.AllocSysString(); 

	if(m_pRecordset->State == 1)
	{
		m_pRecordset->Close(); 
	}
	try
	{
		m_pRecordset->Open(m_bstrSqlStr, (IDispatch*)m_pConnection, adOpenDynamic, adLockOptimistic, adCmdText); 
	}
	catch(_com_error e)
	{
		// #ifdef _DEBUG
		//         CString errorMessage;
		//         errorMessage.Format("%s", (LPTSTR)e.Description());
		// //        AfxMessageBox(errorMessage);
		// #endif
		return false;
	}
	SysFreeString(m_bstrSqlStr);
	return true;
}
bool ADODATA::GetRecordset(CString sqlstr)
{
	CString m_sqlStr;
	BSTR    m_bstrSqlStr;

	m_sqlStr     = sqlstr;
	m_bstrSqlStr = m_sqlStr.AllocSysString(); 

	if(m_pRecordset->State == 1)
	{
		m_pRecordset->Close(); 
	}
	try
	{
		m_pRecordset->Open(m_bstrSqlStr, (IDispatch*)m_pConnection, adOpenDynamic, adLockOptimistic, adCmdText); 
	}
	catch(_com_error e)
	{
		// #ifdef _DEBUG
		//         CString errorMessage;
		//         errorMessage.Format("%s", (LPTSTR)e.Description());
		// //        AfxMessageBox(errorMessage);
		// #endif
		return false;
	}
	SysFreeString(m_bstrSqlStr);
	return true;
}

bool ADODATA::GetRecordset(BSTR sqlstr)
{
	CString m_sqlStr;
	BSTR    m_bstrSqlStr;

	m_sqlStr     = sqlstr;
	m_bstrSqlStr = m_sqlStr.AllocSysString(); 

	if(m_pRecordset->State == 1)
	{
		m_pRecordset->Close(); 
	}
	try
	{
		m_pRecordset->Open(m_bstrSqlStr, (IDispatch*)m_pConnection, adOpenDynamic, adLockOptimistic, adCmdText); 
	}
	catch(_com_error e)
	{
		// #ifdef _DEBUG
		//         CString errorMessage;
		//         errorMessage.Format("%s", (LPTSTR)e.Description());
		// //        AfxMessageBox(errorMessage);
		// #endif
		return false;
	}
	SysFreeString(m_bstrSqlStr);
	return true;
}

bool ADODATA::Execute(char *sqlstr)
{
	CString    m_sqlStr;
	BSTR       m_bstrSqlStr;
	_variant_t RecordsAffected;

	m_sqlStr     = sqlstr;
	m_bstrSqlStr = m_sqlStr.AllocSysString(); 

	if(m_pConnection->State != 1)
	{
		return false;
	}
	try
	{
		m_pConnection->BeginTrans(); 
		m_pConnection->Execute(m_bstrSqlStr, &RecordsAffected, adCmdText);
		m_pConnection->CommitTrans(); 
	}
	catch(_com_error e)
	{
		m_pConnection->RollbackTrans(); 
		// #ifdef _DEBUG
		//         CString errorMessage;
		//         errorMessage.Format("%s", (LPTSTR)e.Description());
		// //        AfxMessageBox(errorMessage);
		// #endif
		return false;
	}
	SysFreeString(m_bstrSqlStr);
	return true;
}

bool ADODATA::Execute(CString sqlstr)
{
	CString    m_sqlStr;
	BSTR       m_bstrSqlStr;
	_variant_t RecordsAffected;

	m_sqlStr     = sqlstr;
	m_bstrSqlStr = m_sqlStr.AllocSysString(); 

	if(m_pConnection->State != 1)
	{
		return false;
	}
	try
	{
		m_pConnection->BeginTrans(); 
		m_pConnection->Execute(m_bstrSqlStr, &RecordsAffected, adCmdText);
		m_pConnection->CommitTrans(); 
	}
	catch(_com_error e)
	{
		m_pConnection->RollbackTrans(); 
		// #ifdef _DEBUG
		//         CString errorMessage;
		//         errorMessage.Format("%s", (LPTSTR)e.Description());
		// //        AfxMessageBox(errorMessage);
		// #endif
		return false;
	}
	SysFreeString(m_bstrSqlStr);
	return true;
}

bool ADODATA::Execute(BSTR sqlstr)
{
	CString    m_sqlStr;
	BSTR       m_bstrSqlStr;
	_variant_t RecordsAffected;

	m_sqlStr     = sqlstr;
	m_bstrSqlStr = m_sqlStr.AllocSysString(); 

	if(m_pConnection->State != 1)
	{
		return false;
	}
	try
	{
		m_pConnection->BeginTrans(); 
		m_pConnection->Execute(m_bstrSqlStr, &RecordsAffected, adCmdText);
		m_pConnection->CommitTrans(); 
	}
	catch(_com_error e)
	{
		m_pConnection->RollbackTrans(); 
		// #ifdef _DEBUG
		//         CString errorMessage;
		//         errorMessage.Format("%s",(LPTSTR)e.Description());
		// //        AfxMessageBox(errorMessage);
		// #endif
		return false;
	}
	SysFreeString(m_bstrSqlStr);
	return true;
}

int ADODATA::GetRecordCount(void)
{
	int n=0;
	try
	{
		if(m_pRecordset->adoEOF)
		{
			return 0;
		}
		m_pRecordset->MoveFirst();
		while(!m_pRecordset->adoEOF)
		{
			n++;
			m_pRecordset->MoveNext();
		}
		m_pRecordset->MoveFirst();
		return n;
	}
	catch(_com_error e)
	{
		// #ifdef _DEBUG
		//         CString errorMessage;
		//         errorMessage.Format("%s", (LPTSTR)e.Description());
		// //        AfxMessageBox(errorMessage);
		// #endif
		return -1;
	}
}

bool ADODATA::ExecProc(void)
{
	_ParameterPtr m_pParameter1;
	_ParameterPtr m_pParameter2;
	_ParameterPtr m_pParameter3;
	m_pParameter1.CreateInstance("ADODB.Parameter");
	m_pParameter2.CreateInstance("ADODB.Parameter");
	m_pParameter3.CreateInstance("ADODB.Parameter");

	_CommandPtr m_pCommand1; 
	m_pCommand1.CreateInstance(_uuidof(Command));

	CString mystr = TEXT("addNum");

	if(m_pConnection->State != 1)
	{
		return false;
	}
	try
	{
		m_pParameter1->Name = "n1";
		m_pParameter1->Type = adSingle;
		m_pParameter1->Size = 4;            
		m_pParameter1->Direction = adParamInput;
		m_pParameter1->Value=_variant_t(10.0);
		m_pCommand1->Parameters->Append(m_pParameter1);

		m_pParameter2->Name = "n2";
		m_pParameter2->Type = adSingle;
		m_pParameter2->Size = 4;            
		m_pParameter2->Direction = adParamInput;
		m_pParameter2->Value = _variant_t(20.0);
		m_pCommand1->Parameters->Append(m_pParameter2);

		m_pParameter3->Name = "result";
		m_pParameter3->Type = adSingle;
		m_pParameter3->Size = 4;            
		m_pParameter3->Direction = adParamOutput;
		m_pCommand1->Parameters->Append(m_pParameter3);

		m_pCommand1->ActiveConnection = m_pConnection;
		m_pCommand1->CommandText = mystr.GetBuffer(0);
		m_pCommand1->CommandType = adCmdStoredProc;
		m_pCommand1->Execute(NULL, NULL, adCmdStoredProc);

		float result = (float)(m_pParameter3->Value);

		m_pParameter1.Release();
		m_pParameter2.Release();
		m_pParameter3.Release();
		m_pCommand1.Release();
	}
	catch(_com_error e)
	{
		// #ifdef _DEBUG
		//         CString errorMessage;
		//         errorMessage.Format("%s",(LPTSTR)e.Description());
		// //        AfxMessageBox(errorMessage);
		// #endif
		m_pParameter1.Release();
		m_pParameter2.Release();
		m_pParameter3.Release();
		m_pCommand1.Release();

		return false;
	}

	return true;
}