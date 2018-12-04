// C++ DataConnect.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"
#include "iostream"  
#include "string" 
#include "vector"  
//步骤1：添加对ADO的支持
#import "C:\Program Files\Common Files\System\ado\msado15.dll" no_namespace rename("EOF","adoEOF")  
using namespace std;

int _tmain(int argc, _TCHAR* argv[])
{
	CoInitialize(NULL); //初始化COM环境           
	_ConnectionPtr pMyConnect(__uuidof(Connection));//定义连接对象并实例化对象 
	_RecordsetPtr pRst(__uuidof(Recordset));//定义记录集对象并实例化对象               
	try           
	{              
		//步骤2：创建数据源连接
		/*打开数据库“SQLServer”，这里需要根据自己PC的数据库的情况 */             
		pMyConnect->Open("Provider=SQLOLEDB; Server=.;Database=db_test; uid=sa; pwd=123;","","",adModeUnknown);           
	} 
	catch (_com_error &e)           
	{               
		cout<<"Initiate failed!"<<endl;               
		cout<<e.Description()<<endl;               
		cout<<e.HelpFile()<<endl;               
		return 0;           
	}           
	cout<<"Connect succeed!"<<endl;                 

	//步骤3：对数据源中的数据库/表进行操作
	try           
	{

		//对SQL语句进行操作
		pRst = pMyConnect->Execute("select * from sheet1",NULL,adCmdText);//执行SQL语句

		//对SQL存储过程进行操作

		//这一句是执行存储过程的关键
		//GetSchemas2便是存储过程的名称
		//后面括号中的()便是其参数列表
		//注意Open方法的最后一个参数是adCmdStoredProc，而我们平常常用的是adCmdText
		//pRst->Open(L"GetSchemas2()",_variant_t((IDispatch*)pMyConnect),adOpenKeyset,adLockOptimistic,adCmdStoredProc);
		/*if(1 == pRst->State)
		{
			long n = pRst->GetRecordCount();
			wcout<<L"记录行数："<<n<<endl;
			if(n>0)
			{
				pRst->MoveFirst();
				while(!pRst->rsEOF)
				{
					int ID = pRst->GetCollect(L"ID");
					wstring sname = (_bstr_t)pRst->GetCollect(L"SchemaName");
					int DID = pRst->GetCollect(L"DesignerID");
					wstring name = (_bstr_t)pRst->GetCollect(L"DesignerTrueName");
					int PIR = pRst->GetCollect(L"Prior");
					wcout<<ID<<L"    "<<sname.c_str()<<L"   "<<DID<<L"   "<<name.c_str()<<L"   "<<PIR<<endl;
					pRst->MoveNext();
				}
			}
			pRst->Close();
		}*/


		if(!pRst->BOF) 
		{
			pRst->MoveFirst(); 
		}               
		else
		{                    
			cout<<"Data is empty!"<<endl;                     
			return 0;                
		}               
		vector<_bstr_t> column_name;      

		/*存储表的所有列名，显示表的列名*/               
		for(int i=0; i< pRst->Fields->GetCount();i++)               
		{                    
			cout<<pRst->Fields->GetItem(_variant_t((long)i))->Name<<" ";                    
			column_name.push_back(pRst->Fields->GetItem(_variant_t((long)i))->Name);               
		}   
		cout<<endl;

		/*对表进行遍历访问,显示表中每一行的内容*/               
		while(!pRst->adoEOF)               
		{                    
			vector<_bstr_t>::iterator iter=column_name.begin();                    
			for(iter;iter!=column_name.end();iter++)                    
			{                         
				if(pRst->GetCollect(*iter).vt !=VT_NULL)                         
				{  
					cout<<(_bstr_t)pRst->GetCollect(*iter)<<" ";                         
				}                         
				else
				{
					cout<<"NULL"<<endl;  
				}                  
			}
			pRst->MoveNext();                   
			cout<<endl;              
		}           
	}
	catch(_com_error &e)           
	{               
		cout<<e.Description()<<endl;               
		cout<<e.HelpFile()<<endl;               
		return 0;          
	}  

	//步骤4：关闭数据源
	/*关闭数据库并释放指针*/        
	try           
	{               
		pRst->Close();     //关闭记录集               
		pMyConnect->Close();//关闭数据库               
		pRst.Release();//释放记录集对象指针               
		pMyConnect.Release();//释放连接对象指针
	}
	catch(_com_error &e)           
	{               
		cout<<e.Description()<<endl;               
		cout<<e.HelpFile()<<endl;               
		return 0;           
	}                  
	CoUninitialize(); //释放COM环境
	return 0;
}

