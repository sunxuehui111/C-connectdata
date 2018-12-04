// C++ DataConnect.cpp : �������̨Ӧ�ó������ڵ㡣
//

#include "stdafx.h"
#include "iostream"  
#include "string" 
#include "vector"  
//����1����Ӷ�ADO��֧��
#import "C:\Program Files\Common Files\System\ado\msado15.dll" no_namespace rename("EOF","adoEOF")  
using namespace std;

int _tmain(int argc, _TCHAR* argv[])
{
	CoInitialize(NULL); //��ʼ��COM����           
	_ConnectionPtr pMyConnect(__uuidof(Connection));//�������Ӷ���ʵ�������� 
	_RecordsetPtr pRst(__uuidof(Recordset));//�����¼������ʵ��������               
	try           
	{              
		//����2����������Դ����
		/*�����ݿ⡰SQLServer����������Ҫ�����Լ�PC�����ݿ����� */             
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

	//����3��������Դ�е����ݿ�/����в���
	try           
	{

		//��SQL�����в���
		pRst = pMyConnect->Execute("select * from sheet1",NULL,adCmdText);//ִ��SQL���

		//��SQL�洢���̽��в���

		//��һ����ִ�д洢���̵Ĺؼ�
		//GetSchemas2���Ǵ洢���̵�����
		//���������е�()����������б�
		//ע��Open���������һ��������adCmdStoredProc��������ƽ�����õ���adCmdText
		//pRst->Open(L"GetSchemas2()",_variant_t((IDispatch*)pMyConnect),adOpenKeyset,adLockOptimistic,adCmdStoredProc);
		/*if(1 == pRst->State)
		{
			long n = pRst->GetRecordCount();
			wcout<<L"��¼������"<<n<<endl;
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

		/*�洢���������������ʾ�������*/               
		for(int i=0; i< pRst->Fields->GetCount();i++)               
		{                    
			cout<<pRst->Fields->GetItem(_variant_t((long)i))->Name<<" ";                    
			column_name.push_back(pRst->Fields->GetItem(_variant_t((long)i))->Name);               
		}   
		cout<<endl;

		/*�Ա���б�������,��ʾ����ÿһ�е�����*/               
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

	//����4���ر�����Դ
	/*�ر����ݿⲢ�ͷ�ָ��*/        
	try           
	{               
		pRst->Close();     //�رռ�¼��               
		pMyConnect->Close();//�ر����ݿ�               
		pRst.Release();//�ͷż�¼������ָ��               
		pMyConnect.Release();//�ͷ����Ӷ���ָ��
	}
	catch(_com_error &e)           
	{               
		cout<<e.Description()<<endl;               
		cout<<e.HelpFile()<<endl;               
		return 0;           
	}                  
	CoUninitialize(); //�ͷ�COM����
	return 0;
}

