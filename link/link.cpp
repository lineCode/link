// link.cpp : �������̨Ӧ�ó������ڵ㡣
// author : gouhongjie gohuge@qq.com
//
#include "stdafx.h"
#include <iostream>  
#include <ShlObj.h>  
#include <atlbase.h>  

using namespace std;

/********************************************************************
�������ܣ����������ݷ�ʽ
********************************************************************/
static HRESULT CreateLink(TCHAR *lpszPathObj, TCHAR *lpszPathLink)
{
	HRESULT hres;
	IShellLink* psl;

	hres = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID*)&psl);
	if (SUCCEEDED(hres))
	{
		IPersistFile* ppf;
		psl->SetPath(lpszPathObj);
		hres = psl->QueryInterface(IID_IPersistFile, (LPVOID*)&ppf);

		if (SUCCEEDED(hres))
		{
			hres = ppf->Save(lpszPathLink, TRUE);
			ppf->Release();
		}
		psl->Release();
	}
	return hres;
}

/********************************************************************
�������ܣ�������ݷ�ʽ
������
szExeFileName   ��ִ���ļ�����·��
szWorkDir       ��ǰ����Ŀ¼(ͨ����szExeFileName����Ŀ¼)
szDescription   ��ݷ�ʽ��������Ϣ����ע��
szIconPath      ���ÿ�ݷ�ʽ��ͼ��
iIcon           ָ��ͼ����szIconPath�е���ţ�ͨ��Ϊ0
szLnkFileName   �洢��ݷ�ʽ������·��
����ֵ��
TRUE    �ɹ�
FALSE   ʧ��
********************************************************************/
BOOL CreateShortcut(PCTSTR szExeFileName, PCTSTR szWorkDir, PCTSTR szDescription,
	PCTSTR szIconPath, int iIcon, PCTSTR szLnkFileName)
{
	BOOL bOk;
	HRESULT hr;

	IShellLink *pLink;
	IPersistFile *pFile;

	//��ʼ��COM���  
	if (FAILED(CoInitialize(NULL)))
	{
		cerr << "COM��ʼ��ʧ�ܣ�" << endl;
		return FALSE;
	}

	//����IShellLink��ʵ��  
	hr = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (void **)&pLink);
	if (SUCCEEDED(hr))
	{
		//��IShellLink�����ȡIPersistFile�ӿ�  
		hr = pLink->QueryInterface(IID_IPersistFile, (void **)&pFile);
		if (SUCCEEDED(hr))
		{
			//���ÿ�ݷ�ʽ��Ŀ��  
			pLink->SetPath(szExeFileName);
			//���ÿ�ݷ�ʽ����ʼλ�ã���ǰ����Ŀ¼��  
			//pLink->SetWorkingDirectory(szWorkDir);
			//���ÿ�ݷ�ʽ��������Ϣ����ע��  
			pLink->SetDescription(szDescription);
			//���ÿ�ݷ�ʽ��ͼ��  
			pLink->SetIconLocation(szIconPath, iIcon);

			//�����ݷ�ʽ  
			hr = pFile->Save(szLnkFileName, TRUE);
			if (SUCCEEDED(hr))
				bOk = TRUE;

			//�ͷ�IPersistFile�ӿ�  
			pFile->Release();
		}
		else
		{
			cerr << "�޷���ȡIPersistFile�ӿڣ�" << endl;
			bOk = FALSE;
		}

		//�ͷ�IShellLink�ӿ�  
		pLink->Release();
	}
	else
	{
		cerr << "�޷�����IShellLinkʵ����" << endl;
		bOk = FALSE;
	}

	//�ͷ�COM���  
	CoUninitialize();

	return bOk;
}

/********************************************************************
�������ܣ���һ����ִ�г���̶���Win7/Win8������
������
szExeFileName   ��ִ���ļ�����·��
szWorkDir       ��ǰ����Ŀ¼(ͨ����szExeFileName����Ŀ¼)
szDescription   ��ݷ�ʽ��������Ϣ����ע��
szIconName      ������ͼ�������
szIconPath      ���ÿ�ݷ�ʽ��ͼ��
iIcon           ָ��ͼ����szIconPath�е���ţ�ͨ��Ϊ0
����ֵ��
TRUE    �ɹ�
FALSE   ʧ��
��ע��
1.  ������ͼ�������ָ���ǵ����ͣ������������ť��ʱ��ʾ����ʾ��Ϣ
�����á����±������á�Notepad��������
2.  ShellExecute������֧��ֱ�ӽ���ִ���ļ��̶����������������ȴ�
��һ����Ӧ�Ŀ�ݷ�ʽ��Ȼ���ٽ���ݷ�ʽ�̶����������������ɾ��
��ݷ�ʽ��
********************************************************************/
BOOL PinToTaskBar(PCTSTR szExeFileName, PCTSTR szWorkDir, PCTSTR szDescription,
	PCTSTR szIconName, PCTSTR szIconPath, int iIcon)
{
	BOOL bOk = FALSE;

	//��ȡ��ʱ�ļ���(����Ϊ"C:\\Temp")  
	TCHAR szLnkFileName[MAX_PATH];
	GetTempPath(MAX_PATH, szLnkFileName);

	//ƴ�ӿ�ݷ�ʽ·��"C:\\Temp\\"  
	if (_tcscat_s(szLnkFileName, _T("\\")) != 0)
	{
		cerr << "ָ���Ŀ�ִ���ļ�·����ͼ������̫����" << endl;
		return FALSE;
	}

	//ƴ�ӿ�ݷ�ʽ·��"C:\\Temp\\���±�"  
	if (_tcscat_s(szLnkFileName, szIconName) != 0)
	{
		cerr << "ָ���Ŀ�ִ���ļ�·����ͼ������̫����" << endl;
		return FALSE;
	}

	//ƴ�ӿ�ݷ�ʽ·��"C:\\Temp\\���±�.lnk"  
	if (_tcscat_s(szLnkFileName, _T(".lnk")) != 0)
	{
		cerr << "ָ���Ŀ�ִ���ļ�·����ͼ������̫����" << endl;
		return FALSE;
	}

	//��1��������ʱ�ļ����д�����ݷ�ʽ  
	bOk = CreateShortcut(szExeFileName, szWorkDir, szDescription, szIconPath, iIcon, szLnkFileName);
	if (bOk)
	{
		//��2�����̶���������  
		int nRet = (int)::ShellExecute(NULL, _T("TaskbarPin"), szLnkFileName, NULL, NULL, SW_SHOW);
		//����ֵ����32��ʾ�ɹ�  
		if (nRet <= 32)
			cerr << "�޷�����ݷ�ʽ�̶�����������" << endl;
		//��3����ɾ����ʱ�ļ����д����Ŀ�ݷ�ʽ  
		DeleteFile(szLnkFileName);
	}

	return bOk;
}


//��������ȡ���̶�ָ�����Ƶĳ���  
BOOL UnPinFromTaskBar(PCTSTR szIconName)
{
	BOOL bOk;

	//��ȡ�û�Ӧ�ó����������ݵĴ洢λ��  
	TCHAR szPath[MAX_PATH];
	SHGetFolderPath(NULL, CSIDL_APPDATA, NULL, SHGFP_TYPE_CURRENT, szPath);

	//��λ��������ͼ������λ��  
	_tcscat_s(szPath, _T("\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar\\"));

	//��λ��ָ��ͼ��  
	if (_tcscat_s(szPath, szIconName) != 0)
	{
		cerr << "ָ����ͼ������̫����" << endl;
		return FALSE;
	}
	if (_tcscat_s(szPath, _T(".lnk")) != 0)
	{
		cerr << "ָ����ͼ������̫����" << endl;
		return FALSE;
	}

	//��������ȡ���̶�ĳ����ݷ�ʽ  
	int nRet = (int)::ShellExecute(NULL, _T("TaskbarUnPin"), szPath, NULL, NULL, SW_SHOW);
	if (nRet <= 32)
	{
		cerr << "�޷�ɾ��ָ����ͼ�꣡" << endl;
		bOk = FALSE;
	}
	else
	{
		bOk = TRUE;
	}

	return bOk;

}


void cre()
{
	cout << "���ڴ���......\n" << endl;
	PinToTaskBar(_T("C:\\Windows\\Notepad.exe"), _T("C:\\Windows"), _T("ʹ�ü��±���д���ı�"), _T("���±�"), _T("C:\\Users\\mycomputer\\Pictures\\1.ico"), 0);
	exit(0);
}
void del()
{
	cout << "����ɾ��......\n" << endl;
	UnPinFromTaskBar(_T("���±�"));
	exit(0);
}


int _tmain()
{
	//cout << "��ʾ����:�̶������±���������������ָ��ʹ�á��������������ͼ��\n\n" << endl;  
	//system("pause");  


	setlocale(LC_ALL, "chs");
	
	// �ȴ��������ݷ�ʽ
	TCHAR szLinkFilePath[MAX_PATH] = L"C:\\Documents and Settings\\All Users\\����\\��սħ��H5.lnk";
	TCHAR szThisFilePath[MAX_PATH] = L"http://game.jzyx.com/jzmy";
	CoInitialize(NULL);
	CreateLink(szThisFilePath, szLinkFilePath);
	CoUninitialize();

	// �󶨳����ݷ�ʽ��������
	PinToTaskBar(_T("C:\\Documents and Settings\\All Users\\����\\��սħ��H5.lnk"), _T("C:\\Documents and Settings\\All Users\\")
		, _T("������Ϸ����սħ��H5"), _T("��սħ��H5"), szLinkFilePath, 0);

	int i;
	while (1) {
		printf("ѡ��˵���\n1 ��������������̨��ݷ�ʽ\n2 ɾ���������������ĺ�̨��ݷ�ʽ\n3 �˳�\n");
		scanf_s("%d", &i);
		if (i == 1)
		{
			cre();
		}
		if (i == 2)
		{
			del();
		}
		if (i == 3)
		{
			exit(0);
		}
	}

}

/*  ��¼

IShellLink��Ҫ��Ա��
1��GetArguments����ò�����Ϣ
2��GetDescription�����������Ϣ����ע�У�
3��GetHotkey����ÿ�ݼ�
4��GetIconLocation�����ͼ��
5��GetIDList����ÿ�ݷ�ʽ��Ŀ������item identifier list
(Windows����е�ÿ���������ļ���Ŀ¼�ʹ�ӡ���ȶ���Ψһ��item identifiler list)
6��GetPath: ��ÿ�ݷ�ʽ��Ŀ���ļ���Ŀ¼��ȫ·��
7��GetShowCmd����ÿ�ݷ�ʽ�����з�ʽ�����糣�洰�ڣ����
8��GetWorkingDirectory����ù���Ŀ¼
9��Resolve������һ��������������ͼ���Ŀ����󣬼�ʹĿ������Ѿ���ɾ�����ƶ���������
�����Ƕ�Ӧ��Ϣ�����÷���
10��SetArguments
11��SetDescription
12��SetHotkey
13��SetIconLocation
14��SetIDList
15��SetPath
16��SetRelativePath
17��SetShowCmd
18��SetWorkingDirectory

*/
