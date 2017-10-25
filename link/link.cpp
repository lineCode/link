// link.cpp : 定义控制台应用程序的入口点。
// author : gouhongjie gohuge@qq.com
//
#include "stdafx.h"
#include <iostream>  
#include <ShlObj.h>  
#include <atlbase.h>  

using namespace std;

/********************************************************************
函数功能：创建桌面快捷方式
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
函数功能：创建快捷方式
参数：
szExeFileName   可执行文件完整路径
szWorkDir       当前工作目录(通常是szExeFileName所在目录)
szDescription   快捷方式的描述信息（备注）
szIconPath      设置快捷方式的图标
iIcon           指定图标在szIconPath中的序号，通常为0
szLnkFileName   存储快捷方式的完整路径
返回值：
TRUE    成功
FALSE   失败
********************************************************************/
BOOL CreateShortcut(PCTSTR szExeFileName, PCTSTR szWorkDir, PCTSTR szDescription,
	PCTSTR szIconPath, int iIcon, PCTSTR szLnkFileName)
{
	BOOL bOk;
	HRESULT hr;

	IShellLink *pLink;
	IPersistFile *pFile;

	//初始化COM组件  
	if (FAILED(CoInitialize(NULL)))
	{
		cerr << "COM初始化失败！" << endl;
		return FALSE;
	}

	//创建IShellLink的实例  
	hr = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (void **)&pLink);
	if (SUCCEEDED(hr))
	{
		//从IShellLink对象获取IPersistFile接口  
		hr = pLink->QueryInterface(IID_IPersistFile, (void **)&pFile);
		if (SUCCEEDED(hr))
		{
			//设置快捷方式的目标  
			pLink->SetPath(szExeFileName);
			//设置快捷方式的起始位置（当前工作目录）  
			//pLink->SetWorkingDirectory(szWorkDir);
			//设置快捷方式的描述信息（备注）  
			pLink->SetDescription(szDescription);
			//设置快捷方式的图标  
			pLink->SetIconLocation(szIconPath, iIcon);

			//保存快捷方式  
			hr = pFile->Save(szLnkFileName, TRUE);
			if (SUCCEEDED(hr))
				bOk = TRUE;

			//释放IPersistFile接口  
			pFile->Release();
		}
		else
		{
			cerr << "无法获取IPersistFile接口！" << endl;
			bOk = FALSE;
		}

		//释放IShellLink接口  
		pLink->Release();
	}
	else
	{
		cerr << "无法创建IShellLink实例！" << endl;
		bOk = FALSE;
	}

	//释放COM组件  
	CoUninitialize();

	return bOk;
}

/********************************************************************
函数功能：将一个可执行程序固定到Win7/Win8任务栏
参数：
szExeFileName   可执行文件完整路径
szWorkDir       当前工作目录(通常是szExeFileName所在目录)
szDescription   快捷方式的描述信息（备注）
szIconName      任务栏图标的名称
szIconPath      设置快捷方式的图标
iIcon           指定图标在szIconPath中的序号，通常为0
返回值：
TRUE    成功
FALSE   失败
备注：
1.  任务栏图标的名称指的是当鼠标停留在任务栏按钮上时显示的提示信息
例如用“记事本”比用“Notepad”更合理。
2.  ShellExecute函数不支持直接将可执行文件固定到任务栏，必须先创
建一个对应的快捷方式，然后再将快捷方式固定到任务栏，最后再删除
快捷方式。
********************************************************************/
BOOL PinToTaskBar(PCTSTR szExeFileName, PCTSTR szWorkDir, PCTSTR szDescription,
	PCTSTR szIconName, PCTSTR szIconPath, int iIcon)
{
	BOOL bOk = FALSE;

	//获取临时文件夹(假设为"C:\\Temp")  
	TCHAR szLnkFileName[MAX_PATH];
	GetTempPath(MAX_PATH, szLnkFileName);

	//拼接快捷方式路径"C:\\Temp\\"  
	if (_tcscat_s(szLnkFileName, _T("\\")) != 0)
	{
		cerr << "指定的可执行文件路径或图标名称太长！" << endl;
		return FALSE;
	}

	//拼接快捷方式路径"C:\\Temp\\记事本"  
	if (_tcscat_s(szLnkFileName, szIconName) != 0)
	{
		cerr << "指定的可执行文件路径或图标名称太长！" << endl;
		return FALSE;
	}

	//拼接快捷方式路径"C:\\Temp\\记事本.lnk"  
	if (_tcscat_s(szLnkFileName, _T(".lnk")) != 0)
	{
		cerr << "指定的可执行文件路径或图标名称太长！" << endl;
		return FALSE;
	}

	//第1步：在临时文件夹中创建快捷方式  
	bOk = CreateShortcut(szExeFileName, szWorkDir, szDescription, szIconPath, iIcon, szLnkFileName);
	if (bOk)
	{
		//第2步：固定到任务栏  
		int nRet = (int)::ShellExecute(NULL, _T("TaskbarPin"), szLnkFileName, NULL, NULL, SW_SHOW);
		//返回值大于32表示成功  
		if (nRet <= 32)
			cerr << "无法将快捷方式固定到任务栏！" << endl;
		//第3步：删除临时文件夹中创建的快捷方式  
		DeleteFile(szLnkFileName);
	}

	return bOk;
}


//从任务栏取消固定指定名称的程序  
BOOL UnPinFromTaskBar(PCTSTR szIconName)
{
	BOOL bOk;

	//获取用户应用程序配置数据的存储位置  
	TCHAR szPath[MAX_PATH];
	SHGetFolderPath(NULL, CSIDL_APPDATA, NULL, SHGFP_TYPE_CURRENT, szPath);

	//定位到任务栏图标所在位置  
	_tcscat_s(szPath, _T("\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar\\"));

	//定位到指定图标  
	if (_tcscat_s(szPath, szIconName) != 0)
	{
		cerr << "指定的图标名称太长！" << endl;
		return FALSE;
	}
	if (_tcscat_s(szPath, _T(".lnk")) != 0)
	{
		cerr << "指定的图标名称太长！" << endl;
		return FALSE;
	}

	//从任务栏取消固定某个快捷方式  
	int nRet = (int)::ShellExecute(NULL, _T("TaskbarUnPin"), szPath, NULL, NULL, SW_SHOW);
	if (nRet <= 32)
	{
		cerr << "无法删除指定的图标！" << endl;
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
	cout << "正在创建......\n" << endl;
	PinToTaskBar(_T("C:\\Windows\\Notepad.exe"), _T("C:\\Windows"), _T("使用记事本编写简单文本"), _T("记事本"), _T("C:\\Users\\mycomputer\\Pictures\\1.ico"), 0);
	exit(0);
}
void del()
{
	cout << "正在删除......\n" << endl;
	UnPinFromTaskBar(_T("记事本"));
	exit(0);
}


int _tmain()
{
	//cout << "演示程序:固定“记事本”到任务栏，并指定使用“计算器”程序的图标\n\n" << endl;  
	//system("pause");  


	setlocale(LC_ALL, "chs");
	
	// 先创建桌面快捷方式
	TCHAR szLinkFilePath[MAX_PATH] = L"C:\\Documents and Settings\\All Users\\桌面\\决战魔域H5.lnk";
	TCHAR szThisFilePath[MAX_PATH] = L"http://game.jzyx.com/jzmy";
	CoInitialize(NULL);
	CreateLink(szThisFilePath, szLinkFilePath);
	CoUninitialize();

	// 绑定程序快捷方式到任务栏
	PinToTaskBar(_T("C:\\Documents and Settings\\All Users\\桌面\\决战魔域H5.lnk"), _T("C:\\Documents and Settings\\All Users\\")
		, _T("极致游戏，决战魔域H5"), _T("决战魔域H5"), szLinkFilePath, 0);

	int i;
	while (1) {
		printf("选择菜单：\n1 在任务栏创建后台快捷方式\n2 删除在任务栏创建的后台快捷方式\n3 退出\n");
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

/*  附录

IShellLink主要成员：
1、GetArguments：获得参数信息
2、GetDescription：获得描述信息（备注行）
3、GetHotkey：获得快捷键
4、GetIconLocation：获得图标
5、GetIDList：获得快捷方式的目标对象的item identifier list
(Windows外壳中的每个对象如文件，目录和打印机等都有唯一的item identifiler list)
6、GetPath: 获得快捷方式的目标文件或目录的全路径
7、GetShowCmd：获得快捷方式的运行方式，比如常规窗口，最大化
8、GetWorkingDirectory：获得工作目录
9、Resolve：按照一定的搜索规则试图获得目标对象，即使目标对象已经被删除或移动，重命名
下面是对应信息的设置方法
10、SetArguments
11、SetDescription
12、SetHotkey
13、SetIconLocation
14、SetIDList
15、SetPath
16、SetRelativePath
17、SetShowCmd
18、SetWorkingDirectory

*/
