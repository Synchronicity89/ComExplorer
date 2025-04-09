// ComExplorer.cpp : Implementation of WinMain


#include "pch.h"
#include "framework.h"
#include "resource.h"
#include "ComExplorer_i.h"


using namespace ATL;


class CComExplorerModule : public ATL::CAtlExeModuleT< CComExplorerModule >
{
public :
	DECLARE_LIBID(LIBID_ComExplorerLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_COMEXPLORER, "{00c8a28a-f581-4b00-8663-bdb995221441}")
};

CComExplorerModule _AtlModule;



//
extern "C" int WINAPI _tWinMain(HINSTANCE /*hInstance*/, HINSTANCE /*hPrevInstance*/,
								LPTSTR /*lpCmdLine*/, int nShowCmd)
{
	return _AtlModule.WinMain(nShowCmd);
}

