#include <windows.h>

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	  if(fdwReason==DLL_PROCESS_ATTACH){
		  MessageBox(0,"InDLlMain!","",0);
	  }
	  return 1;
}
