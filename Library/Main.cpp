#include "main.h"

#pragma data_seg(".shared")
	HHOOK messageHook = 0 ;
	int linesNumber = 3 ;
	int controlKey = 1 ;
	int pressKey = 0 ;
	int holdKey = 0 ;
#pragma data_seg()

HINSTANCE dllInstance ;

HWND GetScrollBar(HWND parent, long position)
{
	HWND handle = FindWindowEx(parent, NULL, "ScrollBar", NULL) ;
	long style ;
	POINT p ;
	RECT r ;

	while (handle != NULL)
	{
		style = GetWindowLong(handle, GWL_STYLE) ;

		if (((style & position) == position) && ((style & WS_VISIBLE) == WS_VISIBLE))
		{
			if (position == SBS_VERT)
			{
				GetCursorPos(&p) ;
				GetWindowRect(handle, &r) ;

				if (p.y >= r.top && p.y <= r.bottom)
				{
					return handle ;
				}
			}
			else
			{
				return handle ;
			}
		}

		handle = FindWindowEx(parent, handle, "ScrollBar", NULL) ;
	}

	return NULL ;
}

BOOL IsMouseCloseToVScrollBar(POINT p, RECT r)	//by ding
{
	if(p.x > r.left - 25 && p.y > r.top + 25 && p.y < r.bottom - 25)
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

LRESULT CALLBACK GetMsgProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	if (nCode != HC_ACTION || wParam != PM_NOREMOVE)
	{
		//return CallNextHookEx(messageHook, nCode, wParam, lParam) ;	//by ding
		return 0;
	}

	long getMessage = ((MSG*)lParam)->message ;

	if (getMessage == WM_MOUSEWHEEL)
	{
		HWND thisWindow = ((MSG*)lParam)->hwnd ;
		char className[255] ;

		GetClassName(thisWindow, className, 255) ;

		if (strcmp(className, "VbaWindow") == 0)
		{
			HWND scrollBar = NULL ;
			long windowMessage = WM_VSCROLL ;
			RECT r ;	//by ding
			GetWindowRect(GetScrollBar(thisWindow, SBS_VERT), &r) ;
			if(IsMouseCloseToVScrollBar(((MSG*)lParam)->pt, r))
			{
				windowMessage = WM_HSCROLL ;
			}
			short scrollDelta = (short)HIWORD(((MSG*)lParam)->wParam) ;

			if (GetAsyncKeyState(VK_SHIFT) & 0x80000000)	//if (GetAsyncKeyState(VK_CONTROL) & 0x80000000)	//by ding
			{
				switch (controlKey)
				{
				case 0:
					scrollBar = GetScrollBar(thisWindow, SBS_VERT) ;
				break ;
				case 1:
					//scrollBar = GetScrollBar(thisWindow, SBS_HORZ) ;	//by ding
					//windowMessage = WM_HSCROLL ;
					if(windowMessage == WM_HSCROLL)
					{
						scrollBar = GetScrollBar(thisWindow, SBS_VERT) ;
						windowMessage = WM_VSCROLL;
					}
					else
					{
						scrollBar = GetScrollBar(thisWindow, SBS_HORZ) ;
						windowMessage = WM_HSCROLL;
					}
				break ;
				case 2:
					BYTE keyCode ;
					
					keyCode = (scrollDelta < 0 ? VK_DOWN : VK_UP ) ;

					keybd_event(keyCode, 0, 0, 0) ;
					keybd_event(keyCode, 0, KEYEVENTF_KEYUP, 0) ;
				break ;
				case 3:
					if (scrollDelta > 0)
					{
						keybd_event(VK_SHIFT, 0, 0, 0) ;
					}
					keybd_event(VK_TAB, 0, 0, 0) ;
					keybd_event(VK_TAB, 0, KEYEVENTF_KEYUP, 0) ;
					if (scrollDelta > 0)
					{
						keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0) ;
					}
				break ;
				}
			}
			else
			{
				scrollBar = GetScrollBar(thisWindow, SBS_VERT) ;
			}

			if (scrollBar)
			{
				long scrollMessage ;
				int scrollLoops ;

				scrollLoops = (windowMessage == WM_HSCROLL ? linesNumber : 1) ;

				if (scrollDelta < 0)
				{
					if (windowMessage == WM_HSCROLL)
					{
						scrollMessage = (linesNumber ? SB_LINERIGHT : SB_PAGERIGHT) ;
					}
					else
					{
						scrollMessage = (linesNumber ? SB_LINEDOWN : SB_PAGEDOWN) ;
					}
				}
				else
				{
					if (windowMessage == WM_HSCROLL)
					{
						scrollMessage = (linesNumber ? SB_LINELEFT : SB_PAGELEFT) ;
					}
					else
					{
						scrollMessage = (linesNumber ? SB_LINEUP : SB_PAGEUP) ;
					}
				}

				do
				{
					SendMessage(thisWindow, windowMessage, scrollMessage, (LPARAM)scrollBar) ;
					SendMessage(thisWindow, windowMessage, SB_ENDSCROLL, (LPARAM)scrollBar) ;
				}
				while (++scrollLoops <= linesNumber) ;
			}
		}
	}
	else if (getMessage == WM_MBUTTONUP)
	{
		if (pressKey)
		{
			if (holdKey)
			{
				keybd_event(holdKey, 0, 0, 0) ;
			}
			keybd_event(pressKey, 0, 0, 0) ;
			keybd_event(pressKey, 0, KEYEVENTF_KEYUP, 0) ;
			if (holdKey)
			{
				keybd_event(holdKey, 0, KEYEVENTF_KEYUP, 0) ;
			}
		}
	}

	//return CallNextHookEx(messageHook, nCode, wParam, lParam) ;	//by ding
	return 0;
}

HHOOK WINAPI EnableScroll()
{
	if (messageHook == 0)
	{
		messageHook = SetWindowsHookEx(WH_GETMESSAGE, (HOOKPROC)GetMsgProc, dllInstance, 0) ;
	}

	return messageHook ;
}

HHOOK WINAPI DisableScroll()
{
	if (UnhookWindowsHookEx(messageHook))
	{
		messageHook = 0 ;
	}

	return messageHook ;
}

void WINAPI ScrollLines(int number)
{
	linesNumber = number ;
}

void WINAPI SetCtrlKey(int value)
{
	controlKey = value ;
}

void WINAPI SetWheelButton(int press, int hold)
{
	pressKey = press ;
	holdKey = hold ;
}

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	if (fdwReason == DLL_PROCESS_ATTACH)
	{
		dllInstance = (HINSTANCE)hinstDLL ;
		DisableThreadLibraryCalls(hinstDLL) ;
	}

	return TRUE ;
}

extern "C" BOOL __stdcall _DllMainCRTStartup(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
    return DllMain(hinstDLL, fdwReason, lpvReserved) ;
}
