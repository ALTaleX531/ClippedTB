// ClippedTB.cpp : 定义应用程序的入口点。
//

#include "pch.h"
#include "ClippedTB.h"

namespace ClippedTB
{
	using namespace std;

	struct ServiceInfo
	{
		HWND primaryRedirectorHost{ nullptr };
		HWND ui{ nullptr };
	};
	using unique_service_info = unique_ptr < ServiceInfo, decltype([](ServiceInfo* ptr)
	{
		if (ptr)
		{
			UnmapViewOfFile(ptr);
		}
	}) > ;

	const UINT WM_CTBSTOP{ RegisterWindowMessageW(L"ClippedTB.Stop") };
	const UINT WM_CTBSHOWUI{ RegisterWindowMessageW(L"ClippedTB.ShowUI") };
	const UINT WM_TASKBARCREATED{ RegisterWindowMessageW(L"TaskbarCreated")};
	constexpr UINT WM_TASKBARICON{ WM_APP + 1 };

	constexpr auto fileMappingName{ L"Local\\ClippedTB"sv };
	constexpr auto taskName{L"ClippedTB Autorun Task"sv};
	constexpr auto registryItemName{L"SOFTWARE\\ClippedTB"sv};

	constexpr auto marginsLeftRegistryItemName{L"MarginsLeft"sv};
	constexpr auto marginsTopRegistryItemName{ L"MarginsTop"sv };
	constexpr auto marginsRightRegistryItemName{ L"MarginsRight"sv };
	constexpr auto marginsBottomRegistryItemName{ L"MarginsBottom"sv };

	constexpr auto topLeftCornerRadiusXRegistryItemName{ L"TopLeftCornerRadiusX"sv };
	constexpr auto topLeftCornerRadiusYRegistryItemName{ L"topLeftCornerRadiusY"sv };
	constexpr auto bottomLeftCornerRadiusXRegistryItemName{ L"BottomLeftCornerRadiusX"sv };
	constexpr auto bottomLeftCornerRadiusYRegistryItemName{ L"BottomLeftCornerRadiusY"sv };
	constexpr auto topRightCornerRadiusXRegistryItemName{ L"TopRightCornerRadiusX"sv };
	constexpr auto topRightCornerRadiusYRegistryItemName{ L"TopRightCornerRadiusY"sv };
	constexpr auto bottomRightCornerRadiusXRegistryItemName{ L"BottomRightCornerRadiusX"sv };
	constexpr auto bottomRightCornerRadiusYRegistryItemName{ L"BottomRightCornerRadiusY"sv };

	struct Config
	{
		float marginsLeft{ 12.5f };
		float marginsTop{ 2.5f };
		float marginsRight{ 12.5f };
		float marginsBottom{ 2.5f };

		float topLeftCornerRadiusX{ 12.f };
		float topLeftCornerRadiusY{ 12.f };
		float bottomLeftCornerRadiusX{ 12.f };
		float bottomLeftCornerRadiusY{ 12.f };
		float topRightCornerRadiusX{ 12.f };
		float topRightCornerRadiusY{ 12.f };
		float bottomRightCornerRadiusX{ 12.f };
		float bottomRightCornerRadiusY{ 12.f };

		HRESULT WriteToRegistry()
		{
#pragma push_macro("WriteRegAndCheckHR")
#undef WriteRegAndCheckHR
#define WriteRegAndCheckHR(item) \
	RETURN_IF_FAILED(\
					 wil::reg::set_value_dword_nothrow(\
							 HKEY_CURRENT_USER,\
							 registryItemName.data(),\
							 item##RegistryItemName .data(),\
							 lroundf(item * 10.f)\
													  ) \
					)

			WriteRegAndCheckHR(marginsLeft);
			WriteRegAndCheckHR(marginsTop);
			WriteRegAndCheckHR(marginsRight);
			WriteRegAndCheckHR(marginsBottom);

			WriteRegAndCheckHR(topLeftCornerRadiusX);
			WriteRegAndCheckHR(topLeftCornerRadiusY);
			WriteRegAndCheckHR(bottomLeftCornerRadiusX);
			WriteRegAndCheckHR(bottomLeftCornerRadiusY);
			WriteRegAndCheckHR(topRightCornerRadiusX);
			WriteRegAndCheckHR(topRightCornerRadiusY);
			WriteRegAndCheckHR(bottomRightCornerRadiusX);
			WriteRegAndCheckHR(bottomRightCornerRadiusY);

#pragma pop_macro("WriteRegAndCheckHR")

			return S_OK;
		}
		HRESULT ReadFromRegistry()
		{
			*this = Config{};
#pragma push_macro("ReadRegAndCheckHR")
#undef ReadRegAndCheckHR
#define ReadRegAndCheckHR(item) \
	RETURN_IF_FAILED( \
					  [&]() \
	{ \
		DWORD value{0}; \
		HRESULT hr \
		{ \
			wil::reg::get_value_dword_nothrow( \
											   HKEY_CURRENT_USER, \
											   registryItemName.data(), \
											   item##RegistryItemName.data(),\
											   &value \
											 ) \
		}; \
		if (SUCCEEDED(hr)) \
		{ \
			item = static_cast<float>(value) / 10.f; \
			return hr; \
		} \
		return (hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND)) ? S_OK : hr; \
	} () \
					)

			ReadRegAndCheckHR(marginsLeft);
			ReadRegAndCheckHR(marginsTop);
			ReadRegAndCheckHR(marginsRight);
			ReadRegAndCheckHR(marginsBottom);

			ReadRegAndCheckHR(topLeftCornerRadiusX);
			ReadRegAndCheckHR(topLeftCornerRadiusY);
			ReadRegAndCheckHR(bottomLeftCornerRadiusX);
			ReadRegAndCheckHR(bottomLeftCornerRadiusY);
			ReadRegAndCheckHR(topRightCornerRadiusX);
			ReadRegAndCheckHR(topRightCornerRadiusY);
			ReadRegAndCheckHR(bottomRightCornerRadiusX);
			ReadRegAndCheckHR(bottomRightCornerRadiusY);

#pragma pop_macro("ReadRegAndCheckHR")

			return S_OK;
		}
	};

	class TBContentRedirector
	{
	public:
		TBContentRedirector() = default;
		~TBContentRedirector()
		{
			Attach(nullptr, nullptr);
		};

		HWND GetSourceTaskbar() const
		{
			return m_srcWnd;
		}
		HWND GetRedirectorHost() const
		{
			return m_targetWnd;
		}
		HRESULT Attach(HWND srcWnd, HWND targetWnd);
		HRESULT Update();
	private:
		HRESULT EnsureDevice(bool& needToRecreate);
		HRESULT UpdateClipRegion();

		Config                                      m_lastConfig{};
		HWND                                        m_srcWnd{ nullptr };
		HWND                                        m_targetWnd{ nullptr };

		wil::com_ptr<IDXGIDevice3>                  m_dxgiDevice{ nullptr };
		wil::com_ptr<ID3D11Device>                  m_d3dDevice{ nullptr };
		wil::com_ptr<IDCompositionDesktopDevice>    m_dcompDevice{ nullptr };

		wil::com_ptr<IDCompositionVisual2>			m_rootVisual{nullptr};
		wil::com_ptr<IDCompositionTarget>	        m_dcompTarget{ nullptr };

		HTHUMBNAIL						            m_thumbnail{ nullptr };
		wil::com_ptr<IDCompositionVisual2>	        m_thumbnailVisual{ nullptr };
	};

	inline HRESULT ReturnAndCheckHR(HRESULT hr)
	{
		if (FAILED(hr))
		{
			WCHAR szError[MAX_PATH + 1] {};
			FormatMessageW(
				FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
				nullptr,
				hr,
				0,
				szError,
				_countof(szError),
				nullptr
			);
			ShellMessageBoxW(
				HINST_THISCOMPONENT,
				nullptr,
				szError,
				nullptr,
				MB_ICONERROR | MB_SYSTEMMODAL
			);
		}

		return hr;
	}

	//
	HRESULT ShowUI();
	HRESULT StopService();
	HRESULT StartService();
	//
	HRESULT Install();
	HRESULT Uninstall();
	bool IsCTBInstalled();

	inline bool IsServiceRunning();
	inline unique_service_info GetServiceInfo();

	namespace UI
	{
		HRESULT TaskbarCreateIcon(HWND hWnd)
		{
			HICON icon{nullptr};
			RETURN_IF_FAILED(LoadIconMetric(HINST_THISCOMPONENT, MAKEINTRESOURCE(IDI_ICON1), LIM_SMALL, &icon));
			NOTIFYICONDATAW data{ sizeof(data), hWnd, IDI_ICON1, NIF_ICON | NIF_MESSAGE | NIF_TIP, WM_TASKBARICON, icon, {L"ClippedTB"}};
			RETURN_IF_WIN32_BOOL_FALSE(Shell_NotifyIconW(NIM_ADD, &data));

			return S_OK;
		}
		HRESULT TaskbarDestroyIcon(HWND hWnd)
		{
			NOTIFYICONDATAW data{ sizeof(data), hWnd, IDI_ICON1 };
			RETURN_IF_WIN32_BOOL_FALSE(Shell_NotifyIconW(NIM_DELETE, &data));

			return S_OK;
		}
		HRESULT Initialize();
		HRESULT ReadRegistry();
		INT_PTR CALLBACK DialogProc(HWND, UINT, WPARAM, LPARAM);
	}
}

int APIENTRY wWinMain(
	_In_ HINSTANCE hInstance,
	_In_opt_ HINSTANCE hPrevInstance,
	_In_ LPWSTR lpCmdLine,
	_In_ int nShowCmd
)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(nShowCmd);

	// Until now, We only support Chinese and English...
	if (GetUserDefaultUILanguage() != MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED))
	{
		SetThreadUILanguage(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US));
	}
	RETURN_IF_WIN32_BOOL_FALSE(SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2));

	SetUnhandledExceptionFilter([](EXCEPTION_POINTERS * ExceptionInfo) -> LONG
	{
		WCHAR msg[MAX_PATH];
		THROW_LAST_ERROR_IF(!LoadStringW(HINST_THISCOMPONENT, IDS_STRING132, msg, MAX_PATH));

		ShellMessageBoxW(
			HINST_THISCOMPONENT,
			nullptr,
			std::format(
				L"{}\nExceptionCode: {:#x}\nExceptionAddress: {}\nNumberParameters: {}",
				msg,
				ExceptionInfo->ExceptionRecord->ExceptionCode,
				ExceptionInfo->ExceptionRecord->ExceptionAddress,
				ExceptionInfo->ExceptionRecord->NumberParameters
			).c_str(),
			nullptr,
			MB_ICONERROR | MB_SYSTEMMODAL
		);

		// Make all the taskbars visible
		EnumWindows([](HWND hWnd, LPARAM lParam) -> BOOL
		{
			WCHAR className[MAX_PATH + 1]{};
			GetClassNameW(hWnd, className, MAX_PATH);
			if (!_wcsicmp(className, L"Shell_SecondaryTrayWnd") || !_wcsicmp(className, L"Shell_TrayWnd"))
			{
				SetLayeredWindowAttributes(hWnd, 0, 255, LWA_ALPHA);
				SetWindowLongPtrW(hWnd, GWL_EXSTYLE, GetWindowLongPtrW(hWnd, GWL_EXSTYLE) & ~WS_EX_LAYERED);
			}

			return TRUE;
		}, 0);

		return EXCEPTION_EXECUTE_HANDLER;
	});

	if (!_wcsicmp(lpCmdLine, L"/stop") || !_wcsicmp(lpCmdLine, L"-stop"))
	{
		return ClippedTB::ReturnAndCheckHR(ClippedTB::StopService());
	}

	if (ClippedTB::IsServiceRunning())
	{
		return ClippedTB::ReturnAndCheckHR(ClippedTB::ShowUI());
	}

	return ClippedTB::ReturnAndCheckHR(ClippedTB::StartService());
}

HRESULT ClippedTB::StartService()
{
	wil::unique_handle fileMapping{ nullptr };
	unique_service_info serviceInfo{ nullptr };

	RETURN_HR_IF(HRESULT_FROM_WIN32(ERROR_SERVICE_ALREADY_RUNNING), IsServiceRunning());
	fileMapping.reset(CreateFileMappingW(INVALID_HANDLE_VALUE, nullptr, PAGE_READWRITE, 0, sizeof(ServiceInfo), fileMappingName.data()));
	RETURN_LAST_ERROR_IF_NULL(fileMapping);
	serviceInfo.reset(reinterpret_cast<ServiceInfo*>(MapViewOfFile(fileMapping.get(), FILE_MAP_ALL_ACCESS, 0, 0, 0)));
	RETURN_LAST_ERROR_IF_NULL(serviceInfo);

	RETURN_IF_WIN32_BOOL_FALSE(SetProcessShutdownParameters(0, 0));

	auto primaryRedirector{ make_unique<TBContentRedirector>() };
	vector<HWND>							secondaryRedirectorHostList{};
	vector<unique_ptr<TBContentRedirector>> secondaryRedirectors{};

	// Attach to the primary taskbar if it is existed
	{
		serviceInfo->primaryRedirectorHost = CreateWindowExW(
				WS_EX_NOACTIVATE | WS_EX_NOREDIRECTIONBITMAP | WS_EX_TOOLWINDOW | WS_EX_LAYERED | WS_EX_TRANSPARENT,
				L"Static",
				L"Primary",
				WS_POPUP,
				0, 0, 0, 0,
				nullptr,
				nullptr,
				HINST_THISCOMPONENT,
				nullptr
											 );
		RETURN_LAST_ERROR_IF_NULL(serviceInfo->primaryRedirectorHost);
		RETURN_IF_WIN32_BOOL_FALSE(SetLayeredWindowAttributes(serviceInfo->primaryRedirectorHost, 0, 255, LWA_ALPHA));
		BOOL value{TRUE};
		RETURN_IF_FAILED(DwmSetWindowAttribute(serviceInfo->primaryRedirectorHost, DWMWA_EXCLUDED_FROM_PEEK, &value, sizeof(BOOL)));
		RETURN_IF_FAILED(DwmSetWindowAttribute(serviceInfo->primaryRedirectorHost, DWMWA_DISALLOW_PEEK, &value, sizeof(BOOL)));

		auto callback = [](HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) -> LRESULT
		{
			auto& redirector{*reinterpret_cast<TBContentRedirector*>(dwRefData)};
			if (uMsg == WM_ENDSESSION || uMsg == WM_CTBSTOP)
			{
				redirector.Attach(nullptr, nullptr);
				DestroyWindow(hWnd);
				return 0;
			}
			if (uMsg == WM_DESTROY)
			{
				HWND ui{ GetServiceInfo()->ui };
				if (ui)
				{
					SendMessageW(ui, WM_CLOSE, 0, 0);
				}
				UI::TaskbarDestroyIcon(hWnd);
				PostQuitMessage(0);
			}
			if (uMsg == WM_CTBSHOWUI)
			{
				HWND ui{ GetServiceInfo()->ui };
				if (ui)
				{
					FlashWindow(ui, TRUE);
					ShowWindow(ui, SW_SHOWNORMAL);
					SetActiveWindow(ui);
				}
				else
				{
					thread t{[]()
					{
						if (GetUserDefaultUILanguage() != MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED))
						{
							SetThreadUILanguage(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US));
						}
						return static_cast<LRESULT>(DialogBoxW(HINST_THISCOMPONENT, MAKEINTRESOURCEW(IDD_DIALOG1), nullptr, ClippedTB::UI::DialogProc));
					}};
					t.detach();
				}
				return 0;
			}
			if (uMsg == WM_TASKBARCREATED)
			{
				return ReturnAndCheckHR(UI::TaskbarCreateIcon(hWnd));
			}
			if (uMsg == WM_TASKBARICON)
			{
				if (wParam == IDI_ICON1)
				{
					if (lParam == WM_LBUTTONUP)
					{
						return SendMessageW(hWnd, WM_CTBSHOWUI, 0, 0);
					}
					if (lParam == WM_RBUTTONUP)
					{

					}
				}
				return 0;
			}
			if (uMsg == WM_DWMCOMPOSITIONCHANGED)
			{
				redirector.Attach(redirector.GetSourceTaskbar(), hWnd);
				return 0;
			}
			return DefSubclassProc(hWnd, uMsg, wParam, lParam);
		};
		RETURN_HR_IF(E_FAIL, !SetWindowSubclass(serviceInfo->primaryRedirectorHost, callback, 0, reinterpret_cast<DWORD_PTR>(primaryRedirector.get())));
		RETURN_IF_FAILED(UI::TaskbarCreateIcon(serviceInfo->primaryRedirectorHost));

		HWND primaryTaskbar{ FindWindowW(L"Shell_TrayWnd", nullptr) };
		if (primaryTaskbar)
		{
			RETURN_IF_FAILED(primaryRedirector->Attach(primaryTaskbar, serviceInfo->primaryRedirectorHost));
			RETURN_IF_FAILED(primaryRedirector->Update());
		}
	}

	auto createSecondaryRedirector = [&](HWND taskbar) -> HRESULT
	{
		HWND redirectorHost{nullptr};
		redirectorHost = CreateWindowExW(
							 WS_EX_NOACTIVATE | WS_EX_NOREDIRECTIONBITMAP | WS_EX_TOOLWINDOW | WS_EX_LAYERED | WS_EX_TRANSPARENT,
							 L"Static",
							 L"Primary",
							 WS_POPUP,
							 0, 0, 0, 0,
							 nullptr,
							 nullptr,
							 HINST_THISCOMPONENT,
							 nullptr
						 );
		RETURN_LAST_ERROR_IF_NULL(redirectorHost);
		RETURN_IF_WIN32_BOOL_FALSE(SetLayeredWindowAttributes(redirectorHost, 0, 255, LWA_ALPHA));
		BOOL value{ TRUE };
		RETURN_IF_FAILED(DwmSetWindowAttribute(redirectorHost, DWMWA_EXCLUDED_FROM_PEEK, &value, sizeof(BOOL)));
		RETURN_IF_FAILED(DwmSetWindowAttribute(redirectorHost, DWMWA_DISALLOW_PEEK, &value, sizeof(BOOL)));

		auto callback = [](HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) -> LRESULT
		{
			auto& redirector{ *reinterpret_cast<TBContentRedirector*>(dwRefData) };
			if (uMsg == WM_ENDSESSION || uMsg == WM_CTBSTOP)
			{
				if (redirector.GetRedirectorHost())
				{
					SendMessageW(redirector.GetRedirectorHost(), WM_CLOSE, 0, 0);
				}
				DestroyWindow(hWnd);
				redirector.Attach(nullptr, nullptr);
				return 0;
			}
			if (uMsg == WM_DWMCOMPOSITIONCHANGED)
			{
				redirector.Attach(redirector.GetSourceTaskbar(), hWnd);
				return 0;
			}
			return DefSubclassProc(hWnd, uMsg, wParam, lParam);
		};

		secondaryRedirectors.emplace_back(new TBContentRedirector{});
		auto& newlyCreatedRedirector{ secondaryRedirectors.back() };

		RETURN_HR_IF(E_FAIL, !SetWindowSubclass(redirectorHost, callback, 0, reinterpret_cast<DWORD_PTR>(newlyCreatedRedirector.get())));

		RETURN_IF_FAILED(newlyCreatedRedirector->Attach(taskbar, redirectorHost));
		RETURN_IF_FAILED(newlyCreatedRedirector->Update());

		secondaryRedirectorHostList.push_back(redirectorHost);
		return S_OK;
	};

	// Enumerate all the secondary taskbars and attach
	{
		vector<HWND> secondaryTaskbarList{};
		RETURN_IF_WIN32_BOOL_FALSE(
			EnumWindows([](HWND hWnd, LPARAM lParam) -> BOOL
		{
			WCHAR className[MAX_PATH + 1]{};
			GetClassNameW(hWnd, className, MAX_PATH);
			if (!_wcsicmp(className, L"Shell_SecondaryTrayWnd"))
			{
				auto& taskbarList{ *reinterpret_cast<vector<HWND>*>(lParam) };
				taskbarList.push_back(hWnd);
			}

			return TRUE;
		}, reinterpret_cast<LPARAM>(&secondaryTaskbarList))
		);

		if (!secondaryTaskbarList.empty())
		{
			for (const auto& taskbar : secondaryTaskbarList)
			{
				RETURN_IF_FAILED(createSecondaryRedirector(taskbar));
			}
		}
	}

	MSG msg{};
	while (true)
	{
		if (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE))
		{
			if (msg.message == WM_QUIT)
				break;
			DispatchMessageW(&msg);
		}
		else
		{
			// Primary taskbar
			{
				HWND primaryTaskbar{ FindWindowW(L"Shell_TrayWnd", nullptr) };

				// Taskbar is still there
				if (primaryRedirector->GetSourceTaskbar() && IsWindow(primaryRedirector->GetSourceTaskbar()))
				{
					RETURN_IF_FAILED(primaryRedirector->Update());
				}
				// Redirector is empty and there is a taskbar exist
				else if (primaryTaskbar)
				{
					RETURN_IF_FAILED(primaryRedirector->Attach(primaryTaskbar, serviceInfo->primaryRedirectorHost));
					RETURN_IF_FAILED(primaryRedirector->Update());
				}
			}

			// Secondary taskbars
			{
				// Store unredirected taskbars
				vector<HWND> secondaryTaskbarList{};
				RETURN_IF_WIN32_BOOL_FALSE(
					EnumWindows([](HWND hWnd, LPARAM lParam) -> BOOL
				{
					WCHAR className[MAX_PATH + 1]{};
					GetClassNameW(hWnd, className, MAX_PATH);
					if (!_wcsicmp(className, L"Shell_SecondaryTrayWnd"))
					{
						auto& taskbarList{ *reinterpret_cast<vector<HWND>*>(lParam) };
						taskbarList.push_back(hWnd);
					}

					return TRUE;
				}, reinterpret_cast<LPARAM>(&secondaryTaskbarList))
				);

				// Do some cleanup
				for (auto it = secondaryRedirectorHostList.begin(); it != secondaryRedirectorHostList.end(); it++)
				{
					if (*it == nullptr || !IsWindow(*it))
					{
						it = secondaryRedirectorHostList.erase(it);
						break;
					}
				}

				for (auto it = secondaryRedirectors.begin(); it != secondaryRedirectors.end(); it++)
				{
					auto& redirector{*it};

					// Taskbar is still there
					if (redirector->GetSourceTaskbar() && IsWindow(redirector->GetSourceTaskbar()))
					{
						// This taskbar has already been redirected, erase it
						for (auto& taskbar : secondaryTaskbarList)
						{
							if (taskbar == redirector->GetSourceTaskbar())
							{
								taskbar = nullptr;
							}
						}
						RETURN_IF_FAILED(redirector->Update());
					}
					// Unused redirectors, erase them
					else
					{
						it = secondaryRedirectors.erase(it);
						break;
					}
				}

				// There are still some secondary taskbars,
				// but we have run out all the redirectors
				if (!secondaryTaskbarList.empty())
				{
					for (const auto& taskbar : secondaryTaskbarList)
					{
						if (taskbar && IsWindow(taskbar))
						{
							RETURN_IF_FAILED(createSecondaryRedirector(taskbar));
						}
					}
				}
			}

			DwmFlush();
		}
	}

	for (const auto& redirectorHost : secondaryRedirectorHostList)
	{
		DestroyWindow(redirectorHost);
	}

	serviceInfo->primaryRedirectorHost = nullptr;

	return S_OK;
}

bool ClippedTB::IsServiceRunning()
{
	return wil::unique_handle{ OpenFileMappingW(FILE_MAP_ALL_ACCESS, FALSE, fileMappingName.data()) }.get() != nullptr;
}

ClippedTB::unique_service_info ClippedTB::GetServiceInfo()
{
	wil::unique_handle fileMapping{ OpenFileMappingW(FILE_MAP_ALL_ACCESS, FALSE, fileMappingName.data()) };
	return unique_service_info{ reinterpret_cast<ServiceInfo*>(MapViewOfFile(fileMapping.get(), FILE_MAP_ALL_ACCESS, 0, 0, 0)) };
}

HRESULT ClippedTB::ShowUI()
{
	RETURN_HR_IF(HRESULT_FROM_WIN32(ERROR_SERVICE_NOT_ACTIVE), !IsServiceRunning());
	auto serviceInfo{ GetServiceInfo() };
	RETURN_LAST_ERROR_IF_NULL(serviceInfo);
	RETURN_HR_IF_NULL(HRESULT_FROM_WIN32(ERROR_SERVICE_START_HANG), serviceInfo->primaryRedirectorHost);
	RETURN_HR_IF(HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE), !IsWindow(serviceInfo->primaryRedirectorHost));
	RETURN_LAST_ERROR_IF(!SendNotifyMessageW(serviceInfo->primaryRedirectorHost, WM_CTBSHOWUI, 0, 0));

	return S_OK;
}

HRESULT ClippedTB::StopService()
{
	RETURN_HR_IF(HRESULT_FROM_WIN32(ERROR_SERVICE_NOT_ACTIVE), !IsServiceRunning());
	auto serviceInfo{ GetServiceInfo() };
	RETURN_LAST_ERROR_IF_NULL(serviceInfo);
	RETURN_HR_IF_NULL(HRESULT_FROM_WIN32(ERROR_SERVICE_START_HANG), serviceInfo->primaryRedirectorHost);
	RETURN_HR_IF(HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE), !IsWindow(serviceInfo->primaryRedirectorHost));
	RETURN_LAST_ERROR_IF(!SendNotifyMessageW(serviceInfo->primaryRedirectorHost, WM_CTBSTOP, 0, 0));

	UINT frameCount{ 0 };
	constexpr UINT maxWaitFrameCount{ 1500 };
	while (serviceInfo->primaryRedirectorHost && frameCount < maxWaitFrameCount)
	{
		DwmFlush();
		if ((++frameCount) == maxWaitFrameCount)
		{
			RETURN_HR(HRESULT_FROM_WIN32(ERROR_SERVICE_REQUEST_TIMEOUT));
		}
	}

	return S_OK;
}

HRESULT ClippedTB::Install()
{
	auto CleanUp{ wil::CoInitializeEx() };

	{
		wil::com_ptr<ITaskService> taskService{ wil::CoCreateInstance<ITaskService>(CLSID_TaskScheduler) };
		RETURN_IF_FAILED(taskService->Connect(_variant_t{}, _variant_t{}, _variant_t{}, _variant_t{}));

		wil::com_ptr<ITaskFolder> rootFolder{ nullptr };
		RETURN_IF_FAILED(taskService->GetFolder(_bstr_t("\\"), &rootFolder));

		wil::com_ptr<ITaskDefinition> taskDefinition{ nullptr };
		RETURN_IF_FAILED(taskService->NewTask(0, &taskDefinition));

		wil::com_ptr<IRegistrationInfo> regInfo{ nullptr };
		RETURN_IF_FAILED(taskDefinition->get_RegistrationInfo(&regInfo));
		RETURN_IF_FAILED(regInfo->put_Author(const_cast<BSTR>(L"ALTaleX")));
		RETURN_IF_FAILED(regInfo->put_Description(const_cast<BSTR>(L"This task provides round effect for your taskbars.")));

		{
			wil::com_ptr<IPrincipal> principal{ nullptr };
			RETURN_IF_FAILED(taskDefinition->get_Principal(&principal));

			RETURN_IF_FAILED(principal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN));
			RETURN_IF_FAILED(principal->put_RunLevel(TASK_RUNLEVEL_LUA));
		}

		{
			wil::com_ptr<ITaskSettings> setting{ nullptr };
			RETURN_IF_FAILED(taskDefinition->get_Settings(&setting));

			RETURN_IF_FAILED(setting->put_StopIfGoingOnBatteries(VARIANT_FALSE));
			RETURN_IF_FAILED(setting->put_DisallowStartIfOnBatteries(VARIANT_FALSE));
			RETURN_IF_FAILED(setting->put_AllowDemandStart(VARIANT_TRUE));
			RETURN_IF_FAILED(setting->put_StartWhenAvailable(VARIANT_FALSE));
			RETURN_IF_FAILED(setting->put_MultipleInstances(TASK_INSTANCES_STOP_EXISTING));
		}

		{
			wil::com_ptr<IExecAction> execAction{ nullptr };
			{
				wil::com_ptr<IAction> action{ nullptr };
				{
					wil::com_ptr<IActionCollection> actionColl{ nullptr };
					RETURN_IF_FAILED(taskDefinition->get_Actions(&actionColl));
					RETURN_IF_FAILED(actionColl->Create(TASK_ACTION_EXEC, &action));
				}
				action.query_to(&execAction);
			}

			WCHAR modulePath[MAX_PATH + 1] {};
			RETURN_LAST_ERROR_IF(!GetModuleFileName(HINST_THISCOMPONENT, modulePath, MAX_PATH));

			RETURN_IF_FAILED(
				execAction->put_Path(
					const_cast<BSTR>(modulePath)
				)
			);

			RETURN_IF_FAILED(
				execAction->put_Arguments(
					const_cast<BSTR>(L"")
				)
			);
		}

		wil::com_ptr<ITriggerCollection> triggerColl{ nullptr };
		RETURN_IF_FAILED(taskDefinition->get_Triggers(&triggerColl));

		wil::com_ptr<ITrigger> trigger{ nullptr };
		RETURN_IF_FAILED(triggerColl->Create(TASK_TRIGGER_LOGON, &trigger));

		wil::com_ptr<IRegisteredTask> registeredTask{ nullptr };
		RETURN_IF_FAILED(
			rootFolder->RegisterTaskDefinition(
				const_cast<BSTR>(taskName.data()),
				taskDefinition.get(),
				TASK_CREATE_OR_UPDATE,
				_variant_t{},
				_variant_t{},
				TASK_LOGON_INTERACTIVE_TOKEN,
				_variant_t{},
				&registeredTask
			)
		);
	}

	return S_OK;
}

HRESULT ClippedTB::Uninstall()
{
	auto CleanUp{ wil::CoInitializeEx() };

	{
		wil::com_ptr<ITaskService> taskService{ wil::CoCreateInstance<ITaskService>(CLSID_TaskScheduler) };
		RETURN_IF_FAILED(taskService->Connect(_variant_t{}, _variant_t{}, _variant_t{}, _variant_t{}));

		wil::com_ptr<ITaskFolder> rootFolder{ nullptr };
		RETURN_IF_FAILED(taskService->GetFolder(_bstr_t("\\"), &rootFolder));

		RETURN_IF_FAILED(
			rootFolder->DeleteTask(
				const_cast<BSTR>(taskName.data()),
				0
			)
		);
	}

	WCHAR msg[MAX_PATH];
	THROW_LAST_ERROR_IF(!LoadStringW(HINST_THISCOMPONENT, IDS_STRING131, msg, MAX_PATH));
	if (
		ShellMessageBoxW(
			HINST_THISCOMPONENT,
			GetForegroundWindow(),
			msg,
			nullptr,
			MB_ICONINFORMATION | MB_YESNO
		) == IDYES
	)
	{
		SHDeleteKeyW(
			HKEY_CURRENT_USER, registryItemName.data()
		);
	}

	return S_OK;
}

bool ClippedTB::IsCTBInstalled()
{
	auto CleanUp{ wil::CoInitializeEx() };

	HRESULT hr
	{
		[]()
		{
			wil::com_ptr<ITaskService> taskService{ wil::CoCreateInstance<ITaskService>(CLSID_TaskScheduler) };
			RETURN_IF_FAILED(taskService->Connect(_variant_t{}, _variant_t{}, _variant_t{}, _variant_t{}));

			wil::com_ptr<ITaskFolder> rootFolder{ nullptr };
			RETURN_IF_FAILED(taskService->GetFolder(_bstr_t("\\"), &rootFolder));

			wil::com_ptr<IRegisteredTask> task{nullptr};
			RETURN_IF_FAILED_EXPECTED(rootFolder->GetTask(const_cast<BSTR>(taskName.data()), &task));

			return S_OK;
		}
		()
	};

	return SUCCEEDED(hr);
}

HRESULT ClippedTB::UI::Initialize()
{
#pragma push_macro("max")
#undef max

	HWND ui{GetServiceInfo()->ui};
	if (IsCTBInstalled())
	{
		CheckDlgButton(ui, IDC_CHECK1, BST_CHECKED);
	}
	else
	{
		CheckDlgButton(ui, IDC_CHECK1, BST_UNCHECKED);
	}

	SendDlgItemMessageW(ui, IDC_SLIDER1, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER2, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER3, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER4, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER5, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER6, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER7, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER8, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER9, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER10, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER11, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER12, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER13, TBM_SETRANGE, 0, MAKELPARAM(0, 500));
	SendDlgItemMessageW(ui, IDC_SLIDER14, TBM_SETRANGE, 0, MAKELPARAM(0, 500));

	RETURN_IF_FAILED(ReadRegistry());

#pragma push_macro("max")

	return S_OK;
}

HRESULT ClippedTB::UI::ReadRegistry()
{
	HWND ui{ GetServiceInfo()->ui };

	Config cfg{};
	RETURN_IF_FAILED(ReturnAndCheckHR(cfg.ReadFromRegistry()));

	SetDlgItemInt(ui, IDC_EDIT1, lroundf(cfg.marginsLeft * 10.f), FALSE);
	SetDlgItemInt(ui, IDC_EDIT2, lroundf(cfg.topLeftCornerRadiusX * 10.f), FALSE);
	SetDlgItemInt(ui, IDC_EDIT3, lroundf(cfg.marginsTop * 10.f), FALSE);
	SetDlgItemInt(ui, IDC_EDIT4, lroundf(cfg.marginsRight * 10.f), FALSE);
	SetDlgItemInt(ui, IDC_EDIT5, lroundf(cfg.marginsBottom * 10.f), FALSE);
	SetDlgItemInt(ui, IDC_EDIT6, lroundf(cfg.topLeftCornerRadiusY * 10.f), FALSE);
	SetDlgItemInt(ui, IDC_EDIT7, lroundf(cfg.topRightCornerRadiusX * 10.f), FALSE);
	SetDlgItemInt(ui, IDC_EDIT8, lroundf(cfg.topRightCornerRadiusY * 10.f), FALSE);
	SetDlgItemInt(ui, IDC_EDIT9, lroundf(cfg.bottomLeftCornerRadiusX * 10.f), FALSE);
	SetDlgItemInt(ui, IDC_EDIT10, lroundf(cfg.bottomLeftCornerRadiusY * 10.f), FALSE);
	SetDlgItemInt(ui, IDC_EDIT11, lroundf(cfg.bottomRightCornerRadiusX * 10.f), FALSE);
	SetDlgItemInt(ui, IDC_EDIT12, lroundf(cfg.bottomRightCornerRadiusY * 10.f), FALSE);

	SendDlgItemMessageW(ui, IDC_SLIDER1, TBM_SETPOS, TRUE, lroundf(cfg.marginsLeft * 10.f));
	SendDlgItemMessageW(ui, IDC_SLIDER2, TBM_SETPOS, TRUE, lroundf(cfg.topLeftCornerRadiusX * 10.f));
	SendDlgItemMessageW(ui, IDC_SLIDER3, TBM_SETPOS, TRUE, lroundf(cfg.marginsTop * 10.f));
	SendDlgItemMessageW(ui, IDC_SLIDER4, TBM_SETPOS, TRUE, lroundf(cfg.marginsRight * 10.f));
	SendDlgItemMessageW(ui, IDC_SLIDER5, TBM_SETPOS, TRUE, lroundf(cfg.marginsBottom * 10.f));
	SendDlgItemMessageW(ui, IDC_SLIDER6, TBM_SETPOS, TRUE, lroundf(cfg.topLeftCornerRadiusY * 10.f));
	SendDlgItemMessageW(ui, IDC_SLIDER7, TBM_SETPOS, TRUE, lroundf(cfg.topRightCornerRadiusX * 10.f));
	SendDlgItemMessageW(ui, IDC_SLIDER8, TBM_SETPOS, TRUE, lroundf(cfg.topRightCornerRadiusY * 10.f));
	SendDlgItemMessageW(ui, IDC_SLIDER9, TBM_SETPOS, TRUE, lroundf(cfg.bottomLeftCornerRadiusX * 10.f));
	SendDlgItemMessageW(ui, IDC_SLIDER10, TBM_SETPOS, TRUE, lroundf(cfg.bottomLeftCornerRadiusY * 10.f));
	SendDlgItemMessageW(ui, IDC_SLIDER11, TBM_SETPOS, TRUE, lroundf(cfg.bottomRightCornerRadiusX * 10.f));
	SendDlgItemMessageW(ui, IDC_SLIDER12, TBM_SETPOS, TRUE, lroundf(cfg.bottomRightCornerRadiusY * 10.f));

	return S_OK;
}

INT_PTR CALLBACK ClippedTB::UI::DialogProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
		case WM_NOTIFY:
		{
			if (((LPNMHDR)lParam)->idFrom == IDC_SYSLINK1)
			{
				switch (((LPNMHDR)lParam)->code)
				{
					case NM_CLICK:
					case NM_RETURN:
					{
						auto pNMLink{ (PNMLINK)lParam };
						ShellExecuteW(hDlg, L"open", pNMLink->item.szUrl, nullptr, nullptr, SW_SHOW);
						break;
					}
				}
			}
			if (((LPNMHDR)lParam)->code == TRBN_THUMBPOSCHANGING)
			{
				auto trackbarThumbInfo{(NMTRBTHUMBPOSCHANGING*)lParam};
				switch (((LPNMHDR)lParam)->idFrom)
				{
					case IDC_SLIDER1:
					case IDC_SLIDER2:
					case IDC_SLIDER3:
					case IDC_SLIDER4:
					case IDC_SLIDER5:
					case IDC_SLIDER6:
					case IDC_SLIDER7:
					case IDC_SLIDER8:
					case IDC_SLIDER9:
					case IDC_SLIDER10:
					case IDC_SLIDER11:
					case IDC_SLIDER12:
					{
						auto value{ trackbarThumbInfo->dwPos };

						switch (LOWORD(wParam))
						{
							case IDC_SLIDER1:
							{
								SetDlgItemInt(hDlg, IDC_EDIT1, value, FALSE);
								break;
							}
							case IDC_SLIDER2:
							{
								SetDlgItemInt(hDlg, IDC_EDIT2, value, FALSE);
								break;
							}
							case IDC_SLIDER3:
							{
								SetDlgItemInt(hDlg, IDC_EDIT3, value, FALSE);
								break;
							}
							case IDC_SLIDER4:
							{
								SetDlgItemInt(hDlg, IDC_EDIT4, value, FALSE);
								break;
							}
							case IDC_SLIDER5:
							{
								SetDlgItemInt(hDlg, IDC_EDIT5, value, FALSE);
								break;
							}
							case IDC_SLIDER6:
							{
								SetDlgItemInt(hDlg, IDC_EDIT6, value, FALSE);
								break;
							}
							case IDC_SLIDER7:
							{
								SetDlgItemInt(hDlg, IDC_EDIT7, value, FALSE);
								break;
							}
							case IDC_SLIDER8:
							{
								SetDlgItemInt(hDlg, IDC_EDIT8, value, FALSE);
								break;
							}
							case IDC_SLIDER9:
							{
								SetDlgItemInt(hDlg, IDC_EDIT9, value, FALSE);
								break;
							}
							case IDC_SLIDER10:
							{
								SetDlgItemInt(hDlg, IDC_EDIT10, value, FALSE);
								break;
							}
							case IDC_SLIDER11:
							{
								SetDlgItemInt(hDlg, IDC_EDIT11, value, FALSE);
								break;
							}
							case IDC_SLIDER12:
							{
								SetDlgItemInt(hDlg, IDC_EDIT12, value, FALSE);
								break;
							}
						}

						break;
					}

					case IDC_SLIDER13:
					{
						auto value{ trackbarThumbInfo->dwPos };
						SetDlgItemInt(hDlg, IDC_EDIT13, value, FALSE);

						break;
					}
					case IDC_SLIDER14:
					{
						auto value{ trackbarThumbInfo->dwPos };
						SetDlgItemInt(hDlg, IDC_EDIT14, value, FALSE);

						break;
					}
				}

				return FALSE;
			}
			break;
		}
		case WM_COMMAND:
		{
			switch (LOWORD(wParam))
			{
				case IDC_BUTTON1:
				{
					SendNotifyMessageW(GetServiceInfo()->primaryRedirectorHost, WM_CTBSTOP, 0, 0);
					break;
				}
				case IDC_BUTTON2:
				{
					SHDeleteKeyW(
						HKEY_CURRENT_USER, registryItemName.data()
					);
					RETURN_IF_FAILED(ReturnAndCheckHR(ReadRegistry()));
					break;
				}

				case IDC_EDIT1:
				case IDC_EDIT2:
				case IDC_EDIT3:
				case IDC_EDIT4:
				case IDC_EDIT5:
				case IDC_EDIT6:
				case IDC_EDIT7:
				case IDC_EDIT8:
				case IDC_EDIT9:
				case IDC_EDIT10:
				case IDC_EDIT11:
				case IDC_EDIT12:
				{
					switch (HIWORD(wParam))
					{
						case EN_CHANGE:
						{
							auto value{ GetDlgItemInt(hDlg, LOWORD(wParam), nullptr, FALSE) };

							Config cfg{};
							RETURN_IF_FAILED(ReturnAndCheckHR(cfg.ReadFromRegistry()));
							switch (LOWORD(wParam))
							{
								case IDC_EDIT1:
								{
									cfg.marginsLeft = static_cast<float>(value) / 10.f;
									SendDlgItemMessageW(hDlg, IDC_SLIDER1, TBM_SETPOS, TRUE, value);

									break;
								}
								case IDC_EDIT2:
								{
									cfg.topLeftCornerRadiusX = static_cast<float>(value) / 10.f;
									SendDlgItemMessageW(hDlg, IDC_SLIDER2, TBM_SETPOS, TRUE, value);
									break;
								}
								case IDC_EDIT3:
								{
									cfg.marginsTop = static_cast<float>(value) / 10.f;
									SendDlgItemMessageW(hDlg, IDC_SLIDER3, TBM_SETPOS, TRUE, value);
									break;
								}
								case IDC_EDIT4:
								{
									cfg.marginsRight = static_cast<float>(value) / 10.f;
									SendDlgItemMessageW(hDlg, IDC_SLIDER4, TBM_SETPOS, TRUE, value);
									break;
								}
								case IDC_EDIT5:
								{
									cfg.marginsBottom = static_cast<float>(value) / 10.f;
									SendDlgItemMessageW(hDlg, IDC_SLIDER5, TBM_SETPOS, TRUE, value);
									break;
								}
								case IDC_EDIT6:
								{
									cfg.topLeftCornerRadiusY = static_cast<float>(value) / 10.f;
									SendDlgItemMessageW(hDlg, IDC_SLIDER6, TBM_SETPOS, TRUE, value);
									break;
								}
								case IDC_EDIT7:
								{
									cfg.topRightCornerRadiusX = static_cast<float>(value) / 10.f;
									SendDlgItemMessageW(hDlg, IDC_SLIDER7, TBM_SETPOS, TRUE, value);
									break;
								}
								case IDC_EDIT8:
								{
									cfg.topRightCornerRadiusY = static_cast<float>(value) / 10.f;
									SendDlgItemMessageW(hDlg, IDC_SLIDER8, TBM_SETPOS, TRUE, value);
									break;
								}
								case IDC_EDIT9:
								{
									cfg.bottomLeftCornerRadiusX = static_cast<float>(value) / 10.f;
									SendDlgItemMessageW(hDlg, IDC_SLIDER9, TBM_SETPOS, TRUE, value);
									break;
								}
								case IDC_EDIT10:
								{
									cfg.bottomLeftCornerRadiusY = static_cast<float>(value) / 10.f;
									SendDlgItemMessageW(hDlg, IDC_SLIDER10, TBM_SETPOS, TRUE, value);
									break;
								}
								case IDC_EDIT11:
								{
									cfg.bottomRightCornerRadiusX = static_cast<float>(value) / 10.f;
									SendDlgItemMessageW(hDlg, IDC_SLIDER11, TBM_SETPOS, TRUE, value);
									break;
								}
								case IDC_EDIT12:
								{
									cfg.bottomRightCornerRadiusY = static_cast<float>(value) / 10.f;
									SendDlgItemMessageW(hDlg, IDC_SLIDER12, TBM_SETPOS, TRUE, value);
									break;
								}
							}

							RETURN_IF_FAILED(ReturnAndCheckHR(cfg.WriteToRegistry()));
							break;
						}
					}
					break;
				}
				case IDC_EDIT13:
				{
					switch (HIWORD(wParam))
					{
						case EN_CHANGE:
						{
							auto value{ GetDlgItemInt(hDlg, LOWORD(wParam), nullptr, FALSE) };

							Config cfg{};
							RETURN_IF_FAILED(ReturnAndCheckHR(cfg.ReadFromRegistry()));
							cfg.marginsLeft = value / 10.f;
							cfg.marginsTop = value / 10.f;
							cfg.marginsRight = value / 10.f;
							cfg.marginsBottom = value / 10.f;
							RETURN_IF_FAILED(ReturnAndCheckHR(cfg.WriteToRegistry()));
							RETURN_IF_FAILED(ReturnAndCheckHR(ReadRegistry()));
							SendDlgItemMessageW(hDlg, IDC_SLIDER13, TBM_SETPOS, TRUE, value);

							break;
						}
					}
					break;
				}
				case IDC_EDIT14:
				{
					switch (HIWORD(wParam))
					{
						case EN_CHANGE:
						{
							auto value{ GetDlgItemInt(hDlg, LOWORD(wParam), nullptr, FALSE) };


							Config cfg{};
							RETURN_IF_FAILED(ReturnAndCheckHR(cfg.ReadFromRegistry()));
							cfg.topLeftCornerRadiusX = value / 10.f;
							cfg.topLeftCornerRadiusY = value / 10.f;
							cfg.topRightCornerRadiusX = value / 10.f;
							cfg.topRightCornerRadiusY = value / 10.f;
							cfg.bottomLeftCornerRadiusX = value / 10.f;
							cfg.bottomLeftCornerRadiusY = value / 10.f;
							cfg.bottomRightCornerRadiusX = value / 10.f;
							cfg.bottomRightCornerRadiusY = value / 10.f;
							RETURN_IF_FAILED(ReturnAndCheckHR(cfg.WriteToRegistry()));
							RETURN_IF_FAILED(ReturnAndCheckHR(ReadRegistry()));
							SendDlgItemMessageW(hDlg, IDC_SLIDER14, TBM_SETPOS, TRUE, value);

							break;
						}
					}
					break;
				}

				case IDC_CHECK1:
				{
					if (IsDlgButtonChecked(hDlg, IDC_CHECK1))
					{
						return ReturnAndCheckHR(Install());
					}
					else
					{
						return ReturnAndCheckHR(Uninstall());
					}
					break;
				}
			}
			break;
		}
		case WM_INITDIALOG:
		{
			GetServiceInfo()->ui = hDlg;
			Initialize();
			break;
		}
		case WM_CLOSE:
		{
			GetServiceInfo()->ui = nullptr;
			EndDialog(hDlg, 0);
			break;
		}
		default:
			return FALSE;
	}

	return TRUE;
}

HRESULT ClippedTB::TBContentRedirector::Attach(HWND srcWnd, HWND targetWnd)
{
	if (m_thumbnail)
	{
		DwmUnregisterThumbnail(m_thumbnail);
		m_thumbnail = nullptr;
	}

	// Destruction part
	if (!srcWnd || !targetWnd)
	{
		if (m_targetWnd)
		{
			// Handle window z-order
			RETURN_IF_WIN32_BOOL_FALSE(
				SetWindowPos(
					m_targetWnd, nullptr,
					0, 0, 0, 0,
					SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_NOZORDER | SWP_NOSENDCHANGING | SWP_HIDEWINDOW
				)
			);
			SetWindowLongPtrW(m_targetWnd, GWLP_HWNDPARENT, 0);
		}
		if (m_srcWnd)
		{
			// Show taskbar
			RETURN_IF_WIN32_BOOL_FALSE(SetLayeredWindowAttributes(m_srcWnd, 0, 255, LWA_ALPHA));
			SetWindowLongPtrW(m_srcWnd, GWL_EXSTYLE, GetWindowLongPtrW(m_srcWnd, GWL_EXSTYLE) & ~WS_EX_LAYERED);
		}

		m_thumbnailVisual.reset();
		m_rootVisual.reset();
		m_dcompTarget.reset();
		m_dcompDevice.reset();
		m_dxgiDevice.reset();
		m_d3dDevice.reset();

		m_srcWnd = srcWnd;
		m_targetWnd = targetWnd;

		return S_OK;
	}

	if (!IsCompositionActive())
	{
		return S_OK;
	}

	m_srcWnd = srcWnd;
	m_targetWnd = targetWnd;

	bool needToRecreate{false};
	RETURN_IF_FAILED(
		EnsureDevice(needToRecreate)
	);

	// Hide taskbar
	SetWindowLongPtrW(m_srcWnd, GWL_EXSTYLE, GetWindowLongPtrW(m_srcWnd, GWL_EXSTYLE) | WS_EX_LAYERED);
	RETURN_IF_WIN32_BOOL_FALSE(SetLayeredWindowAttributes(m_srcWnd, 0, 1, LWA_ALPHA));
	// Handle window z-order
	RETURN_IF_WIN32_BOOL_FALSE(
		SetWindowPos(
			targetWnd, HWND_BOTTOM,
			0, 0, 0, 0,
			SWP_NOACTIVATE | SWP_NOZORDER | SWP_NOSENDCHANGING | SWP_HIDEWINDOW
		)
	);
	SetWindowLongPtrW(m_targetWnd, GWLP_HWNDPARENT, reinterpret_cast<LONG_PTR>(m_srcWnd));

	// Initialize thumbnai visual
	RETURN_IF_FAILED(
		m_dcompDevice->CreateTargetForHwnd(m_targetWnd, TRUE, &m_dcompTarget)
	);
	DWM_THUMBNAIL_PROPERTIES thumbnailProperties
	{
		DWM_TNP_VISIBLE | DWM_TNP_RECTDESTINATION | DwmThumbnailAPI::DWM_TNP_ENABLE3D,
		{0, 0, 0, 0},
		{},
		255,
		TRUE,
		FALSE
	};
	RETURN_IF_FAILED(
		DwmThumbnailAPI::g_actualDwmpCreateSharedThumbnailVisual(
			m_targetWnd,
			m_srcWnd,
			0,
			&thumbnailProperties,
			m_dcompDevice.get(),
			m_thumbnailVisual.put_void(),
			&m_thumbnail
		)
	);

	RETURN_IF_FAILED(
		m_dcompDevice->CreateVisual(&m_rootVisual)
	);
	RETURN_IF_FAILED(
		m_rootVisual->AddVisual(m_thumbnailVisual.get(), FALSE, nullptr)
	);
	// Enable antialiasing!
	RETURN_IF_FAILED(
		m_rootVisual->SetCompositeMode(DCOMPOSITION_COMPOSITE_MODE_SOURCE_OVER)
	);
	RETURN_IF_FAILED(
		m_rootVisual->SetBitmapInterpolationMode(DCOMPOSITION_BITMAP_INTERPOLATION_MODE_LINEAR)
	);
	RETURN_IF_FAILED(
		m_rootVisual->SetBorderMode(DCOMPOSITION_BORDER_MODE_SOFT)
	);

	RETURN_IF_FAILED(
		m_dcompTarget->SetRoot(m_rootVisual.get())
	);
	RETURN_IF_FAILED(
		m_dcompDevice->Commit()
	);

	return S_OK;
}

HRESULT ClippedTB::TBContentRedirector::UpdateClipRegion()
{
	if (!m_srcWnd || !IsWindow(m_srcWnd))
	{
		return Attach(nullptr, nullptr);
	}

	if (!IsCompositionActive())
	{
		return S_OK;
	}

	bool needToRecreate{ false };
	RETURN_IF_FAILED(
		EnsureDevice(needToRecreate)
	);
	if (needToRecreate)
	{
		Attach(m_srcWnd, m_targetWnd);
	}

	SIZE size{};
	RETURN_IF_FAILED(
		DwmThumbnailAPI::g_actualDwmpQueryWindowThumbnailSourceSize(m_srcWnd, FALSE, &size)
	);
	float width{static_cast<float>(size.cx)};
	float height{ static_cast<float>(size.cy) };
	constexpr float cornerRadius{30.f};

	wil::com_ptr<IDCompositionRectangleClip> clip{ nullptr };
	RETURN_IF_FAILED(m_dcompDevice->CreateRectangleClip(&clip));
	RETURN_IF_FAILED(clip->SetLeft(m_lastConfig.marginsLeft));
	RETURN_IF_FAILED(clip->SetTop(m_lastConfig.marginsTop));
	RETURN_IF_FAILED(clip->SetRight(width - m_lastConfig.marginsRight));
	RETURN_IF_FAILED(clip->SetBottom(height - m_lastConfig.marginsBottom));
	RETURN_IF_FAILED(clip->SetBottomLeftRadiusX(m_lastConfig.bottomLeftCornerRadiusX));
	RETURN_IF_FAILED(clip->SetBottomLeftRadiusY(m_lastConfig.bottomLeftCornerRadiusY));
	RETURN_IF_FAILED(clip->SetBottomRightRadiusX(m_lastConfig.bottomRightCornerRadiusX));
	RETURN_IF_FAILED(clip->SetBottomRightRadiusY(m_lastConfig.bottomRightCornerRadiusY));
	RETURN_IF_FAILED(clip->SetTopLeftRadiusX(m_lastConfig.topLeftCornerRadiusX));
	RETURN_IF_FAILED(clip->SetTopLeftRadiusY(m_lastConfig.topLeftCornerRadiusY));
	RETURN_IF_FAILED(clip->SetTopRightRadiusX(m_lastConfig.topRightCornerRadiusX));
	RETURN_IF_FAILED(clip->SetTopRightRadiusY(m_lastConfig.topRightCornerRadiusY));

	RETURN_IF_FAILED(
		m_rootVisual->SetClip(clip.get())
	);

	RETURN_IF_FAILED(
		m_dcompDevice->Commit()
	);

	return S_OK;
}

HRESULT ClippedTB::TBContentRedirector::Update()
{
	if (!m_srcWnd || !IsWindow(m_srcWnd))
	{
		return Attach(nullptr, nullptr);
	}

	if (!IsCompositionActive())
	{
		return S_OK;
	}

	bool needToRecreate{ false };
	RETURN_IF_FAILED(
		EnsureDevice(needToRecreate)
	);
	if (needToRecreate)
	{
		Attach(m_srcWnd, m_targetWnd);
	}

	RECT taskbarWindowRect{}, windowRect{};
	RETURN_IF_WIN32_BOOL_FALSE(
		GetWindowRect(m_srcWnd, &taskbarWindowRect)
	);
	RETURN_IF_WIN32_BOOL_FALSE(
		GetWindowRect(m_targetWnd, &windowRect)
	);
	SIZE size{ taskbarWindowRect.right - taskbarWindowRect.left, taskbarWindowRect.bottom - taskbarWindowRect.top };

	bool updateClip{false};

	// User has changed the config information
	Config currentConfig{};
	RETURN_IF_FAILED(currentConfig.ReadFromRegistry());
	if (memcmp(&m_lastConfig, &currentConfig, sizeof(Config)) != 0)
	{
		m_lastConfig = currentConfig;
		updateClip = true;
	}
	// Taskbar visibility changed or window rect changed
	if (auto windowVisible{IsWindowVisible(m_srcWnd)}; m_targetWnd && (!EqualRect(&windowRect, &taskbarWindowRect) || windowVisible != IsWindowVisible(m_targetWnd)))
	{
		RETURN_IF_WIN32_BOOL_FALSE(
			SetWindowPos(
				m_targetWnd, HWND_BOTTOM,
				taskbarWindowRect.left, taskbarWindowRect.top, size.cx, size.cy,
				SWP_NOACTIVATE | SWP_NOZORDER | SWP_NOSENDCHANGING | (windowVisible ? SWP_SHOWWINDOW : SWP_HIDEWINDOW)
			)
		);

		updateClip = true;
	}
	// Window rect changed
	if (m_thumbnail && !EqualRect(&windowRect, &taskbarWindowRect))
	{
		DWM_THUMBNAIL_PROPERTIES thumbnailProperties
		{
			DWM_TNP_VISIBLE | DWM_TNP_RECTDESTINATION | DwmThumbnailAPI::DWM_TNP_ENABLE3D,
			{0, 0, size.cx, size.cy},
			{0, 0, 0, 0},
			255,
			TRUE,
			FALSE
		};
		RETURN_IF_FAILED(DwmUpdateThumbnailProperties(m_thumbnail, &thumbnailProperties));

		updateClip = true;
	}

	if (updateClip)
	{
		UpdateClipRegion();
		updateClip = false;
	}

	// Taskbar cloak state changed
	DWORD taskbarCloakedStatus{0}, cloakedStatus{0};
	RETURN_IF_FAILED(DwmGetWindowAttribute(m_srcWnd, DWMWA_CLOAKED, &taskbarCloakedStatus, sizeof(DWORD)));
	RETURN_IF_FAILED(DwmGetWindowAttribute(m_srcWnd, DWMWA_CLOAKED, &cloakedStatus, sizeof(DWORD)));
	BOOL taskbarCloaked{ taskbarCloakedStatus != 0 }, cloaked{ cloakedStatus != 0 };
	if (taskbarCloaked != cloaked)
	{
		RETURN_IF_FAILED(DwmSetWindowAttribute(m_targetWnd, DWMWA_CLOAK, &taskbarCloaked, sizeof(BOOL)));
	}
	return S_OK;
}

HRESULT ClippedTB::TBContentRedirector::EnsureDevice(bool& needToRecreate)
{
	if (m_dcompDevice)
	{
		// We have already created the related devices,
		// then verify whether it is valid or not
		wil::com_ptr<IDCompositionDevice> dcompDevice{nullptr};
		RETURN_HR_IF(
			E_FAIL,
			!m_dcompDevice.try_query_to(&dcompDevice)
		);
		BOOL valid{FALSE};
		RETURN_IF_FAILED(
			dcompDevice->CheckDeviceState(&valid)
		);

		needToRecreate = !valid;

		if (valid)
		{
			return S_OK;
		}
	}

	RETURN_IF_FAILED(
		D3D11CreateDevice(
			nullptr,
			D3D_DRIVER_TYPE_HARDWARE,
			nullptr,
			D3D11_CREATE_DEVICE_BGRA_SUPPORT | D3D11_CREATE_DEVICE_SINGLETHREADED,
			nullptr,
			0,
			D3D11_SDK_VERSION,
			&m_d3dDevice,
			nullptr,
			nullptr
		)
	);
	RETURN_HR_IF(
		E_FAIL,
		!m_d3dDevice.try_query_to(&m_dxgiDevice)
	);
	RETURN_IF_FAILED(
		DCompositionCreateDevice3(
			m_dxgiDevice.get(),
			IID_PPV_ARGS(&m_dcompDevice)
		)
	);

	return S_OK;
}