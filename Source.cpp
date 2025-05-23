#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "d2d1")
#pragma comment(lib, "d3d11")
#pragma comment(lib, "dxgi")
#pragma comment(lib, "windowscodecs")
#pragma comment(lib, "dwrite")

#include <windows.h>
#include <wrl/client.h>
#include <d2d1_1.h>
#include <d3d11.h>
#include <wincodec.h>
#include <documenttarget.h>
#include <dwrite.h>

using namespace Microsoft::WRL;

constexpr float DPI = 96.0f;
const TCHAR szClassName[] = TEXT("Window");

HRESULT CreatePrintPageContent(
    ID2D1DeviceContext* deviceContext,
    IDWriteTextFormat* textFormat,
    ID2D1CommandList** outCommandList
)
{
    ComPtr<ID2D1CommandList> commandList;
    HRESULT hr = deviceContext->CreateCommandList(&commandList);
    if (FAILED(hr)) return hr;

    deviceContext->SetTarget(commandList.Get());
    deviceContext->BeginDraw();
    deviceContext->Clear(D2D1::ColorF(D2D1::ColorF::White));

    ComPtr<ID2D1SolidColorBrush> brush;
    hr = deviceContext->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::Black), &brush);
    if (FAILED(hr)) return hr;

    D2D1_RECT_F rect = D2D1::RectF(100, 100, 500, 200);
    deviceContext->DrawTextW(
        L"印刷テストページ",
        (UINT32)wcslen(L"印刷テストページ"),
        textFormat,
        rect,
        brush.Get()
    );

    hr = deviceContext->EndDraw();
    if (FAILED(hr)) return hr;

    hr = commandList->Close();
    if (FAILED(hr)) return hr;

    *outCommandList = commandList.Detach();
    return S_OK;
}

HRESULT PrintWithDirect2D(ID2D1DeviceContext* d2dContext, ID2D1Device* d2dDevice, IWICImagingFactory* wicFactory, IDWriteTextFormat* textFormat, LPCWSTR printerName)
{
    ComPtr<IPrintDocumentPackageTargetFactory> pkgFactory;
    HRESULT hr = CoCreateInstance(__uuidof(PrintDocumentPackageTargetFactory),
        nullptr, CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&pkgFactory));
    if (FAILED(hr)) return hr;

    ComPtr<IPrintDocumentPackageTarget> docTarget;
    hr = pkgFactory->CreateDocumentPackageTargetForPrintJob(
        printerName,
        L"My Print Job",
        nullptr,
        nullptr,
        &docTarget);
    if (FAILED(hr)) return hr;

    ComPtr<ID2D1PrintControl> printControl;
    hr = d2dDevice->CreatePrintControl(wicFactory, docTarget.Get(), nullptr, &printControl);
    if (FAILED(hr)) return hr;

    ComPtr<ID2D1CommandList> cmdList;
    hr = CreatePrintPageContent(d2dContext, textFormat, &cmdList);
    if (FAILED(hr)) return hr;

    float pageWidth = 8.5f * DPI;
    float pageHeight = 11.0f * DPI;

    hr = printControl->AddPage(cmdList.Get(), D2D1::SizeF(pageWidth, pageHeight), nullptr, nullptr, nullptr);
    if (FAILED(hr)) return hr;

    return printControl->Close();
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    static HWND hButton;
    switch (msg)
    {
    case WM_CREATE:
        if (FAILED(CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED))) {
            MessageBox(hWnd, L"COM initialization failed", L"Error", MB_OK);
            return -1;
        }
        hButton = CreateWindow(TEXT("BUTTON"), TEXT("印刷"), WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hWnd, (HMENU)IDOK, ((LPCREATESTRUCT)lParam)->hInstance, 0);
        break;

    case WM_SIZE:
        MoveWindow(hButton, 10, 10, 256, 32, TRUE);
        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK) {
            ComPtr<ID3D11Device> d3dDevice;
            ComPtr<ID3D11DeviceContext> d3dContext;
            HRESULT hr = D3D11CreateDevice(
                nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr, D3D11_CREATE_DEVICE_BGRA_SUPPORT,
                nullptr, 0, D3D11_SDK_VERSION,
                &d3dDevice, nullptr, &d3dContext
            );
            if (FAILED(hr)) {
                MessageBox(hWnd, L"Failed to create D3D11 device", L"Error", MB_OK);
                CoUninitialize();
                return 0;
            }

            ComPtr<IDXGIDevice> dxgiDevice;
            hr = d3dDevice.As(&dxgiDevice);
            if (FAILED(hr)) {
                CoUninitialize();
                return 0;
            }

            ComPtr<ID2D1Factory1> d2dFactory;
            D2D1_FACTORY_OPTIONS options = {};
            options.debugLevel = D2D1_DEBUG_LEVEL_NONE;
            hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, __uuidof(ID2D1Factory1), &options, reinterpret_cast<void**>(d2dFactory.GetAddressOf()));
            if (FAILED(hr)) {
                CoUninitialize();
                return 0;
            }

            ComPtr<ID2D1Device> d2dDevice;
            ComPtr<ID2D1DeviceContext> d2dContext;
            hr = d2dFactory->CreateDevice(dxgiDevice.Get(), &d2dDevice);
            if (FAILED(hr)) {
                CoUninitialize();
                return 0;
            }

            hr = d2dDevice->CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE, &d2dContext);
            if (FAILED(hr)) {
                CoUninitialize();
                return 0;
            }

            ComPtr<IWICImagingFactory> wicFactory;
            hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&wicFactory));
            if (FAILED(hr)) {
                CoUninitialize();
                return 0;
            }

            ComPtr<IDWriteFactory> dwriteFactory;
            ComPtr<IDWriteTextFormat> textFormat;
            hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(dwriteFactory.GetAddressOf()));
            if (FAILED(hr)) {
                CoUninitialize();
                return 0;
            }

            hr = dwriteFactory->CreateTextFormat(
                L"Yu Gothic",
                nullptr,
                DWRITE_FONT_WEIGHT_NORMAL,
                DWRITE_FONT_STYLE_NORMAL,
                DWRITE_FONT_STRETCH_NORMAL,
                24.0f,
                L"ja-jp",
                &textFormat
            );
            if (FAILED(hr)) {
                CoUninitialize();
                return 0;
            }

            LPCWSTR printerName = L"Microsoft Print to PDF";
            hr = PrintWithDirect2D(d2dContext.Get(), d2dDevice.Get(), wicFactory.Get(), textFormat.Get(), printerName);
            if (FAILED(hr)) {
                MessageBox(hWnd, L"Print failed", L"Error", MB_OK);
            }
            else {
                MessageBox(hWnd, L"Print completed successfully", L"Info", MB_OK);
            }
        }
        break;

    case WM_DESTROY:
        CoUninitialize();
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    return 0;
}

int WINAPI wWinMain(
    _In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE,
    _In_ LPWSTR,
    _In_ int nShowCmd)
{
    MSG msg;
    WNDCLASS wndclass = {
        0,
        WndProc,
        0,
        0,
        hInstance,
        0,
        LoadCursor(0, IDC_ARROW),
        (HBRUSH)(COLOR_WINDOW + 1),
        0,
        szClassName
    };
    RegisterClass(&wndclass);

    HWND hWnd = CreateWindow(
        szClassName,
        TEXT("Window"),
        WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
        CW_USEDEFAULT, 0,
        CW_USEDEFAULT, 0,
        0, 0,
        hInstance,
        0
    );

    ShowWindow(hWnd, SW_SHOWDEFAULT);
    UpdateWindow(hWnd);

    while (GetMessage(&msg, 0, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return (int)msg.wParam;
}