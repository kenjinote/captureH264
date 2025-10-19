#include <windows.h>
#include <d3d11.h>
#include <dxgi1_2.h>
#include <initguid.h>
#include <mfapi.h>
#include <mfidl.h>
#include <mfreadwrite.h>
#include <mferror.h>
#include <codecapi.h>
#include <iostream>
#include <thread>
#include <chrono>
#include <roapi.h>

// Media Foundation H.264 Profile
#include <wmcodecdsp.h>

// WinRT / C++/WinRT
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Graphics.Capture.h>
#include <winrt/Windows.Graphics.DirectX.Direct3D11.h>
#include <windows.graphics.directx.direct3d11.interop.h>
#include <d3d11_1.h>
#include <winrt/Windows.Graphics.h>
#include <windows.graphics.capture.interop.h>

#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "dxgi.lib")
#pragma comment(lib, "mfplat.lib")
#pragma comment(lib, "mfreadwrite.lib")
#pragma comment(lib, "mfuuid.lib")
#pragma comment(lib, "WindowsApp.lib")

using namespace winrt;
using namespace winrt::Windows::Foundation;
using namespace winrt::Windows::Graphics::Capture;
using namespace winrt::Windows::Graphics::DirectX::Direct3D11;

// グローバル
IDirect3DDevice g_d3dDevice = nullptr;
GraphicsCaptureSession g_session = nullptr;
Direct3D11CaptureFramePool g_framePool = nullptr;
GraphicsCaptureItem g_captureItem = nullptr;
winrt::com_ptr<ID3D11Device> g_nativeD3D11Device;
winrt::com_ptr<ID3D11DeviceContext> g_nativeD3D11Context;

// 関数宣言
IDirect3DDevice CreateDirect3DDevice();
winrt::com_ptr<ID3D11Texture2D> GetStagingTexture(IDirect3DSurface const& surface, int width, int height);
HRESULT CreateSinkWriterWithD3D(IMFSinkWriter** ppSinkWriter, ID3D11Device* pD3DDevice);
winrt::com_ptr<ID3D11Device> GetD3D11DeviceFromWinrtDevice(IDirect3DDevice const& winrtDevice);

//=============================================================================
// D3D/Interop ヘルパー
//=============================================================================

struct __declspec(uuid("A9B3D012-3DF2-4EE3-B8D1-8695F457D3C1"))
    IDirect3DDxgiInterfaceAccess : ::IUnknown
{
    virtual HRESULT __stdcall GetInterface(REFIID iid, void** p) = 0;
};

winrt::com_ptr<ID3D11Device> GetD3D11DeviceFromWinrtDevice(IDirect3DDevice const& winrtDevice)
{
    winrt::com_ptr<::IDirect3DDxgiInterfaceAccess> access;
    HRESULT hr = reinterpret_cast<::IUnknown*>(winrt::get_abi(winrtDevice))->QueryInterface(
        __uuidof(IDirect3DDxgiInterfaceAccess),
        reinterpret_cast<void**>(access.put())
    );

    if (FAILED(hr) || !access) {
        std::cerr << "Failed to get IDirect3DDxgiInterfaceAccess. HRESULT=0x" << std::hex << hr << std::endl;
        return nullptr;
    }

    winrt::com_ptr<ID3D11Device> d3dDevice;
    hr = access->GetInterface(__uuidof(ID3D11Device), reinterpret_cast<void**>(d3dDevice.put()));
    if (FAILED(hr) || !d3dDevice) {
        std::cerr << "Failed to get ID3D11Device from IDirect3DDxgiInterfaceAccess. HRESULT=0x" << std::hex << hr << std::endl;
        return nullptr;
    }

    return d3dDevice;
}

IDirect3DDevice CreateDirect3DDevice()
{
    ID3D11Device* rawDevice = nullptr;
    ID3D11DeviceContext* rawContext = nullptr;
    D3D_FEATURE_LEVEL featureLevels[] = { D3D_FEATURE_LEVEL_11_1, D3D_FEATURE_LEVEL_11_0 };

    HRESULT hr = D3D11CreateDevice(
        nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr,
        D3D11_CREATE_DEVICE_BGRA_SUPPORT,
        featureLevels, ARRAYSIZE(featureLevels),
        D3D11_SDK_VERSION, &rawDevice, nullptr, &rawContext);

    if (FAILED(hr) || !rawDevice) {
        std::cerr << "D3D11CreateDevice failed. Falling back to WARP." << std::endl;
        hr = D3D11CreateDevice(
            nullptr, D3D_DRIVER_TYPE_WARP, nullptr,
            D3D11_CREATE_DEVICE_BGRA_SUPPORT,
            featureLevels, ARRAYSIZE(featureLevels),
            D3D11_SDK_VERSION, &rawDevice, nullptr, &rawContext);
        if (FAILED(hr) || !rawDevice) {
            throw hresult_error(hr, L"Failed to create D3D11 device (WARP fallback failed).");
        }
    }

    g_nativeD3D11Device.attach(rawDevice);
    g_nativeD3D11Context.attach(rawContext);

    winrt::com_ptr<IDXGIDevice> dxgiDevice;
    g_nativeD3D11Device.as(dxgiDevice);
    if (!dxgiDevice) {
        throw hresult_error(E_FAIL, L"Failed to get IDXGIDevice from native D3D11Device.");
    }

    IDirect3DDevice ret = nullptr;
    check_hresult(CreateDirect3D11DeviceFromDXGIDevice(dxgiDevice.get(), reinterpret_cast<::IInspectable**>(put_abi(ret))));
    return ret;
}

winrt::com_ptr<ID3D11Texture2D> GetStagingTexture(IDirect3DSurface const& surface, int width, int height)
{
    winrt::com_ptr<ID3D11Texture2D> originalTexture;
    surface.as(originalTexture);

    D3D11_TEXTURE2D_DESC desc;
    originalTexture->GetDesc(&desc);

    D3D11_TEXTURE2D_DESC descCopy = desc;
    descCopy.Width = width;
    descCopy.Height = height;
    descCopy.Usage = D3D11_USAGE_STAGING;
    descCopy.BindFlags = 0;
    descCopy.CPUAccessFlags = D3D11_CPU_ACCESS_READ;
    descCopy.MiscFlags = 0;
    descCopy.SampleDesc.Count = 1;
    descCopy.SampleDesc.Quality = 0;
    descCopy.MipLevels = 1;
    descCopy.ArraySize = 1;

    winrt::com_ptr<ID3D11Texture2D> texCopy;
    check_hresult(g_nativeD3D11Device->CreateTexture2D(&descCopy, nullptr, texCopy.put()));
    g_nativeD3D11Context->CopyResource(texCopy.get(), originalTexture.get());
    return texCopy;
}

//=============================================================================
// Media Foundation ヘルパー
//=============================================================================

HRESULT CreateSinkWriterWithD3D(IMFSinkWriter** ppSinkWriter, ID3D11Device* pD3DDevice)
{
    if (!ppSinkWriter) return E_POINTER;
    *ppSinkWriter = nullptr;

    winrt::com_ptr<IMFDXGIDeviceManager> spDXGIDevMgr;
    UINT resetToken = 0;
    HRESULT hr = MFCreateDXGIDeviceManager(&resetToken, spDXGIDevMgr.put());
    if (FAILED(hr)) return hr;

    hr = spDXGIDevMgr->ResetDevice(pD3DDevice, resetToken);
    if (FAILED(hr)) return hr;

    winrt::com_ptr<IMFAttributes> spAttributes;
    hr = MFCreateAttributes(spAttributes.put(), 2);
    if (FAILED(hr)) return hr;

    hr = spAttributes->SetUnknown(MF_SINK_WRITER_D3D_MANAGER, spDXGIDevMgr.get());
    if (FAILED(hr)) return hr;

    hr = spAttributes->SetUINT32(MF_READWRITE_ENABLE_HARDWARE_TRANSFORMS, TRUE);
    if (FAILED(hr)) return hr;

    hr = MFCreateSinkWriterFromURL(L"output.mp4", nullptr, spAttributes.get(), ppSinkWriter);
    return hr;
}

//=============================================================================
// メインキャプチャロジック
//=============================================================================

void Capture(int durationSeconds)
{
    std::cout << "Starting capture for " << durationSeconds << " seconds..." << std::endl;

    g_d3dDevice = CreateDirect3DDevice();
    g_nativeD3D11Device = GetD3D11DeviceFromWinrtDevice(g_d3dDevice);
    if (!g_nativeD3D11Device) {
        throw hresult_error(E_FAIL, L"Failed to obtain native ID3D11Device (Interop failed).");
    }

    g_nativeD3D11Device->GetImmediateContext(g_nativeD3D11Context.put());

    auto itemSize = g_captureItem.Size();
    int width = (static_cast<int>(itemSize.Width) + 15) / 16 * 16;
    int height = (static_cast<int>(itemSize.Height) + 15) / 16 * 16;

    g_framePool = Direct3D11CaptureFramePool::Create(
        g_d3dDevice,
        winrt::Windows::Graphics::DirectX::DirectXPixelFormat::B8G8R8A8UIntNormalized,
        1, itemSize
    );
    g_session = g_framePool.CreateCaptureSession(g_captureItem);
    g_session.IsCursorCaptureEnabled(false);
    g_session.StartCapture();

    winrt::com_ptr<IMFSinkWriter> writer;
    DWORD streamIndex = 0;
    IMFSinkWriter* pSinkWriterRaw = nullptr;
    check_hresult(CreateSinkWriterWithD3D(&pSinkWriterRaw, g_nativeD3D11Device.get()));
    writer.attach(pSinkWriterRaw);

    winrt::com_ptr<IMFMediaType> outType;
    check_hresult(MFCreateMediaType(outType.put()));
    check_hresult(outType->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Video));
    check_hresult(outType->SetGUID(MF_MT_SUBTYPE, MFVideoFormat_H264));
    check_hresult(outType->SetUINT32(MF_MT_AVG_BITRATE, 8000000));
    check_hresult(MFSetAttributeSize(outType.get(), MF_MT_FRAME_SIZE, width, height));
    check_hresult(MFSetAttributeRatio(outType.get(), MF_MT_FRAME_RATE, 30, 1));
    check_hresult(outType->SetUINT32(MF_MT_INTERLACE_MODE, MFVideoInterlace_Progressive));
    check_hresult(outType->SetUINT32(MF_MT_MPEG2_PROFILE, eAVEncH264VProfile_High));
    check_hresult(writer->AddStream(outType.get(), &streamIndex));

    winrt::com_ptr<IMFMediaType> inType;
    check_hresult(MFCreateMediaType(inType.put()));
    check_hresult(inType->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Video));
    check_hresult(inType->SetGUID(MF_MT_SUBTYPE, MFVideoFormat_ARGB32));
    check_hresult(MFSetAttributeSize(inType.get(), MF_MT_FRAME_SIZE, width, height));
    check_hresult(MFSetAttributeRatio(inType.get(), MF_MT_FRAME_RATE, 30, 1));
    check_hresult(inType->SetUINT32(MF_MT_INTERLACE_MODE, MFVideoInterlace_Progressive));
    check_hresult(writer->SetInputMediaType(streamIndex, inType.get(), nullptr));

    check_hresult(writer->BeginWriting());

    const UINT64 frameDuration = 10'000'000ULL / 30ULL;
    UINT64 frameCount = 0;
    auto startTime = std::chrono::steady_clock::now();
    auto endTime = startTime + std::chrono::seconds(durationSeconds);

    std::cout << "Writing frames. Target size: " << width << "x" << height << std::endl;

    while (std::chrono::steady_clock::now() < endTime)
    {
        auto frame = g_framePool.TryGetNextFrame();
        if (frame)
        {
            IDirect3DSurface currentSurface = frame.Surface();
            UINT64 sampleTime = frameCount * frameDuration;

            winrt::com_ptr<ID3D11Texture2D> texCopy = GetStagingTexture(currentSurface, width, height);

            D3D11_MAPPED_SUBRESOURCE mapped = {};
            check_hresult(g_nativeD3D11Context->Map(texCopy.get(), 0, D3D11_MAP_READ, 0, &mapped));

            winrt::com_ptr<IMFMediaBuffer> buffer;
            check_hresult(MFCreateMemoryBuffer(width * height * 4, buffer.put()));

            BYTE* pDest = nullptr;
            DWORD maxLength = 0;
            check_hresult(buffer->Lock(&pDest, &maxLength, nullptr));

            for (int y = 0; y < itemSize.Height; ++y) {
                const BYTE* pSrc = reinterpret_cast<const BYTE*>(mapped.pData) + y * mapped.RowPitch;
                BYTE* pDst = pDest + y * width * 4;
                memcpy(pDst, pSrc, itemSize.Width * 4);
                if (width > itemSize.Width)
                    memset(pDst + itemSize.Width * 4, 0, (width - itemSize.Width) * 4);
            }
            for (int y = static_cast<int>(itemSize.Height); y < height; ++y)
                memset(pDest + y * width * 4, 0, width * 4);

            buffer->Unlock();
            buffer->SetCurrentLength(width * height * 4);
            g_nativeD3D11Context->Unmap(texCopy.get(), 0);

            winrt::com_ptr<IMFSample> sample;
            check_hresult(MFCreateSample(sample.put()));
            check_hresult(sample->AddBuffer(buffer.get()));
            check_hresult(sample->SetSampleTime(sampleTime));
            check_hresult(sample->SetSampleDuration(frameDuration));
            check_hresult(writer->WriteSample(streamIndex, sample.get()));

            frameCount++;
        }
        else std::this_thread::sleep_for(std::chrono::milliseconds(5));
    }

    std::cout << "Finished writing " << frameCount << " frames. Finalizing..." << std::endl;
    check_hresult(writer->Finalize());
    writer = nullptr;

    if (g_session) { g_session.Close(); g_session = nullptr; }
    g_framePool = nullptr;
    g_captureItem = nullptr;

    std::cout << "Capture finished successfully. Output saved to output.mp4 (" << frameCount << " frames)." << std::endl;
}

//=============================================================================
// メイン
//=============================================================================

int main()
{
    winrt::init_apartment(winrt::apartment_type::single_threaded);
    HRESULT mfhr = MFStartup(MF_VERSION);
    if (FAILED(mfhr)) {
        std::cerr << "MFStartup failed: 0x" << std::hex << mfhr << std::dec << std::endl;
        return 1;
    }

    try {
        std::cout << "Attempting to capture primary display directly..." << std::endl;

        POINT ptZero = { 0, 0 };
        HMONITOR hMonitor = MonitorFromPoint(ptZero, MONITOR_DEFAULTTOPRIMARY);

        winrt::com_ptr<IGraphicsCaptureItemInterop> interop =
            winrt::get_activation_factory<winrt::Windows::Graphics::Capture::GraphicsCaptureItem,
            IGraphicsCaptureItemInterop>().as<IGraphicsCaptureItemInterop>();

        ABI::Windows::Graphics::Capture::IGraphicsCaptureItem* rawItemPtr = nullptr;
        check_hresult(interop->CreateForMonitor(
            hMonitor,
            __uuidof(ABI::Windows::Graphics::Capture::IGraphicsCaptureItem),
            reinterpret_cast<void**>(&rawItemPtr)
        ));

        g_captureItem = { rawItemPtr, winrt::take_ownership_from_abi };

        if (!g_captureItem) {
            throw winrt::hresult_error(E_FAIL, L"Failed to create capture item for primary monitor.");
        }

        Capture(5);
    }
    catch (const hresult_error& e) {
        std::wcerr << L"Capture Error (HRESULT=0x" << std::hex << e.code() << L"): " << e.message().c_str() << std::endl;
        return 1;
    }
    catch (const std::exception& e) {
        std::cerr << "Standard Exception: " << e.what() << std::endl;
        return 1;
    }

    MFShutdown();
    winrt::uninit_apartment();
    return 0;
}
