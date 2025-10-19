#include <windows.h>
#include <d3d11.h>
#include <dxgi1_2.h>
#include <mfapi.h>
#include <mfidl.h>
#include <mfreadwrite.h>
#include <mferror.h>
#include <chrono>
#include <thread>
#include <vector>
#include <string>
#include <iostream>
#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "dxgi.lib")
#pragma comment(lib, "mfplat.lib")
#pragma comment(lib, "mfreadwrite.lib")
#pragma comment(lib, "mfuuid.lib")

// 安全にリリースするマクロ
#define SAFE_RELEASE(p) if (p) { p->Release(); p = nullptr; }

HRESULT CreateD3DDevice(ID3D11Device** ppDevice, ID3D11DeviceContext** ppContext)
{
    D3D_FEATURE_LEVEL featureLevels[] = {
        D3D_FEATURE_LEVEL_11_1, D3D_FEATURE_LEVEL_11_0
    };
    D3D_FEATURE_LEVEL featureLevel;

    return D3D11CreateDevice(
        nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr,
        D3D11_CREATE_DEVICE_BGRA_SUPPORT,
        featureLevels, ARRAYSIZE(featureLevels),
        D3D11_SDK_VERSION,
        ppDevice, &featureLevel, ppContext
    );
}

HRESULT InitDesktopDuplication(ID3D11Device* device, IDXGIOutputDuplication** ppDuplication, DXGI_OUTPUT_DESC* pOutputDesc)
{
    IDXGIDevice* dxgiDevice = nullptr;
    IDXGIAdapter* dxgiAdapter = nullptr;
    IDXGIOutput* dxgiOutput = nullptr;
    IDXGIOutput1* dxgiOutput1 = nullptr;

    HRESULT hr = device->QueryInterface(__uuidof(IDXGIDevice), (void**)&dxgiDevice);
    if (FAILED(hr)) return hr;

    hr = dxgiDevice->GetParent(__uuidof(IDXGIAdapter), (void**)&dxgiAdapter);
    if (FAILED(hr)) goto cleanup;

    hr = dxgiAdapter->EnumOutputs(0, &dxgiOutput);
    if (FAILED(hr)) goto cleanup;

    hr = dxgiOutput->GetDesc(pOutputDesc);
    if (FAILED(hr)) goto cleanup;

    hr = dxgiOutput->QueryInterface(__uuidof(IDXGIOutput1), (void**)&dxgiOutput1);
    if (FAILED(hr)) goto cleanup;

    hr = dxgiOutput1->DuplicateOutput(device, ppDuplication);

cleanup:
    SAFE_RELEASE(dxgiDevice);
    SAFE_RELEASE(dxgiAdapter);
    SAFE_RELEASE(dxgiOutput);
    SAFE_RELEASE(dxgiOutput1);
    return hr;
}

HRESULT InitSinkWriter(IMFSinkWriter** ppWriter, DWORD* pStreamIndex, UINT32 width, UINT32 height, UINT32 fps)
{
    IMFAttributes* attr = nullptr;
    IMFMediaType* outType = nullptr;
    IMFMediaType* inType = nullptr;
    IMFSinkWriter* writer = nullptr;

    HRESULT hr = MFCreateAttributes(&attr, 1);
    if (FAILED(hr)) return hr;

    hr = attr->SetUINT32(MF_READWRITE_ENABLE_HARDWARE_TRANSFORMS, TRUE);
    if (FAILED(hr)) return hr;

    hr = MFCreateSinkWriterFromURL(L"capture.mp4", nullptr, attr, &writer);
    SAFE_RELEASE(attr);
    if (FAILED(hr)) return hr;

    // 出力 (H.264)
    hr = MFCreateMediaType(&outType);
    if (FAILED(hr)) return hr;
    outType->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Video);
    outType->SetGUID(MF_MT_SUBTYPE, MFVideoFormat_H264);
    outType->SetUINT32(MF_MT_AVG_BITRATE, 8000000);
    outType->SetUINT32(MF_MT_INTERLACE_MODE, MFVideoInterlace_Progressive);
    MFSetAttributeSize(outType, MF_MT_FRAME_SIZE, width, height);
    MFSetAttributeRatio(outType, MF_MT_FRAME_RATE, fps, 1);
    MFSetAttributeRatio(outType, MF_MT_PIXEL_ASPECT_RATIO, 1, 1);
    hr = writer->AddStream(outType, pStreamIndex);
    SAFE_RELEASE(outType);
    if (FAILED(hr)) return hr;

    // 入力 (RGB32)
    hr = MFCreateMediaType(&inType);
    if (FAILED(hr)) return hr;
    inType->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Video);
    inType->SetGUID(MF_MT_SUBTYPE, MFVideoFormat_ARGB32);
    MFSetAttributeSize(inType, MF_MT_FRAME_SIZE, width, height);
    MFSetAttributeRatio(inType, MF_MT_FRAME_RATE, fps, 1);
    MFSetAttributeRatio(inType, MF_MT_PIXEL_ASPECT_RATIO, 1, 1);
    inType->SetUINT32(MF_MT_INTERLACE_MODE, MFVideoInterlace_Progressive);
    hr = writer->SetInputMediaType(*pStreamIndex, inType, nullptr);
    SAFE_RELEASE(inType);
    if (FAILED(hr)) return hr;

    hr = writer->BeginWriting();
    if (FAILED(hr)) return hr;

    *ppWriter = writer;
    return S_OK;
}

int main()
{
    HRESULT hr = MFStartup(MF_VERSION);
    if (FAILED(hr)) {
        std::cerr << "MFStartup failed" << std::endl;
        return -1;
    }

    ID3D11Device* device = nullptr;
    ID3D11DeviceContext* context = nullptr;
    IDXGIOutputDuplication* duplication = nullptr;
    DXGI_OUTPUT_DESC outputDesc = {};
    hr = CreateD3DDevice(&device, &context);
    if (FAILED(hr)) { std::cerr << "D3DDevice creation failed\n"; return -1; }

    hr = InitDesktopDuplication(device, &duplication, &outputDesc);
    if (FAILED(hr)) { std::cerr << "Duplication init failed\n"; return -1; }

    UINT width = outputDesc.DesktopCoordinates.right - outputDesc.DesktopCoordinates.left;
    UINT height = outputDesc.DesktopCoordinates.bottom - outputDesc.DesktopCoordinates.top;

    IMFSinkWriter* writer = nullptr;
    DWORD streamIndex = 0;
    hr = InitSinkWriter(&writer, &streamIndex, width, height, 30);
    if (FAILED(hr)) { std::cerr << "SinkWriter init failed\n"; return -1; }

    std::cout << "Capturing for 5 seconds..." << std::endl;

    auto start = std::chrono::steady_clock::now();
    DWORD frameCount = 0;
    while (true) {
        DXGI_OUTDUPL_FRAME_INFO frameInfo = {};
        IDXGIResource* desktopResource = nullptr;
        hr = duplication->AcquireNextFrame(500, &frameInfo, &desktopResource);
        if (hr == DXGI_ERROR_WAIT_TIMEOUT) continue;
        if (FAILED(hr)) break;

        ID3D11Texture2D* frame = nullptr;
        hr = desktopResource->QueryInterface(__uuidof(ID3D11Texture2D), (void**)&frame);
        SAFE_RELEASE(desktopResource);
        if (FAILED(hr)) break;

        D3D11_TEXTURE2D_DESC desc;
        frame->GetDesc(&desc);

        // CPUアクセス可能なテクスチャにコピー
        D3D11_TEXTURE2D_DESC cpuDesc = desc;
        cpuDesc.Usage = D3D11_USAGE_STAGING;
        cpuDesc.BindFlags = 0;
        cpuDesc.CPUAccessFlags = D3D11_CPU_ACCESS_READ;
        cpuDesc.MiscFlags = 0;

        ID3D11Texture2D* cpuTex = nullptr;
        device->CreateTexture2D(&cpuDesc, nullptr, &cpuTex);
        context->CopyResource(cpuTex, frame);
        SAFE_RELEASE(frame);

        // マップしてMedia Foundationバッファにコピー
        D3D11_MAPPED_SUBRESOURCE mapped;
        hr = context->Map(cpuTex, 0, D3D11_MAP_READ, 0, &mapped);
        if (SUCCEEDED(hr)) {
            IMFMediaBuffer* buffer = nullptr;
            MFCreateMemoryBuffer(mapped.RowPitch * height, &buffer);
            BYTE* dest = nullptr;
            buffer->Lock(&dest, nullptr, nullptr);
            memcpy(dest, mapped.pData, mapped.RowPitch * height);
            buffer->Unlock();
            buffer->SetCurrentLength(mapped.RowPitch * height);

            IMFSample* sample = nullptr;
            MFCreateSample(&sample);
            sample->AddBuffer(buffer);
            sample->SetSampleTime(frameCount * 333333);  // 30fps
            sample->SetSampleDuration(333333);
            writer->WriteSample(streamIndex, sample);

            SAFE_RELEASE(sample);
            SAFE_RELEASE(buffer);
            context->Unmap(cpuTex, 0);
        }

        SAFE_RELEASE(cpuTex);
        duplication->ReleaseFrame();

        frameCount++;

        auto now = std::chrono::steady_clock::now();
        if (std::chrono::duration_cast<std::chrono::seconds>(now - start).count() > 5)
            break;
    }

    writer->Finalize();
    SAFE_RELEASE(writer);
    SAFE_RELEASE(duplication);
    SAFE_RELEASE(context);
    SAFE_RELEASE(device);
    MFShutdown();

    std::cout << "Capture complete. Saved as capture.mp4" << std::endl;
    return 0;
}
