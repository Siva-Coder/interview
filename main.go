package main

import (
	"encoding/binary"
	"encoding/json"
	"fmt"
	"io"
	"math"
	"net/http"
	"net/url"
	"strings"
	"sync"
	"syscall"
	"time"
	"unsafe"

	"github.com/gorilla/websocket"
	"golang.org/x/sys/windows"
)

// ═══════════════════════════════════════════════════════════════════════════
//  CONSTANTS
// ═══════════════════════════════════════════════════════════════════════════

const (
	WDA_EXCLUDEFROMCAPTURE = 0x00000011

	WS_POPUP         = 0x80000000
	WS_VISIBLE       = 0x10000000
	WS_EX_LAYERED    = 0x00080000
	WS_EX_TOOLWINDOW = 0x00000080
	WS_EX_NOACTIVATE = 0x08000000
	WS_EX_TOPMOST    = 0x00000008

	LWA_ALPHA    = 0x00000002
	LWA_COLORKEY = 0x00000001 // ← new: black pixels → see-through

	WM_DESTROY     = 0x0002
	WM_PAINT       = 0x000F
	WM_SIZE        = 0x0005
	WM_KEYDOWN     = 0x0100
	WM_LBUTTONDOWN = 0x0201
	WM_TIMER       = 0x0113
	WM_ACTIVATEAPP = 0x001C
	WM_NCHITTEST   = 0x0084
	WM_MOVE        = 0x0003

	VK_ESCAPE = 0x1B

	DWM_BB_ENABLE = 0x00000001

	WIN_W = 460
	WIN_H = 500 // ← was 680

	// Close button
	CLOSE_X    int32 = WIN_W - 36
	CLOSE_Y    int32 = 14
	CLOSE_SIZE int32 = 20

	// Stop button (left of close)
	STOP_X    int32 = WIN_W - 64
	STOP_Y    int32 = 14
	STOP_SIZE int32 = 20

	// Direct2D / DirectWrite
	D2D1_FACTORY_TYPE_SINGLE_THREADED = 0
	DWRITE_FACTORY_TYPE_SHARED        = 0
	D2D1_RENDER_TARGET_TYPE_DEFAULT   = 0
	D2D1_RENDER_TARGET_USAGE_NONE     = 0
	D2D1_FEATURE_LEVEL_DEFAULT        = 0
	D2D1_ALPHA_MODE_PREMULTIPLIED     = 1
	DWRITE_FONT_WEIGHT_NORMAL         = 400
	DWRITE_FONT_WEIGHT_SEMI_BOLD      = 600
	DWRITE_FONT_WEIGHT_BOLD           = 700
	DWRITE_FONT_STYLE_NORMAL          = 0
	DWRITE_FONT_STRETCH_NORMAL        = 5
	DWRITE_TEXT_ALIGNMENT_LEADING     = 0
	DWRITE_PARAGRAPH_ALIGNMENT_NEAR   = 0

	DWRITE_RENDERING_MODE_CLEARTYPE_NATURAL_SYMMETRIC = 6
	D2D1_DRAW_TEXT_OPTIONS_NONE                       = 0
	DWRITE_MEASURING_MODE_NATURAL                     = 0

	// WASAPI
	CLSCTX_ALL                   = 23
	AUDCLNT_SHAREMODE_SHARED     = 0
	AUDCLNT_BUFFERFLAGS_SILENT   = 0x00000002
	AUDCLNT_STREAMFLAGS_LOOPBACK = 0x00020000
	S_OK                         = 0

	// AssemblyAI
	ASSEMBLYAI_WS_URL = "wss://streaming.assemblyai.com/v3/ws"
	SAMPLE_RATE       = 16000
)

// ═══════════════════════════════════════════════════════════════════════════
//  AssemblyAI Message Types  (UNCHANGED)
// ═══════════════════════════════════════════════════════════════════════════

type AAIBaseMessage struct {
	Type string `json:"type"`
}

type AAIBeginMessage struct {
	Type      string  `json:"type"`
	ID        string  `json:"id"`
	ExpiresAt float64 `json:"expires_at"`
}

type AAITurnMessage struct {
	Type                string  `json:"type"`
	TurnOrder           int     `json:"turn_order"`
	TurnIsFormatted     bool    `json:"turn_is_formatted"`
	EndOfTurn           bool    `json:"end_of_turn"`
	Transcript          string  `json:"transcript"`
	EndOfTurnConfidence float64 `json:"end_of_turn_confidence"`
}

type AAITerminationMessage struct {
	Type                   string `json:"type"`
	AudioDurationSeconds   int    `json:"audio_duration_seconds"`
	SessionDurationSeconds int    `json:"session_duration_seconds"`
}

type AAITerminateMessage struct {
	Type string `json:"type"`
}

// ═══════════════════════════════════════════════════════════════════════════
//  COM GUIDs  (UNCHANGED)
// ═══════════════════════════════════════════════════════════════════════════

var (
	IID_ID2D1Factory = windows.GUID{
		Data1: 0x06152247, Data2: 0x6f50, Data3: 0x465a,
		Data4: [8]byte{0x92, 0x45, 0x11, 0x8b, 0xfd, 0x3b, 0x60, 0x07},
	}
	IID_IDWriteFactory = windows.GUID{
		Data1: 0xb859ee5a, Data2: 0xd838, Data3: 0x4b5b,
		Data4: [8]byte{0xa2, 0xe8, 0x1a, 0xdc, 0x7d, 0x93, 0xdb, 0x48},
	}
	CLSID_MMDeviceEnumerator = windows.GUID{
		Data1: 0xbcde0395, Data2: 0xe52f, Data3: 0x467c,
		Data4: [8]byte{0x8e, 0x3d, 0xc4, 0x57, 0x92, 0x91, 0x69, 0x2e},
	}
	IID_IMMDeviceEnumerator = windows.GUID{
		Data1: 0xa95664d2, Data2: 0x9614, Data3: 0x4f35,
		Data4: [8]byte{0xa7, 0x46, 0xde, 0x8d, 0xb6, 0x36, 0x17, 0xe6},
	}
	IID_IAudioClient = windows.GUID{
		Data1: 0x1cb9ad4c, Data2: 0xdbfa, Data3: 0x4c32,
		Data4: [8]byte{0xb1, 0x78, 0xc2, 0xf5, 0x68, 0xa7, 0x03, 0xb2},
	}
	IID_IAudioCaptureClient = windows.GUID{
		Data1: 0xc8adbd64, Data2: 0xe71e, Data3: 0x48a0,
		Data4: [8]byte{0xa4, 0xde, 0x18, 0x5c, 0x39, 0x5c, 0xd3, 0x17},
	}
)

// ═══════════════════════════════════════════════════════════════════════════
//  COM vtable structs  (UNCHANGED)
// ═══════════════════════════════════════════════════════════════════════════

type iUnknownVtbl struct {
	QueryInterface uintptr
	AddRef         uintptr
	Release        uintptr
}

type id2d1FactoryVtbl struct {
	iUnknownVtbl
	ReloadSystemMetrics            uintptr
	GetDesktopDpi                  uintptr
	CreateRectangleGeometry        uintptr
	CreateRoundedRectangleGeometry uintptr
	CreateEllipseGeometry          uintptr
	CreateGeometryGroup            uintptr
	CreateTransformedGeometry      uintptr
	CreatePathGeometry             uintptr
	CreateStrokeStyle              uintptr
	CreateDrawingStateBlock        uintptr
	CreateWicBitmapRenderTarget    uintptr
	CreateHwndRenderTarget         uintptr
}

type id2d1RenderTargetVtbl struct {
	iUnknownVtbl
	GetFactory                   uintptr // 3
	CreateBitmap                 uintptr // 4
	CreateBitmapFromWicBitmap    uintptr // 5
	CreateSharedBitmap           uintptr // 6
	CreateBitmapBrush            uintptr // 7
	CreateSolidColorBrush        uintptr // 8
	CreateGradientStopCollection uintptr // 9
	CreateLinearGradientBrush    uintptr // 10
	CreateRadialGradientBrush    uintptr // 11
	CreateBitmapRenderTarget     uintptr // 12
	CreateLayerRenderTarget      uintptr // 13
	CreateMesh                   uintptr // 14
	DrawLine                     uintptr // 15
	DrawRectangle                uintptr // 16
	FillRectangle                uintptr // 17
	DrawRoundedRectangle         uintptr // 18
	FillRoundedRectangle         uintptr // 19
	DrawEllipse                  uintptr // 20
	FillEllipse                  uintptr // 21
	DrawGeometry                 uintptr // 22
	FillGeometry                 uintptr // 23
	FillMesh                     uintptr // 24
	FillOpacityMask              uintptr // 25
	DrawBitmap                   uintptr // 26
	DrawText                     uintptr // 27
	DrawTextLayout               uintptr // 28
	DrawGlyphRun                 uintptr // 29
	SetTransform                 uintptr // 30
	GetTransform                 uintptr // 31
	SetAntialiasMode             uintptr // 32
	GetAntialiasMode             uintptr // 33
	SetTextAntialiasMode         uintptr // 34
	GetTextAntialiasMode         uintptr // 35
	SetTextRenderingParams       uintptr // 36
	GetTextRenderingParams       uintptr // 37
	SetTags                      uintptr // 38
	GetTags                      uintptr // 39
	PushLayer                    uintptr // 40
	PopLayer                     uintptr // 41
	Flush                        uintptr // 42
	SaveDrawingState             uintptr // 43
	RestoreDrawingState          uintptr // 44
	PushAxisAlignedClip          uintptr // 45
	PopAxisAlignedClip           uintptr // 46
	Clear                        uintptr // 47
	BeginDraw                    uintptr // 48
	EndDraw                      uintptr // 49
	GetPixelFormat               uintptr // 50
	SetDpi                       uintptr // 51
	GetDpi                       uintptr // 52
	GetSize                      uintptr // 53
	GetPixelSize                 uintptr // 54
	GetMaximumBitmapSize         uintptr // 55
	IsSupported                  uintptr // 56
}

type idWriteFactoryVtbl struct {
	iUnknownVtbl
	GetSystemFontCollection        uintptr
	CreateCustomFontCollection     uintptr
	RegisterFontCollectionLoader   uintptr
	UnregisterFontCollectionLoader uintptr
	CreateFontFileReference        uintptr
	CreateCustomFontFileReference  uintptr
	CreateFontFace                 uintptr
	CreateRenderingParams          uintptr
	CreateMonitorRenderingParams   uintptr
	CreateCustomRenderingParams    uintptr
	RegisterFontFileLoader         uintptr
	UnregisterFontFileLoader       uintptr
	CreateTextFormat               uintptr
	CreateTypography               uintptr
	GetGdiInterop                  uintptr
	CreateTextLayout               uintptr
}

type idWriteTextFormatVtbl struct {
	iUnknownVtbl
	SetTextAlignment      uintptr
	SetParagraphAlignment uintptr
}

type immDeviceEnumeratorVtbl struct {
	iUnknownVtbl
	EnumAudioEndpoints                     uintptr
	GetDefaultAudioEndpoint                uintptr
	GetDevice                              uintptr
	RegisterEndpointNotificationCallback   uintptr
	UnregisterEndpointNotificationCallback uintptr
}

type immDeviceVtbl struct {
	iUnknownVtbl
	Activate          uintptr
	OpenPropertyStore uintptr
	GetId             uintptr
	GetState          uintptr
}

type iAudioClientVtbl struct {
	iUnknownVtbl
	Initialize        uintptr
	GetBufferSize     uintptr
	GetStreamLatency  uintptr
	GetCurrentPadding uintptr
	IsFormatSupported uintptr
	GetMixFormat      uintptr
	GetDevicePeriod   uintptr
	Start             uintptr
	Stop              uintptr
	Reset             uintptr
	SetEventHandle    uintptr
	GetService        uintptr
}

type iAudioCaptureClientVtbl struct {
	iUnknownVtbl
	GetBuffer         uintptr
	ReleaseBuffer     uintptr
	GetNextPacketSize uintptr
}

// ═══════════════════════════════════════════════════════════════════════════
//  Structs  (UNCHANGED)
// ═══════════════════════════════════════════════════════════════════════════

type D2D1_COLOR_F struct{ R, G, B, A float32 }
type D2D1_RECT_F struct{ Left, Top, Right, Bottom float32 }
type D2D1_ELLIPSE struct {
	Point            D2D_POINT_2F
	RadiusX, RadiusY float32
}
type D2D_POINT_2F struct{ X, Y float32 }
type D2D1_SIZE_U struct{ Width, Height uint32 }
type D2D1_PIXEL_FORMAT struct {
	Format    uint32
	AlphaMode uint32
}
type D2D1_RENDER_TARGET_PROPERTIES struct {
	RenderTargetType uint32
	PixelFormat      D2D1_PIXEL_FORMAT
	DpiX, DpiY       float32
	Usage            uint32
	MinLevel         uint32
}
type D2D1_HWND_RENDER_TARGET_PROPERTIES struct {
	Hwnd        uintptr
	PixelSize   D2D1_SIZE_U
	PresentOpts uint32
}
type DWM_BLURBEHIND struct {
	DwFlags                uint32
	FEnable                int32
	HRgnBlur               uintptr
	FTransitionOnMaximized int32
}

type WAVEFORMATEX struct {
	WFormatTag      uint16
	NChannels       uint16
	NSamplesPerSec  uint32
	NAvgBytesPerSec uint32
	NBlockAlign     uint16
	WBitsPerSample  uint16
	CBSize          uint16
}

type WNDCLASSEXW struct {
	CbSize        uint32
	Style         uint32
	LpfnWndProc   uintptr
	CbClsExtra    int32
	CbWndExtra    int32
	HInstance     windows.Handle
	HIcon         windows.Handle
	HCursor       windows.Handle
	HbrBackground windows.Handle
	LpszMenuName  *uint16
	LpszClassName *uint16
	HIconSm       windows.Handle
}
type POINT struct{ X, Y int32 }
type MSG struct {
	Hwnd    windows.Handle
	Message uint32
	WParam  uintptr
	LParam  uintptr
	Time    uint32
	Pt      POINT
}
type PAINTSTRUCT struct {
	Hdc         windows.Handle
	Erase       int32
	RcPaint     RECT
	Restore     int32
	IncUpdate   int32
	RgbReserved [32]byte
}
type RECT struct{ Left, Top, Right, Bottom int32 }

// ═══════════════════════════════════════════════════════════════════════════
//  DLLs & procs  (UNCHANGED)
// ═══════════════════════════════════════════════════════════════════════════

var (
	user32    = windows.NewLazySystemDLL("user32.dll")
	kernel32  = windows.NewLazySystemDLL("kernel32.dll")
	gdi32     = windows.NewLazySystemDLL("gdi32.dll")
	dwmapi    = windows.NewLazySystemDLL("dwmapi.dll")
	d2d1dll   = windows.NewLazySystemDLL("d2d1.dll")
	dwritedll = windows.NewLazySystemDLL("dwrite.dll")
	ole32     = windows.NewLazySystemDLL("ole32.dll")

	procSetWindowDisplayAffinity   = user32.NewProc("SetWindowDisplayAffinity")
	procCreateWindowExW            = user32.NewProc("CreateWindowExW")
	procRegisterClassExW           = user32.NewProc("RegisterClassExW")
	procDefWindowProcW             = user32.NewProc("DefWindowProcW")
	procShowWindow                 = user32.NewProc("ShowWindow")
	procUpdateWindow               = user32.NewProc("UpdateWindow")
	procSetLayeredWindowAttributes = user32.NewProc("SetLayeredWindowAttributes")
	procGetMessageW                = user32.NewProc("GetMessageW")
	procTranslateMessage           = user32.NewProc("TranslateMessage")
	procDispatchMessageW           = user32.NewProc("DispatchMessageW")
	procPostQuitMessage            = user32.NewProc("PostQuitMessage")
	procGetClientRect              = user32.NewProc("GetClientRect")
	procSetWindowPos               = user32.NewProc("SetWindowPos")
	procGetSystemMetrics           = user32.NewProc("GetSystemMetrics")
	procBeginPaint                 = user32.NewProc("BeginPaint")
	procEndPaint                   = user32.NewProc("EndPaint")
	procSetTimer                   = user32.NewProc("SetTimer")
	procGetWindowRect              = user32.NewProc("GetWindowRect")
	procGetCursorPos               = user32.NewProc("GetCursorPos")
	procInvalidateRect             = user32.NewProc("InvalidateRect")
	procSetWindowRgn               = user32.NewProc("SetWindowRgn")
	procCreateRoundRectRgn         = gdi32.NewProc("CreateRoundRectRgn")
	procCreateSolidBrush           = gdi32.NewProc("CreateSolidBrush")
	procFillRect                   = user32.NewProc("FillRect")
	procDeleteObject               = gdi32.NewProc("DeleteObject")
	procCreateEllipseRgn           = gdi32.NewProc("CreateEllipseRgn")

	procDwmEnableBlurBehindWindow = dwmapi.NewProc("DwmEnableBlurBehindWindow")
	procD2D1CreateFactory         = d2d1dll.NewProc("D2D1CreateFactory")
	procDWriteCreateFactory       = dwritedll.NewProc("DWriteCreateFactory")
	procCoInitializeEx            = ole32.NewProc("CoInitializeEx")
	procCoCreateInstance          = ole32.NewProc("CoCreateInstance")
	procCoTaskMemFree             = ole32.NewProc("CoTaskMemFree")

	HWND_TOPMOST   = ^uintptr(0)
	SWP_SHOWWINDOW = uintptr(0x0040)
	SWP_NOMOVE     = uintptr(0x0002)
	SWP_NOSIZE     = uintptr(0x0001)
	SW_SHOW        = uintptr(5)
	TIMER_ID       = uintptr(1)
	TIMER_AI       = uintptr(2)
	TIMER_AUDIO    = uintptr(3)
)

// ═══════════════════════════════════════════════════════════════════════════
//  Color helpers — green-only palette + black for colorkey transparency
// ═══════════════════════════════════════════════════════════════════════════

func rgba(r, g, b uint8, a float32) D2D1_COLOR_F {
	return D2D1_COLOR_F{R: float32(r) / 255.0, G: float32(g) / 255.0, B: float32(b) / 255.0, A: a}
}

func rectF(l, t, r, b float32) D2D1_RECT_F { return D2D1_RECT_F{l, t, r, b} }

var (
	// ← opaque black = transparent hole via LWA_COLORKEY
	clrBg          = rgba(0x0D, 0x0D, 0x18, 1.0)
	clrGreen       = rgba(0x34, 0xd3, 0x99, 1) // primary green — labels, AI text
	clrGreenDim    = rgba(0x10, 0x99, 0x66, 1) // dimmer green — "You" label, divider
	clrGreenMuted  = rgba(0x0a, 0x5c, 0x40, 1) // very dim green — partial text
	clrTextMain    = rgba(0xf0, 0xf4, 0xff, 1) // near-white — user transcript text
	clrTextDim     = rgba(0x9a, 0xa2, 0xb8, 1) // dim white — timestamps, meta
	clrRed         = rgba(0xf8, 0x71, 0x71, 1) // close button hover
	clrTransparent = rgba(0, 0, 0, 0)
)

// ═══════════════════════════════════════════════════════════════════════════
//  D2D/DWrite state
// ═══════════════════════════════════════════════════════════════════════════

var (
	pD2DFactory    uintptr
	pDWriteFactory uintptr
	pRenderTarget  uintptr

	brushBg, brushGreen, brushGreenDim, brushGreenMuted uintptr
	brushTextMain, brushTextDim, brushRed               uintptr

	fmtTitle, fmtSub, fmtBody, fmtBadge, fmtSmall uintptr

	hwndMain   windows.Handle
	windowX    int32
	windowY    int32
	closeHover bool
	stopHover  bool // ← new
	appStopped bool // ← new: tracks if user manually stopped
)

// ═══════════════════════════════════════════════════════════════════════════
//  App State  (UNCHANGED)
// ═══════════════════════════════════════════════════════════════════════════

type Message struct {
	Role      string
	Text      string
	Timestamp time.Time
}

var (
	conversation      []Message
	convMutex         sync.RWMutex
	isListening       bool
	isProcessing      bool
	isAIResponding    bool
	currentTranscript string
	transcriptMutex   sync.RWMutex
	partialTranscript string
	partialMutex      sync.RWMutex

	pAudioClient   uintptr
	pCaptureClient uintptr
	audioFormat    *WAVEFORMATEX
	captureRunning bool
	audioBuffer    []byte
	audioMutex     sync.Mutex

	aaiWSConn     *websocket.Conn
	aaiConnected  bool
	aaiMutex      sync.RWMutex
	assemblyAIKey string
)

// ═══════════════════════════════════════════════════════════════════════════
//  COM helpers  (UNCHANGED)
// ═══════════════════════════════════════════════════════════════════════════

func vtbl(obj uintptr, slot int) uintptr {
	vtable := *(*uintptr)(unsafe.Pointer(obj))
	return *(*uintptr)(unsafe.Pointer(vtable + uintptr(slot)*unsafe.Sizeof(uintptr(0))))
}

func comCall(obj uintptr, slot int, args ...uintptr) uintptr {
	fn := vtbl(obj, slot)
	all := make([]uintptr, 0, len(args)+1)
	all = append(all, obj)
	all = append(all, args...)
	ret, _, _ := syscall.SyscallN(fn, all...)
	return ret
}

// ═══════════════════════════════════════════════════════════════════════════
//  Direct2D helpers  (UNCHANGED)
// ═══════════════════════════════════════════════════════════════════════════

func createBrush(rt uintptr, c D2D1_COLOR_F) uintptr {
	var brush uintptr
	comCall(rt, 8, uintptr(unsafe.Pointer(&c)), 0, uintptr(unsafe.Pointer(&brush)))
	return brush
}

func d2dFillRect(rt, brush uintptr, rc D2D1_RECT_F) {
	comCall(rt, 17, uintptr(unsafe.Pointer(&rc)), brush)
}

func d2dFillEllipse(rt, brush uintptr, cx, cy, rx, ry float32) {
	e := D2D1_ELLIPSE{Point: D2D_POINT_2F{cx, cy}, RadiusX: rx, RadiusY: ry}
	comCall(rt, 21, uintptr(unsafe.Pointer(&e)), brush)
}

func d2dText(rt uintptr, text string, fmt_, brush uintptr, rc D2D1_RECT_F) {
	wstr, _ := syscall.UTF16FromString(text)
	comCall(rt, 27,
		uintptr(unsafe.Pointer(&wstr[0])),
		uintptr(len(wstr)-1),
		fmt_,
		uintptr(unsafe.Pointer(&rc)),
		brush,
		D2D1_DRAW_TEXT_OPTIONS_NONE,
		DWRITE_MEASURING_MODE_NATURAL,
	)
}

func createTextFormat(family string, weight, style, stretch uint32, size float32) uintptr {
	familyW, _ := syscall.UTF16PtrFromString(family)
	localeW, _ := syscall.UTF16PtrFromString("en-us")
	var pFmt uintptr
	comCall(pDWriteFactory, 15,
		uintptr(unsafe.Pointer(familyW)), 0,
		uintptr(weight), uintptr(style), uintptr(stretch),
		uintptr(math_float32bits(size)),
		uintptr(unsafe.Pointer(localeW)),
		uintptr(unsafe.Pointer(&pFmt)),
	)
	return pFmt
}

func math_float32bits(f float32) uint32 { return *(*uint32)(unsafe.Pointer(&f)) }

// ═══════════════════════════════════════════════════════════════════════════
//  Direct2D Initialization
// ═══════════════════════════════════════════════════════════════════════════

func initD2D(hwnd uintptr) error {
	hr, _, _ := procD2D1CreateFactory.Call(
		D2D1_FACTORY_TYPE_SINGLE_THREADED,
		uintptr(unsafe.Pointer(&IID_ID2D1Factory)),
		0, uintptr(unsafe.Pointer(&pD2DFactory)),
	)
	if hr != 0 {
		return fmt.Errorf("D2D1CreateFactory: 0x%X", hr)
	}

	hr, _, _ = procDWriteCreateFactory.Call(
		DWRITE_FACTORY_TYPE_SHARED,
		uintptr(unsafe.Pointer(&IID_IDWriteFactory)),
		uintptr(unsafe.Pointer(&pDWriteFactory)),
	)
	if hr != 0 {
		return fmt.Errorf("DWriteCreateFactory: 0x%X", hr)
	}

	rtProps := D2D1_RENDER_TARGET_PROPERTIES{
		RenderTargetType: D2D1_RENDER_TARGET_TYPE_DEFAULT,
		PixelFormat:      D2D1_PIXEL_FORMAT{Format: 87, AlphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED},
		DpiX:             96, DpiY: 96, Usage: D2D1_RENDER_TARGET_USAGE_NONE,
	}
	hwndRtProps := D2D1_HWND_RENDER_TARGET_PROPERTIES{
		Hwnd: hwnd, PixelSize: D2D1_SIZE_U{Width: WIN_W, Height: WIN_H},
	}
	hr = comCall(pD2DFactory, 14,
		uintptr(unsafe.Pointer(&rtProps)),
		uintptr(unsafe.Pointer(&hwndRtProps)),
		uintptr(unsafe.Pointer(&pRenderTarget)),
	)
	if hr != 0 {
		return fmt.Errorf("CreateHwndRenderTarget: 0x%X", hr)
	}

	// ← Use Grayscale (2) instead of ClearType (1): required for transparent backgrounds
	comCall(pRenderTarget, 34, 1)

	// Only the brushes we actually use in the new UI
	brushBg = createBrush(pRenderTarget, clrBg)
	brushGreen = createBrush(pRenderTarget, clrGreen)
	brushGreenDim = createBrush(pRenderTarget, clrGreenDim)
	brushGreenMuted = createBrush(pRenderTarget, clrGreenMuted)
	brushTextMain = createBrush(pRenderTarget, clrTextMain)
	brushTextDim = createBrush(pRenderTarget, clrTextDim)
	brushRed = createBrush(pRenderTarget, clrRed)

	fmtTitle = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 17)
	fmtSub = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 12)
	fmtBody = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 13)
	fmtBadge = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_SEMI_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 10)
	fmtSmall = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 10)

	return nil
}

// ═══════════════════════════════════════════════════════════════════════════
//  RENDERING — transparent chatbot UI
// ═══════════════════════════════════════════════════════════════════════════

func renderFrame() {
	if pRenderTarget == 0 {
		return
	}
	W := float32(WIN_W)
	H := float32(WIN_H)

	comCall(pRenderTarget, 48) // BeginDraw

	// Clear with opaque black → LWA_COLORKEY punches it to fully transparent
	// Only drawn text/lines remain visible; background shows through completely
	comCall(pRenderTarget, 47, uintptr(unsafe.Pointer(&clrBg))) // Clear

	// ── Header ──────────────────────────────────────────────────────────────
	// Title
	d2dText(pRenderTarget, "parakeet", fmtTitle, brushGreen, rectF(18, 12, 120, 42))
	d2dText(pRenderTarget, " ai", fmtTitle, brushGreenDim, rectF(114, 12, 160, 42))

	// Live status dot — green if connected, dim if not
	aaiMutex.RLock()
	aaiConn := aaiConnected
	aaiMutex.RUnlock()
	dotBrush := brushGreenMuted
	if isListening && aaiConn {
		dotBrush = brushGreen
	}
	d2dFillEllipse(pRenderTarget, dotBrush, 170, 27, 4, 4)

	// Stop button ■ — green when active, dim when stopped
	stopBrush := brushGreenDim
	if appStopped {
		stopBrush = brushTextDim
	}
	if stopHover && !appStopped {
		stopBrush = brushGreen
	}
	d2dText(pRenderTarget, "■",
		fmtTitle, stopBrush,
		rectF(float32(STOP_X)-1, float32(STOP_Y), float32(STOP_X)+float32(STOP_SIZE)+2, float32(STOP_Y)+float32(STOP_SIZE)+2))

	// Close button × — dim normally, red on hover
	closeBrush := brushTextDim
	if closeHover {
		closeBrush = brushRed
	}
	d2dText(pRenderTarget, "×",
		fmtTitle, closeBrush,
		rectF(float32(CLOSE_X), float32(CLOSE_Y)-1, float32(CLOSE_X)+float32(CLOSE_SIZE)+4, float32(CLOSE_Y)+float32(CLOSE_SIZE)+2))

	// Thin green divider under header
	d2dFillRect(pRenderTarget, brushGreenDim, rectF(18, 48, W-18, 49))

	// ── Chat messages ────────────────────────────────────────────────────────
	yPos := float32(60)
	lineH := float32(20)  // body line height
	labelH := float32(16) // "You" / "AI" label height
	gap := float32(10)    // gap between messages

	convMutex.RLock()
	messages := make([]Message, len(conversation))
	copy(messages, conversation)
	convMutex.RUnlock()

	// How many messages fit? Show from the end.
	// Each message uses at least labelH + lineH + gap = 46px.
	// Reserve 40px at bottom for partial/processing.
	available := H - yPos - 40
	startIdx := 0
	if len(messages) > 0 {
		// Count from end until we'd exceed available height
		totalH := float32(0)
		startIdx = len(messages)
		for i := len(messages) - 1; i >= 0; i-- {
			lines := wrapText(messages[i].Text, 54)
			msgH := labelH + float32(len(lines))*lineH + gap
			if totalH+msgH > available {
				break
			}
			totalH += msgH
			startIdx = i
		}
	}

	if len(messages) == 0 {
		// Empty state — waiting for speech
		d2dText(pRenderTarget, "Listening for speech...", fmtBody, brushGreenMuted, rectF(18, yPos, W-18, yPos+24))
	} else {
		for i := startIdx; i < len(messages); i++ {
			msg := messages[i]
			if yPos > H-40 {
				break
			}

			if msg.Role == "user" {
				// "You" label in dim green
				d2dText(pRenderTarget, "You", fmtBadge, brushGreenDim, rectF(18, yPos, 50, yPos+labelH))
				yPos += labelH
				// Transcript text in near-white
				for _, line := range wrapText(msg.Text, 54) {
					if yPos > H-40 {
						break
					}
					d2dText(pRenderTarget, line, fmtBody, brushTextMain, rectF(18, yPos, W-18, yPos+lineH))
					yPos += lineH
				}
			} else {
				// "AI" label in bright green
				d2dText(pRenderTarget, "AI", fmtBadge, brushGreen, rectF(18, yPos, 40, yPos+labelH))
				yPos += labelH
				// AI response text in green
				for _, line := range wrapText(msg.Text, 54) {
					if yPos > H-40 {
						break
					}
					d2dText(pRenderTarget, line, fmtBody, brushGreen, rectF(18, yPos, W-18, yPos+lineH))
					yPos += lineH
				}
			}
			yPos += gap
		}
	}

	// ── Live partial transcript (streaming) ─────────────────────────────────
	partialMutex.RLock()
	partial := partialTranscript
	partialMutex.RUnlock()

	if partial != "" && yPos < H-20 {
		d2dText(pRenderTarget, "You", fmtBadge, brushGreenMuted, rectF(18, yPos, 50, yPos+labelH))
		yPos += labelH
		// Cursor blink using time
		cursor := ""
		if (time.Now().UnixNano()/400000000)%2 == 0 {
			cursor = "▌"
		}
		d2dText(pRenderTarget, partial+cursor, fmtBody, brushTextDim, rectF(18, yPos, W-18, yPos+lineH))
	}

	// ── Processing indicator ─────────────────────────────────────────────────
	if isProcessing {
		n := int(time.Now().Unix() % 4)
		dots := strings.Repeat(".", n)
		d2dText(pRenderTarget, "AI"+dots, fmtBadge, brushGreen, rectF(18, H-22, W-18, H-6))
	}

	comCall(pRenderTarget, 49, 0, 0) // EndDraw
}

// ═══════════════════════════════════════════════════════════════════════════
//  AssemblyAI WebSocket  (UNCHANGED)
// ═══════════════════════════════════════════════════════════════════════════

func connectAssemblyAI() error {
	if assemblyAIKey == "" {
		return fmt.Errorf("ASSEMBLYAI_API_KEY not set")
	}

	params := url.Values{}
	params.Add("sample_rate", fmt.Sprintf("%d", SAMPLE_RATE))
	params.Add("format_turns", "true")
	params.Add("speech_model", "u3-rt-pro")

	wsURL := fmt.Sprintf("%s?%s", ASSEMBLYAI_WS_URL, params.Encode())
	fmt.Printf("🔌 Connecting to AssemblyAI: %s\n", wsURL)

	headers := http.Header{}
	headers.Add("Authorization", assemblyAIKey)

	dialer := websocket.Dialer{HandshakeTimeout: 10 * time.Second}
	conn, resp, err := dialer.Dial(wsURL, headers)
	if err != nil {
		if resp != nil {
			body, _ := io.ReadAll(resp.Body)
			fmt.Printf("❌ AssemblyAI connection failed: %v\n   Response: %s\n", err, string(body))
		} else {
			fmt.Printf("❌ AssemblyAI connection failed: %v\n", err)
		}
		return fmt.Errorf("failed to connect to AssemblyAI: %w", err)
	}
	if resp != nil {
		fmt.Printf("✅ WebSocket handshake: %d\n", resp.StatusCode)
	}

	aaiMutex.Lock()
	aaiWSConn = conn
	aaiConnected = true
	aaiMutex.Unlock()

	fmt.Println("✅ Connected to AssemblyAI Universal-Streaming")
	go handleAssemblyAIResponses()
	return nil
}

func handleAssemblyAIResponses() {
	fmt.Println("👂 AssemblyAI response handler started")
	msgCount := 0

	for {
		aaiMutex.RLock()
		conn := aaiWSConn
		aaiMutex.RUnlock()

		if conn == nil {
			fmt.Println("🛑 AssemblyAI connection closed")
			break
		}

		msgType, data, err := conn.ReadMessage()
		if err != nil {
			fmt.Printf("❌ AssemblyAI read error: %v\n", err)
			aaiMutex.Lock()
			aaiConnected = false
			aaiWSConn = nil
			aaiMutex.Unlock()
			break
		}
		if msgType == websocket.BinaryMessage {
			continue
		}

		msgCount++
		fmt.Printf("📨 AssemblyAI message #%d: %s\n", msgCount, string(data))

		var baseMsg AAIBaseMessage
		if err := json.Unmarshal(data, &baseMsg); err != nil {
			fmt.Printf("⚠️  Failed to parse message: %v\n", err)
			continue
		}

		switch baseMsg.Type {
		case "Begin":
			var msg AAIBeginMessage
			json.Unmarshal(data, &msg)
			fmt.Printf("🟢 AssemblyAI session started: %s\n", msg.ID)
		case "Turn":
			var msg AAITurnMessage
			json.Unmarshal(data, &msg)
			handleTurnMessage(msg)
		case "Termination":
			var msg AAITerminationMessage
			json.Unmarshal(data, &msg)
			fmt.Printf("🔴 Session terminated: %ds audio processed\n", msg.AudioDurationSeconds)
			aaiMutex.Lock()
			aaiConnected = false
			aaiWSConn = nil
			aaiMutex.Unlock()
			return
		default:
			fmt.Printf("❓ Unknown message type: %s\n", baseMsg.Type)
		}
	}
}

func handleTurnMessage(msg AAITurnMessage) {
	fmt.Printf("🎯 Turn: formatted=%v endOfTurn=%v transcript='%s'\n",
		msg.TurnIsFormatted, msg.EndOfTurn, msg.Transcript)

	if strings.TrimSpace(msg.Transcript) == "" {
		return
	}

	if msg.TurnIsFormatted {
		transcriptMutex.Lock()
		currentTranscript = msg.Transcript
		transcriptMutex.Unlock()

		partialMutex.Lock()
		partialTranscript = ""
		partialMutex.Unlock()

		fmt.Printf("📝 FINAL: %s\n", msg.Transcript)

		if !isProcessing && !isAIResponding {
			go sendToAI(msg.Transcript)
		}
	} else if !msg.EndOfTurn {
		partialMutex.Lock()
		partialTranscript = msg.Transcript
		partialMutex.Unlock()
		fmt.Printf("📝 PARTIAL: %s\n", msg.Transcript)
	}

	if hwndMain != 0 {
		procInvalidateRect.Call(uintptr(hwndMain), 0, 1)
	}
}

func sendAudioToAssemblyAI(audioData []byte) error {
	aaiMutex.RLock()
	conn := aaiWSConn
	connected := aaiConnected
	aaiMutex.RUnlock()

	if !connected || conn == nil {
		return fmt.Errorf("AssemblyAI not connected")
	}
	return conn.WriteMessage(websocket.BinaryMessage, audioData)
}

func disconnectAssemblyAI() {
	aaiMutex.Lock()
	defer aaiMutex.Unlock()

	if aaiWSConn != nil {
		terminate := AAITerminateMessage{Type: "Terminate"}
		data, _ := json.Marshal(terminate)
		aaiWSConn.WriteMessage(websocket.TextMessage, data)
		time.Sleep(200 * time.Millisecond)
		aaiWSConn.Close()
		aaiWSConn = nil
	}
	aaiConnected = false
}

// ═══════════════════════════════════════════════════════════════════════════
//  Audio Capture  (UNCHANGED)
// ═══════════════════════════════════════════════════════════════════════════

func initAudioCapture() error {
	hr, _, _ := procCoInitializeEx.Call(0, 0)
	if hr != 0 && hr != 1 {
		return fmt.Errorf("CoInitializeEx failed: 0x%X", hr)
	}

	var pEnumerator uintptr
	hr, _, _ = procCoCreateInstance.Call(
		uintptr(unsafe.Pointer(&CLSID_MMDeviceEnumerator)),
		0, CLSCTX_ALL,
		uintptr(unsafe.Pointer(&IID_IMMDeviceEnumerator)),
		uintptr(unsafe.Pointer(&pEnumerator)),
	)
	if hr != 0 {
		return fmt.Errorf("CoCreateInstance(MMDeviceEnumerator): 0x%X", hr)
	}

	var pDevice uintptr
	hr = comCall(pEnumerator, 4, 0, 1, uintptr(unsafe.Pointer(&pDevice)))
	if hr != 0 {
		return fmt.Errorf("GetDefaultAudioEndpoint: 0x%X", hr)
	}

	hr = comCall(pDevice, 3,
		uintptr(unsafe.Pointer(&IID_IAudioClient)),
		CLSCTX_ALL, 0,
		uintptr(unsafe.Pointer(&pAudioClient)),
	)
	if hr != 0 {
		return fmt.Errorf("Activate(IAudioClient): 0x%X", hr)
	}

	var pFormat uintptr
	hr = comCall(pAudioClient, 8, uintptr(unsafe.Pointer(&pFormat)))
	if hr != 0 {
		return fmt.Errorf("GetMixFormat: 0x%X", hr)
	}
	audioFormat = (*WAVEFORMATEX)(unsafe.Pointer(pFormat))

	fmt.Printf("🎧 Loopback capture: %d Hz, %d channels, %d bits, blockAlign=%d\n",
		audioFormat.NSamplesPerSec, audioFormat.NChannels, audioFormat.WBitsPerSample, audioFormat.NBlockAlign)

	hr = comCall(pAudioClient, 3,
		AUDCLNT_SHAREMODE_SHARED,
		AUDCLNT_STREAMFLAGS_LOOPBACK,
		10000000, 0, pFormat, 0,
	)
	if hr != 0 {
		return fmt.Errorf("IAudioClient::Initialize: 0x%X", hr)
	}

	hr = comCall(pAudioClient, 14,
		uintptr(unsafe.Pointer(&IID_IAudioCaptureClient)),
		uintptr(unsafe.Pointer(&pCaptureClient)),
	)
	if hr != 0 {
		return fmt.Errorf("GetService(IAudioCaptureClient): 0x%X", hr)
	}

	hr = comCall(pAudioClient, 10)
	if hr != 0 {
		return fmt.Errorf("IAudioClient::Start: 0x%X", hr)
	}

	captureRunning = true
	isListening = true

	fmt.Println("✅ Audio capture started")

	if err := connectAssemblyAI(); err != nil {
		fmt.Printf("⚠️  AssemblyAI connection failed: %v\n", err)
	}

	go audioCaptureLoop()
	return nil
}

func audioCaptureLoop() {
	const targetChunkMs = 100
	var sendBuffer []byte
	chunkCount := 0
	lastLog := time.Now()

	fmt.Println("🎤 Audio capture loop started")

	for captureRunning {
		var pData uintptr
		var framesAvailable uint32
		var flags uint32
		var devicePos uint64
		var qpcPos uint64

		hr := comCall(pCaptureClient, 3,
			uintptr(unsafe.Pointer(&pData)),
			uintptr(unsafe.Pointer(&framesAvailable)),
			uintptr(unsafe.Pointer(&flags)),
			uintptr(unsafe.Pointer(&devicePos)),
			uintptr(unsafe.Pointer(&qpcPos)),
		)

		if hr == 0 && framesAvailable > 0 {
			bytesAvailable := uint32(framesAvailable) * uint32(audioFormat.NBlockAlign)
			if flags&AUDCLNT_BUFFERFLAGS_SILENT == 0 {
				data := unsafe.Slice((*byte)(unsafe.Pointer(pData)), bytesAvailable)
				converted := convertAudioFormat(data)
				sendBuffer = append(sendBuffer, converted...)
			}
			comCall(pCaptureClient, 4, uintptr(framesAvailable))
		}

		const outBytesPerMs = SAMPLE_RATE * 2 / 1000
		const targetBytes = targetChunkMs * outBytesPerMs

		if len(sendBuffer) >= targetBytes {
			chunk := make([]byte, targetBytes)
			copy(chunk, sendBuffer[:targetBytes])
			sendBuffer = sendBuffer[targetBytes:]

			if err := sendAudioToAssemblyAI(chunk); err != nil {
				if chunkCount%50 == 0 {
					fmt.Printf("⚠️  Send error (chunk %d): %v\n", chunkCount, err)
				}
			} else {
				if time.Since(lastLog) > 2*time.Second {
					fmt.Printf("📤 Sent chunk %d (%d bytes), buffer=%d\n", chunkCount, len(chunk), len(sendBuffer))
					lastLog = time.Now()
				}
			}
			chunkCount++
		}

		time.Sleep(5 * time.Millisecond)
	}
	fmt.Println("🛑 Audio capture loop ended")
}

func convertAudioFormat(data []byte) []byte {
	if audioFormat == nil {
		return data
	}

	srcRate := int(audioFormat.NSamplesPerSec)
	channels := int(audioFormat.NChannels)
	bitsPerSample := int(audioFormat.WBitsPerSample)

	var monoFloat []float32

	if bitsPerSample == 32 && channels == 2 {
		frameCount := len(data) / 8
		monoFloat = make([]float32, frameCount)
		for i := 0; i < frameCount; i++ {
			lBits := binary.LittleEndian.Uint32(data[i*8:])
			rBits := binary.LittleEndian.Uint32(data[i*8+4:])
			l := math.Float32frombits(lBits)
			r := math.Float32frombits(rBits)
			monoFloat[i] = (l + r) * 0.5
		}
	} else if bitsPerSample == 32 && channels == 1 {
		frameCount := len(data) / 4
		monoFloat = make([]float32, frameCount)
		for i := 0; i < frameCount; i++ {
			bits := binary.LittleEndian.Uint32(data[i*4:])
			monoFloat[i] = math.Float32frombits(bits)
		}
	} else if bitsPerSample == 16 && channels == 2 {
		frameCount := len(data) / 4
		monoFloat = make([]float32, frameCount)
		for i := 0; i < frameCount; i++ {
			l := int16(binary.LittleEndian.Uint16(data[i*4:]))
			r := int16(binary.LittleEndian.Uint16(data[i*4+2:]))
			monoFloat[i] = float32(l+r) / 2.0 / 32768.0
		}
	} else if bitsPerSample == 16 && channels == 1 {
		frameCount := len(data) / 2
		monoFloat = make([]float32, frameCount)
		for i := 0; i < frameCount; i++ {
			s := int16(binary.LittleEndian.Uint16(data[i*2:]))
			monoFloat[i] = float32(s) / 32768.0
		}
	} else {
		fmt.Printf("⚠️  Unsupported format: %d Hz %dch %d-bit\n", srcRate, channels, bitsPerSample)
		return data
	}

	outRate := SAMPLE_RATE
	var decimated []float32

	if srcRate == outRate {
		decimated = monoFloat
	} else {
		step := float64(srcRate) / float64(outRate)
		outLen := int(float64(len(monoFloat)) / step)
		decimated = make([]float32, 0, outLen)

		for i := 0; i < outLen; i++ {
			startF := float64(i) * step
			endF := startF + step
			start := int(startF)
			end := int(endF)
			if end > len(monoFloat) {
				end = len(monoFloat)
			}
			if start >= len(monoFloat) {
				break
			}
			var sum float32
			for j := start; j < end; j++ {
				sum += monoFloat[j]
			}
			decimated = append(decimated, sum/float32(end-start))
		}
	}

	out := make([]byte, len(decimated)*2)
	for i, f := range decimated {
		if f > 1.0 {
			f = 1.0
		} else if f < -1.0 {
			f = -1.0
		}
		s := int16(f * 32767.0)
		binary.LittleEndian.PutUint16(out[i*2:], uint16(s))
	}
	return out
}

func stopAudioCapture() {
	captureRunning = false
	isListening = false
	appStopped = true
	disconnectAssemblyAI()
	if pAudioClient != 0 {
		comCall(pAudioClient, 11)
	}
}

// ═══════════════════════════════════════════════════════════════════════════
//  AI Integration  (UNCHANGED)
// ═══════════════════════════════════════════════════════════════════════════

func sendToAI(userText string) {
	isProcessing = true
	isAIResponding = true

	convMutex.Lock()
	conversation = append(conversation, Message{
		Role: "user", Text: userText, Timestamp: time.Now(),
	})
	convMutex.Unlock()

	go func() {
		response := callAIAPI(userText)

		convMutex.Lock()
		conversation = append(conversation, Message{
			Role: "assistant", Text: response, Timestamp: time.Now(),
		})
		convMutex.Unlock()

		isProcessing = false
		isAIResponding = false

		if hwndMain != 0 {
			procInvalidateRect.Call(uintptr(hwndMain), 0, 1)
		}
	}()
}

func callAIAPI(prompt string) string {
	time.Sleep(2 * time.Second)
	return "Use the STAR method:\n1. Situation: Sprint planning conflict\n2. Task: Align PM & eng timelines\n3. Action: Facilitated async RFC doc\n4. Result: Shipped on time, 0 escalations"
}

// ═══════════════════════════════════════════════════════════════════════════
//  Window procedure
// ═══════════════════════════════════════════════════════════════════════════

func wndProc(hwnd windows.Handle, msg uint32, wParam uintptr, lParam uintptr) uintptr {
	switch msg {
	case WM_PAINT:
		var ps PAINTSTRUCT
		procBeginPaint.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&ps)))
		renderFrame()
		procEndPaint.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&ps)))
		return 0

	case WM_NCHITTEST:
		lx := int32(int16(lParam & 0xFFFF))
		ly := int32(int16((lParam >> 16) & 0xFFFF))
		var wr RECT
		procGetWindowRect.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&wr)))
		cx, cy := lx-wr.Left, ly-wr.Top
		if inCloseBtn(cx, cy) || inStopBtn(cx, cy) {
			return 1 // HTCLIENT — so WM_LBUTTONDOWN fires
		}
		return 2 // HTCAPTION — drag anywhere else

	case WM_LBUTTONDOWN:
		x := int32(lParam & 0xFFFF)
		y := int32((lParam >> 16) & 0xFFFF)
		if inCloseBtn(x, y) {
			stopAudioCapture()
			procPostQuitMessage.Call(0)
		} else if inStopBtn(x, y) && !appStopped {
			stopAudioCapture()
			procInvalidateRect.Call(uintptr(hwnd), 0, 1)
		}
		return 0

	case WM_MOVE:
		windowX = int32(int16(lParam & 0xFFFF))
		windowY = int32(int16((lParam >> 16) & 0xFFFF))
		return 0

	case WM_ACTIVATEAPP:
		enforceTopmostOnly(uintptr(hwnd))
		return 0

	case WM_TIMER:
		if wParam == TIMER_ID {
			enforceTopmostOnly(uintptr(hwnd))

			var pt POINT
			procGetCursorPos.Call(uintptr(unsafe.Pointer(&pt)))
			var wr RECT
			procGetWindowRect.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&wr)))
			cx, cy := pt.X-wr.Left, pt.Y-wr.Top

			newClose := inCloseBtn(cx, cy)
			newStop := inStopBtn(cx, cy)
			if newClose != closeHover || newStop != stopHover {
				closeHover = newClose
				stopHover = newStop
				procInvalidateRect.Call(uintptr(hwnd), 0, 1)
			}
		} else if wParam == TIMER_AI || wParam == TIMER_AUDIO {
			procInvalidateRect.Call(uintptr(hwnd), 0, 1)
		}
		return 0

	case WM_KEYDOWN:
		if wParam == VK_ESCAPE {
			stopAudioCapture()
			procPostQuitMessage.Call(0)
		}
		return 0

	case WM_DESTROY:
		stopAudioCapture()
		procPostQuitMessage.Call(0)
		return 0

	default:
		ret, _, _ := procDefWindowProcW.Call(uintptr(hwnd), uintptr(msg), wParam, lParam)
		return ret
	}
}

func enforceTopmostOnly(hwnd uintptr) {
	procSetWindowPos.Call(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_SHOWWINDOW)
}

func enforceTopmost(hwnd uintptr) {
	procSetWindowPos.Call(hwnd, HWND_TOPMOST, uintptr(windowX), uintptr(windowY), WIN_W, WIN_H, SWP_SHOWWINDOW)
}

func inCloseBtn(x, y int32) bool {
	return x >= CLOSE_X && x <= CLOSE_X+CLOSE_SIZE && y >= CLOSE_Y && y <= CLOSE_Y+CLOSE_SIZE
}

func inStopBtn(x, y int32) bool {
	return x >= STOP_X && x <= STOP_X+STOP_SIZE && y >= STOP_Y && y <= STOP_Y+STOP_SIZE
}

func getSystemMetrics(n int32) int32 {
	ret, _, _ := procGetSystemMetrics.Call(uintptr(n))
	return int32(ret)
}

// ═══════════════════════════════════════════════════════════════════════════
//  Utility  (UNCHANGED)
// ═══════════════════════════════════════════════════════════════════════════

func wrapText(text string, maxLen int) []string {
	var lines []string
	words := splitWords(text)
	current := ""
	for _, word := range words {
		if len(current)+len(word)+1 > maxLen {
			if current != "" {
				lines = append(lines, current)
			}
			current = word
		} else {
			if current != "" {
				current += " "
			}
			current += word
		}
	}
	if current != "" {
		lines = append(lines, current)
	}
	if len(lines) == 0 {
		lines = append(lines, text)
	}
	return lines
}

func splitWords(s string) []string {
	var words []string
	var current []rune
	for _, r := range s {
		if r == ' ' || r == '\t' || r == '\n' {
			if len(current) > 0 {
				words = append(words, string(current))
				current = nil
			}
		} else {
			current = append(current, r)
		}
	}
	if len(current) > 0 {
		words = append(words, string(current))
	}
	return words
}

func truncate(s string, n int) string {
	if len(s) <= n {
		return s
	}
	return s[:n-3] + "..."
}

func formatTime(t time.Time) string { return t.Format("3:04 PM") }

// ═══════════════════════════════════════════════════════════════════════════
//  Main
// ═══════════════════════════════════════════════════════════════════════════

func main() {
	// assemblyAIKey = os.Getenv("ASSEMBLYAI_API_KEY")
	assemblyAIKey = "4c9348b3bfe847a797221bbbaff7cfb6"

	fmt.Println("═══════════════════════════════════════════")
	fmt.Println("  Parakeet AI — Interview Copilot")
	fmt.Println("  AssemblyAI Real-Time Streaming")
	fmt.Println("═══════════════════════════════════════════")

	if assemblyAIKey == "" {
		fmt.Println("\n⚠️  WARNING: ASSEMBLYAI_API_KEY not set!")
		fmt.Println("   Set it: $env:ASSEMBLYAI_API_KEY='your_key'")
		fmt.Println("   Continuing without transcription...\n")
	} else {
		fmt.Println("\n✓ AssemblyAI API key configured")
	}

	modHandle, _, _ := kernel32.NewProc("GetModuleHandleW").Call(0)
	hInstance := windows.Handle(modHandle)

	className, _ := syscall.UTF16PtrFromString("ParakeetOverlayClass")
	wcex := WNDCLASSEXW{
		CbSize:        uint32(unsafe.Sizeof(WNDCLASSEXW{})),
		LpfnWndProc:   syscall.NewCallback(wndProc),
		HInstance:     hInstance,
		HbrBackground: 0,
		LpszClassName: className,
	}
	if ret, _, err := procRegisterClassExW.Call(uintptr(unsafe.Pointer(&wcex))); ret == 0 {
		fmt.Println("RegisterClassEx failed:", err)
		return
	}

	screenW := int32(getSystemMetrics(0))
	windowX = screenW - WIN_W - 24
	windowY = 40

	windowName, _ := syscall.UTF16PtrFromString("Parakeet AI")
	hwnd, _, err := procCreateWindowExW.Call(
		uintptr(WS_EX_LAYERED|WS_EX_TOOLWINDOW|WS_EX_NOACTIVATE|WS_EX_TOPMOST),
		uintptr(unsafe.Pointer(className)),
		uintptr(unsafe.Pointer(windowName)),
		uintptr(WS_POPUP|WS_VISIBLE),
		uintptr(windowX), uintptr(windowY),
		WIN_W, WIN_H,
		0, 0, uintptr(hInstance), 0,
	)
	if hwnd == 0 {
		fmt.Println("CreateWindowEx failed:", err)
		return
	}
	hwndMain = windows.Handle(hwnd)

	rgn, _, _ := procCreateRoundRectRgn.Call(0, 0, WIN_W, WIN_H, 20, 20)
	procSetWindowRgn.Call(hwnd, rgn, 1)

	if err := initD2D(hwnd); err != nil {
		fmt.Println("Direct2D init failed:", err)
		return
	}

	// ← LWA_COLORKEY: pure black (0x000000) pixels become see-through.
	//   No DWM blur — user sees the actual content behind the window unblurred.
	procSetLayeredWindowAttributes.Call(hwnd, 0, 215, LWA_ALPHA)

	procShowWindow.Call(hwnd, SW_SHOW)
	procUpdateWindow.Call(hwnd)
	enforceTopmost(hwnd)

	if ok, _, _ := procSetWindowDisplayAffinity.Call(hwnd, WDA_EXCLUDEFROMCAPTURE); ok != 0 {
		fmt.Println("✅ Screen capture protection active")
	} else {
		fmt.Println("⚠️  Needs Windows 10 2004+ for capture protection")
	}

	if err := initAudioCapture(); err != nil {
		fmt.Println("⚠️  Audio capture init failed:", err)
		fmt.Println("    Continuing without audio...")
		transcriptMutex.Lock()
		currentTranscript = "[Audio capture not available]"
		transcriptMutex.Unlock()
	}

	procSetTimer.Call(hwnd, TIMER_ID, 150, 0)
	procSetTimer.Call(hwnd, TIMER_AI, 500, 0)
	procSetTimer.Call(hwnd, TIMER_AUDIO, 100, 0)

	fmt.Println("\n🎧 Listening to system audio...")
	fmt.Println("   Click ■ to stop  |  Click × or ESC to quit\n")

	var msg MSG
	for {
		if r, _, _ := procGetMessageW.Call(uintptr(unsafe.Pointer(&msg)), 0, 0, 0); r == 0 {
			break
		}
		procTranslateMessage.Call(uintptr(unsafe.Pointer(&msg)))
		procDispatchMessageW.Call(uintptr(unsafe.Pointer(&msg)))
	}
}
