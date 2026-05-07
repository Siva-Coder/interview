package main

import (
	"bufio"
	"bytes"
	"os"

	// "os"
	"context"
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
	LWA_COLORKEY = 0x00000001

	WM_DESTROY     = 0x0002
	WM_PAINT       = 0x000F
	WM_SIZE        = 0x0005
	WM_KEYDOWN     = 0x0100
	WM_LBUTTONDOWN = 0x0201
	WM_TIMER       = 0x0113
	WM_ACTIVATEAPP = 0x001C
	WM_NCHITTEST   = 0x0084
	WM_MOVE        = 0x0003
	WM_MOUSEMOVE   = 0x0200
	WM_LBUTTONUP   = 0x0202
	WM_MOUSEWHEEL  = 0x020A

	VK_ESCAPE = 0x1B

	DWM_BB_ENABLE = 0x00000001

	WIN_W = 460
	WIN_H = 500

	// Close button
	CLOSE_X    int32 = WIN_W - 36
	CLOSE_Y    int32 = 14
	CLOSE_SIZE int32 = 20

	// Hand Cursoe
	IDC_HAND = 32649

	// Stop/resume button (left of close)
	STOP_X    int32 = WIN_W - 64
	STOP_Y    int32 = 14
	STOP_SIZE int32 = 20

	// Transparency slider (left of stop button)
	SLIDER_X       int32 = 190
	SLIDER_Y       int32 = 18
	SLIDER_W       int32 = STOP_X - SLIDER_X - 12
	SLIDER_H       int32 = 14
	SLIDER_THUMB_W int32 = 10

	// Power/start button (centre of window)
	POWER_BTN_X int32 = WIN_W/2 - 50
	POWER_BTN_Y int32 = WIN_H/2 - 30
	POWER_BTN_W int32 = 100
	POWER_BTN_H int32 = 40

	// Timer IDs
	TIMER_REFRESH   = 1
	TIMER_AI_POLL   = 2
	TIMER_AUDIO_CAP = 3

	// Direct2D / DirectWrite
	D2D1_FACTORY_TYPE_SINGLE_THREADED                 = 0
	DWRITE_FACTORY_TYPE_SHARED                        = 0
	D2D1_RENDER_TARGET_TYPE_DEFAULT                   = 0
	D2D1_RENDER_TARGET_USAGE_NONE                     = 0
	D2D1_FEATURE_LEVEL_DEFAULT                        = 0
	D2D1_ALPHA_MODE_PREMULTIPLIED                     = 1
	DWRITE_FONT_WEIGHT_NORMAL                         = 400
	DWRITE_FONT_WEIGHT_SEMI_BOLD                      = 600
	DWRITE_FONT_WEIGHT_BOLD                           = 700
	DWRITE_FONT_STYLE_NORMAL                          = 0
	DWRITE_FONT_STRETCH_NORMAL                        = 5
	DWRITE_TEXT_ALIGNMENT_LEADING                     = 0
	DWRITE_PARAGRAPH_ALIGNMENT_NEAR                   = 0
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
//  AssemblyAI Message Types
// ═══════════════════════════════════════════════════════════════════════════

type openaiChatMessage struct {
	Role    string `json:"role"`
	Content string `json:"content"`
}

type openaiChatCompletionRequest struct {
	Model    string              `json:"model"`
	Messages []openaiChatMessage `json:"messages"`
	Stream   bool                `json:"stream"`
}

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
//  COM GUIDs
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
//  COM vtable structs
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
//  Structs
// ═══════════════════════════════════════════════════════════════════════════

type D2D1_COLOR_F struct{ R, G, B, A float32 }
type D2D1_RECT_F struct{ Left, Top, Right, Bottom float32 }

type D2D1_ROUNDED_RECT struct {
	Rect             D2D1_RECT_F
	RadiusX, RadiusY float32
}

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
//  DLLs & procs
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

	procDwmEnableBlurBehindWindow = dwmapi.NewProc("DwmEnableBlurBehindWindow")
	procD2D1CreateFactory         = d2d1dll.NewProc("D2D1CreateFactory")
	procDWriteCreateFactory       = dwritedll.NewProc("DWriteCreateFactory")
	procCoInitializeEx            = ole32.NewProc("CoInitializeEx")
	procCoCreateInstance          = ole32.NewProc("CoCreateInstance")
	procCoTaskMemFree             = ole32.NewProc("CoTaskMemFree")

	procSetCursor   = user32.NewProc("SetCursor")
	procLoadCursorW = user32.NewProc("LoadCursorW")

	HWND_TOPMOST   = ^uintptr(0)
	SWP_SHOWWINDOW = uintptr(0x0040)
	SWP_NOMOVE     = uintptr(0x0002)
	SWP_NOSIZE     = uintptr(0x0001)
	SW_SHOW        = uintptr(5)

	procSetCapture     = user32.NewProc("SetCapture")
	procReleaseCapture = user32.NewProc("ReleaseCapture")
)

// ═══════════════════════════════════════════════════════════════════════════
//  Color helpers
// ═══════════════════════════════════════════════════════════════════════════

func rgba(r, g, b uint8, a float32) D2D1_COLOR_F {
	return D2D1_COLOR_F{R: float32(r) / 255.0, G: float32(g) / 255.0, B: float32(b) / 255.0, A: a}
}
func rectF(l, t, r, b float32) D2D1_RECT_F { return D2D1_RECT_F{l, t, r, b} }

var (
	clrBg         = rgba(0x0D, 0x0D, 0x18, 1.0)
	clrGreen      = rgba(0x34, 0xd3, 0x99, 1)
	clrGreenDim   = rgba(0x10, 0x99, 0x66, 1)
	clrGreenMuted = rgba(0x0a, 0x5c, 0x40, 1)
	clrTextMain   = rgba(0xf0, 0xf4, 0xff, 1)
	clrTextDim    = rgba(0x9a, 0xa2, 0xb8, 1)
	clrRed        = rgba(0xf8, 0x71, 0x71, 1)
	clrBtnBg      = rgba(0x23, 0x8b, 0x5a, 1) // dark green for power button
	clrRedBright  = rgba(0xff, 0x4d, 0x4d, 1)
	clrUserBubble = rgba(0x1a, 0x2d, 0x42, 0.9) // dark blue
	clrAiBubble   = rgba(0x12, 0x2e, 0x25, 0.9) // dark green
)

// ═══════════════════════════════════════════════════════════════════════════
//  D2D/DWrite state
// ═══════════════════════════════════════════════════════════════════════════

var (
	pD2DFactory    uintptr
	pDWriteFactory uintptr
	pRenderTarget  uintptr

	brushBg, brushGreen, brushGreenDim, brushGreenMuted uintptr
	brushTextMain, brushTextDim, brushRed, brushBtnBg   uintptr
	brushRedBright                                      uintptr
	brushUserBubble                                     uintptr
	brushAiBubble                                       uintptr

	fmtTitle, fmtSub, fmtBody, fmtBadge, fmtSmall uintptr

	hwndMain   windows.Handle
	windowX    int32
	windowY    int32
	closeHover bool
	stopHover  bool
	fmtClose   uintptr
)

// ═══════════════════════════════════════════════════════════════════════════
//  App State
// ═══════════════════════════════════════════════════════════════════════════

type Message struct {
	Role      string
	Text      string
	Lines     []string
	Timestamp time.Time
}

var (
	conversation []Message
	// Streaming AI partial response
	partialAIResponse      string
	partialAIResponseMutex sync.Mutex

	geminiKey         string
	convMutex         sync.RWMutex
	isListening       bool
	isProcessing      bool
	isAIResponding    bool
	currentTranscript string
	transcriptMutex   sync.RWMutex
	partialTranscript string
	partialMutex      sync.RWMutex

	// Session control
	sessionActive bool
	sessionMutex  sync.Mutex

	// AI cancellation
	cancelAI    context.CancelFunc
	aiCancelMux sync.Mutex

	// Audio capture (timer-driven on main thread)
	pAudioClient   uintptr
	pCaptureClient uintptr
	audioFormat    *WAVEFORMATEX
	audioStarted   bool
	audioBuffer    []byte
	audioMutex     sync.Mutex

	// AssemblyAI WebSocket
	aaiWSConn      *websocket.Conn
	aaiConnected   bool
	aaiMutex       sync.RWMutex
	assemblyAIKey  string
	scrollOffset   float32
	scrollMutex    sync.Mutex
	totalMsgHeight float32

	//scroll
	scrollDragging  bool
	dragStartY      int32
	dragStartOffset float32
	alphaValue      uint8 = 215 // current transparency (215 = slightly transparent)
	dragSlider      bool
	sliderThumbPos  int32 // not used, but keep if needed

	pendingSendCancel    context.CancelFunc
	pendingSendCancelMux sync.Mutex
)

// ═══════════════════════════════════════════════════════════════════════════
//  COM helpers
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
//  Direct2D helpers
// ═══════════════════════════════════════════════════════════════════════════

func createBrush(rt uintptr, c D2D1_COLOR_F) uintptr {
	var brush uintptr
	comCall(rt, 8, uintptr(unsafe.Pointer(&c)), 0, uintptr(unsafe.Pointer(&brush)))
	return brush
}

func d2dFillRect(rt, brush uintptr, rc D2D1_RECT_F) {
	comCall(rt, 17, uintptr(unsafe.Pointer(&rc)), brush)
}

func d2dFillRoundedRect(rt, brush uintptr, rc D2D1_RECT_F, rx, ry float32) {
	rr := D2D1_ROUNDED_RECT{Rect: rc, RadiusX: rx, RadiusY: ry}
	comCall(rt, 19, uintptr(unsafe.Pointer(&rr)), brush)
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
		uintptr(mathFloat32bits(size)),
		uintptr(unsafe.Pointer(localeW)),
		uintptr(unsafe.Pointer(&pFmt)),
	)
	return pFmt
}

func mathFloat32bits(f float32) uint32 { return *(*uint32)(unsafe.Pointer(&f)) }

// re-calculate total pixel height of all conversation messages
func updateTotalMessageHeight() {
	convMutex.RLock()
	defer convMutex.RUnlock()

	totalMsgHeight = 0
	labelH := float32(16)
	lineH := float32(20)
	gap := float32(10)

	for _, msg := range conversation {
		totalMsgHeight += labelH + float32(len(msg.Lines))*lineH + gap
	}
}

// clamp scrollOffset to valid range
func clampScroll() {
	available := float32(WIN_H) - 60 - 60
	maxScroll := totalMsgHeight - available
	if maxScroll < 0 {
		maxScroll = 0
	}
	if scrollOffset > maxScroll {
		scrollOffset = maxScroll
	}
	if scrollOffset < 0 {
		scrollOffset = 0
	}
}

func inChatArea(x, y int32) bool {
	// chat area: from y=60 to y=WIN_H-60, full width
	return y >= 60 && y <= WIN_H-60
}

func inSliderArea(x, y int32) bool {
	return x >= SLIDER_X && x <= SLIDER_X+SLIDER_W && y >= SLIDER_Y && y <= SLIDER_Y+SLIDER_H
}

func updateAlphaFromMouse(mx int32) {
	thumbRange := float32(SLIDER_W - SLIDER_THUMB_W)
	newAlpha := uint8((float32(mx-SLIDER_X) / thumbRange) * 255)
	if newAlpha < 50 {
		newAlpha = 50 // minimum opacity
	}
	alphaValue = newAlpha
	procSetLayeredWindowAttributes.Call(uintptr(hwndMain), 0, uintptr(alphaValue), LWA_ALPHA)
}

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

	comCall(pRenderTarget, 34, 1) // Greyscale for transparency

	brushBg = createBrush(pRenderTarget, clrBg)
	brushGreen = createBrush(pRenderTarget, clrGreen)
	brushGreenDim = createBrush(pRenderTarget, clrGreenDim)
	brushGreenMuted = createBrush(pRenderTarget, clrGreenMuted)
	brushTextMain = createBrush(pRenderTarget, clrTextMain)
	brushTextDim = createBrush(pRenderTarget, clrTextDim)
	brushRed = createBrush(pRenderTarget, clrRed)
	brushBtnBg = createBrush(pRenderTarget, clrBtnBg)
	brushRedBright = createBrush(pRenderTarget, clrRedBright)
	brushUserBubble = createBrush(pRenderTarget, clrUserBubble)
	brushAiBubble = createBrush(pRenderTarget, clrAiBubble)

	fmtTitle = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 17)
	fmtSub = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 12)
	fmtBody = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 13)
	fmtBadge = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_SEMI_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 10)
	fmtSmall = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 10)
	fmtClose = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 22)

	return nil
}

// ═══════════════════════════════════════════════════════════════════════════
//  RENDERING
// ═══════════════════════════════════════════════════════════════════════════

func renderFrame() {
	if pRenderTarget == 0 {
		return
	}
	W := float32(WIN_W)

	comCall(pRenderTarget, 48)                                  // BeginDraw
	comCall(pRenderTarget, 47, uintptr(unsafe.Pointer(&clrBg))) // Clear

	sessionMutex.Lock()
	active := sessionActive
	sessionMutex.Unlock()

	// ── Header ──────────────────────────────────────────────────────────
	d2dText(pRenderTarget, "Candy", fmtTitle, brushGreen, rectF(18, 12, 70, 42))
	d2dText(pRenderTarget, " Ai", fmtTitle, brushGreenDim, rectF(74, 12, 160, 42))

	// Live status dot
	aaiMutex.RLock()
	aaiConn := aaiConnected
	aaiMutex.RUnlock()
	dotBrush := brushGreenMuted
	if isListening && aaiConn && active {
		dotBrush = brushGreen
	}
	d2dFillEllipse(pRenderTarget, dotBrush, 110, 27, 4, 4)

	// Stop / resume button (RED when active)
	var stopIcon string
	if active {
		stopIcon = "■"
	} else {
		stopIcon = "▶"
	}
	stopBr := brushGreenDim
	if active {
		stopBr = brushRed
		if stopHover {
			stopBr = brushRedBright
		}
	} else {
		if stopHover {
			stopBr = brushGreen
		}
	}
	d2dText(pRenderTarget, stopIcon, fmtTitle, stopBr,
		rectF(float32(STOP_X)-1, float32(STOP_Y), float32(STOP_X+STOP_SIZE)+2, float32(STOP_Y+STOP_SIZE)+2))

	// Close button
	closeBrush := brushTextDim
	if closeHover {
		closeBrush = brushRed
	}
	d2dText(pRenderTarget, "×", fmtClose, closeBrush,
		rectF(float32(CLOSE_X), float32(CLOSE_Y)-3, float32(CLOSE_X+CLOSE_SIZE)+6, float32(CLOSE_Y+CLOSE_SIZE)+4))

	// Transparency slider
	s := float32(SLIDER_X)
	sW := float32(SLIDER_W)
	sY := float32(SLIDER_Y) + 2
	// track
	d2dFillRect(pRenderTarget, brushGreenMuted, rectF(s, sY, s+sW, sY+10))
	// thumb
	thumbW := float32(SLIDER_THUMB_W)
	thumbRange := sW - thumbW
	thumbX := s + float32(alphaValue)/255.0*thumbRange
	thBr := brushGreenDim
	if dragSlider {
		thBr = brushGreen
	}
	d2dFillRoundedRect(pRenderTarget, thBr, rectF(thumbX, sY-2, thumbX+thumbW, sY+14), 4, 4)

	// Divider
	d2dFillRect(pRenderTarget, brushGreenDim, rectF(18, 48, W-18, 49))

	// ── Chat messages (scrollable) ───────────────────────────────────────
	const (
		viewTop       = float32(60)
		bottomReserve = float32(60)
	)
	viewBottom := float32(WIN_H) - bottomReserve

	lineH := float32(20)
	labelH := float32(16)
	gap := float32(10)

	convMutex.RLock()
	messages := make([]Message, len(conversation))
	copy(messages, conversation)
	convMutex.RUnlock()

	scrollMutex.Lock()
	offset := scrollOffset
	scrollMutex.Unlock()

	// Pre‑compute heights from the local slice
	msgHeights := make([]float32, len(messages))
	var localTotalH float32
	for i, msg := range messages {
		h := labelH + float32(len(msg.Lines))*lineH + gap
		msgHeights[i] = h
		localTotalH += h
	}

	availableH := viewBottom - viewTop
	contentY0 := viewTop
	if localTotalH > availableH {
		contentY0 = viewBottom - localTotalH - offset
		if contentY0 > viewTop {
			contentY0 = viewTop
		}
	}

	firstY := viewBottom - localTotalH - offset
	// If content is shorter than view, pin to top
	if localTotalH < viewBottom-viewTop {
		firstY = viewTop
	}

	yPos := firstY
	for i := 0; i < len(messages); i++ {
		msg := messages[i]
		msgH := msgHeights[i]
		// skip if completely above visible area
		if yPos+msgH < viewTop {
			yPos += msgH
			continue
		}
		// stop if below visible area
		if yPos > viewBottom {
			break
		}

		if msg.Role == "user" {
			// user bubble – right aligned
			bubbleW := float32(280)
			bubbleX := float32(WIN_W) - bubbleW - 20
			bubbleY := yPos
			d2dFillRoundedRect(pRenderTarget, brushUserBubble, rectF(bubbleX, bubbleY, bubbleX+bubbleW, bubbleY+msgH), 8, 8)
			d2dText(pRenderTarget, "You", fmtBadge, brushGreenDim, rectF(bubbleX+8, bubbleY+2, bubbleX+50, bubbleY+labelH))
			lineY := bubbleY + labelH
			for _, line := range msg.Lines {
				if lineY > viewBottom {
					break
				}
				d2dText(pRenderTarget, line, fmtBody, brushTextMain, rectF(bubbleX+8, lineY, bubbleX+bubbleW-8, lineY+lineH))
				lineY += lineH
			}
		} else {
			// AI message – full width, no bubble
			d2dText(pRenderTarget, "AI", fmtBadge, brushGreen, rectF(18, yPos, 40, yPos+labelH))
			lineY := yPos + labelH
			for _, line := range msg.Lines {
				if lineY > viewBottom {
					break
				}
				d2dText(pRenderTarget, line, fmtBody, brushTextMain, rectF(18, lineY, float32(WIN_W)-18, lineY+lineH))
				lineY += lineH
			}
		}
		yPos += msgH
	}

	// ── Scrollbar ───────────────────────────────────────────────────────
	if localTotalH > availableH {
		frac := availableH / localTotalH
		barH := frac * (viewBottom - viewTop)
		if barH < 20 {
			barH = 20
		}
		maxScroll := localTotalH - availableH
		scrollFrac := float32(0)
		if maxScroll > 0 {
			scrollFrac = offset / maxScroll
			if scrollFrac > 1 {
				scrollFrac = 1
			}
		}
		barY := viewTop + scrollFrac*((viewBottom-viewTop)-barH)
		d2dFillRect(pRenderTarget, brushGreenDim, rectF(W-6, barY, W-4, barY+barH))
	}

	// ── Live partial transcript (user speaking) ─────────────────────────
	partialMutex.RLock()
	partial := partialTranscript
	partialMutex.RUnlock()

	const partialY = float32(WIN_H) - 58
	if partial != "" {
		d2dText(pRenderTarget, "You", fmtBadge, brushGreenMuted, rectF(18, partialY, 50, partialY+labelH))
		cursor := ""
		if (time.Now().UnixNano()/400000000)%2 == 0 {
			cursor = "▌"
		}
		d2dText(pRenderTarget, partial+cursor, fmtBody, brushTextDim,
			rectF(50, partialY, W-18, partialY+lineH))
	}

	// ── AI streaming / processing indicator ─────────────────────────────
	partialAIResponseMutex.Lock()
	aiPartial := partialAIResponse
	partialAIResponseMutex.Unlock()

	const aiY = float32(WIN_H) - 30
	if isAIResponding && aiPartial != "" {
		d2dText(pRenderTarget, "AI", fmtBadge, brushGreen, rectF(18, aiY, 40, aiY+labelH))
		cursor := ""
		if (time.Now().UnixNano()/400000000)%2 == 0 {
			cursor = "▌"
		}
		d2dText(pRenderTarget, aiPartial+cursor, fmtBody, brushTextMain, rectF(40, aiY, W-18, aiY+lineH))
	} else if isProcessing && !isAIResponding {
		n := int(time.Now().Unix() % 4)
		dots := strings.Repeat(".", n)
		d2dText(pRenderTarget, "AI"+dots, fmtBadge, brushGreen, rectF(18, aiY, W-18, aiY+lineH))
	}

	comCall(pRenderTarget, 49, 0, 0) // EndDraw
}

// ═══════════════════════════════════════════════════════════════════════════
//  AssemblyAI WebSocket
// ═══════════════════════════════════════════════════════════════════════════

func connectAssemblyAI() error {
	if assemblyAIKey == "" {
		return fmt.Errorf("ASSEMBLYAI_API_KEY not set")
	}
	aaiMutex.RLock()
	if aaiConnected {
		aaiMutex.RUnlock()
		return nil
	}
	aaiMutex.RUnlock()

	params := url.Values{}
	params.Add("sample_rate", fmt.Sprintf("%d", SAMPLE_RATE))
	params.Add("format_turns", "true")
	params.Add("speech_model", "u3-rt-pro")
	params.Add("min_turn_silence", "600")  // Wait 600ms of silence before checking for end-of-turn
	params.Add("max_turn_silence", "2000") // Force end-of-turn after 2 seconds of silence

	wsURL := fmt.Sprintf("%s?%s", ASSEMBLYAI_WS_URL, params.Encode())
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
		return err
	}

	aaiMutex.Lock()
	aaiWSConn = conn
	aaiConnected = true
	aaiMutex.Unlock()

	fmt.Println("✅ Connected to AssemblyAI")
	go handleAssemblyAIResponses()
	return nil
}

func handleAssemblyAIResponses() {
	fmt.Println("👂 AssemblyAI response handler started")
	for {
		aaiMutex.RLock()
		conn := aaiWSConn
		aaiMutex.RUnlock()
		if conn == nil {
			break
		}

		// 30‑second deadline – only fired if the connection is genuinely dead.
		if err := conn.SetReadDeadline(time.Now().Add(30 * time.Second)); err != nil {
			fmt.Printf("❌ AssemblyAI set deadline error: %v\n", err)
			break
		}

		msgType, data, err := conn.ReadMessage()
		if err != nil {
			// Any error (timeout, close, network) – treat as disconnection.
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

		var baseMsg AAIBaseMessage
		if err := json.Unmarshal(data, &baseMsg); err != nil {
			continue
		}
		switch baseMsg.Type {
		case "Begin":
			var msg AAIBeginMessage
			json.Unmarshal(data, &msg)
			fmt.Printf("🟢 Session started: %s\n", msg.ID)
		case "Turn":
			var msg AAITurnMessage
			json.Unmarshal(data, &msg)
			handleTurnMessage(msg)
		case "Termination":
			var msg AAITerminationMessage
			json.Unmarshal(data, &msg)
			fmt.Printf("🔴 Session terminated\n")
			aaiMutex.Lock()
			aaiConnected = false
			aaiWSConn = nil
			aaiMutex.Unlock()
			return
		}
	}
}

func handleTurnMessage(msg AAITurnMessage) {
	if strings.TrimSpace(msg.Transcript) == "" {
		return
	}

	sessionMutex.Lock()
	active := sessionActive
	sessionMutex.Unlock()
	if !active {
		return
	}

	if msg.TurnIsFormatted {
		// ── final transcript received ─────────────────────────
		// store the final transcript
		transcriptMutex.Lock()
		currentTranscript = msg.Transcript
		transcriptMutex.Unlock()

		// clear partial
		partialMutex.Lock()
		partialTranscript = ""
		partialMutex.Unlock()

		// Interrupt any ongoing AI processing
		aiCancelMux.Lock()
		if cancelAI != nil {
			cancelAI()
			cancelAI = nil
		}
		aiCancelMux.Unlock()

		isProcessing = false
		isAIResponding = false

		// Clear any streaming partial
		partialAIResponseMutex.Lock()
		partialAIResponse = ""
		partialAIResponseMutex.Unlock()

		// ── debounced send (300ms after final) ─────────────────
		// capture the transcript locally so it never changes while we wait
		finalTranscript := msg.Transcript

		// Cancel any previously scheduled send
		pendingSendCancelMux.Lock()
		if pendingSendCancel != nil {
			pendingSendCancel()
		}
		ctx, cancel := context.WithCancel(context.Background())
		pendingSendCancel = cancel
		pendingSendCancelMux.Unlock()

		go func(transcript string) {
			select {
			case <-time.After(300 * time.Millisecond):
				// delay passed → send to AI
				if transcript != "" {
					sendToAI(transcript)
				}
			case <-ctx.Done():
				// cancelled because new speech arrived
			}
		}(finalTranscript)

	} else if !msg.EndOfTurn {
		// ── partial transcript (speaking) ─────────────────────
		// Cancel any pending final‑turn processor (speaker hasn't finished yet)
		pendingSendCancelMux.Lock()
		if pendingSendCancel != nil {
			pendingSendCancel()
			pendingSendCancel = nil
		}
		pendingSendCancelMux.Unlock()

		partialMutex.Lock()
		partialTranscript = msg.Transcript
		partialMutex.Unlock()
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
		return fmt.Errorf("not connected")
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
//  Audio Capture (timer-driven on main thread)
// ═══════════════════════════════════════════════════════════════════════════

func initAudioDevice() error {
	// COM must already be initialised on this thread (main).
	var pEnumerator uintptr
	hr, _, _ := procCoCreateInstance.Call(
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

	fmt.Printf("🎧 Loopback capture: %d Hz, %d channels, %d bits\n",
		audioFormat.NSamplesPerSec, audioFormat.NChannels, audioFormat.WBitsPerSample)

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

	hr = comCall(pAudioClient, 10) // Start
	if hr != 0 {
		return fmt.Errorf("IAudioClient::Start: 0x%X", hr)
	}

	audioStarted = true
	fmt.Println("✅ Audio capture started")
	return nil
}

func audioTick() {
	if !audioStarted || pCaptureClient == 0 {
		return
	}
	sessionMutex.Lock()
	active := sessionActive
	sessionMutex.Unlock()
	if !active {
		return
	}

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
			audioMutex.Lock()
			audioBuffer = append(audioBuffer, converted...)
			audioMutex.Unlock()
		}
		comCall(pCaptureClient, 4, uintptr(framesAvailable))
	}

	const outBytesPerMs = SAMPLE_RATE * 2 / 1000
	const targetBytes = 100 * outBytesPerMs // 100ms chunks

	audioMutex.Lock()
	bufLen := len(audioBuffer)
	audioMutex.Unlock()

	if bufLen >= targetBytes {
		audioMutex.Lock()
		chunk := make([]byte, targetBytes)
		copy(chunk, audioBuffer[:targetBytes])
		audioBuffer = audioBuffer[targetBytes:]
		audioMutex.Unlock()

		if err := sendAudioToAssemblyAI(chunk); err != nil {
			// silently ignore
		}
	}
}

func stopAudioDevice() {
	if audioStarted {
		if pAudioClient != 0 {
			comCall(pAudioClient, 11) // Stop
		}
		audioStarted = false
		procSetTimer.Call(uintptr(hwndMain), TIMER_AUDIO_CAP, 20, 0) // kill timer
	}
}

// ═══════════════════════════════════════════════════════════════════════════
//  Session management
// ═══════════════════════════════════════════════════════════════════════════

func startSession() {
	sessionMutex.Lock()
	if sessionActive {
		sessionMutex.Unlock()
		return
	}
	sessionActive = true
	sessionMutex.Unlock()

	// Initialise audio if not already
	if !audioStarted {
		if err := connectAssemblyAI(); err != nil {
			fmt.Printf("⚠️  AssemblyAI connection failed: %v\n", err)
		}
		if err := initAudioDevice(); err != nil {
			fmt.Printf("⚠️  Audio init failed: %v\n", err)
			sessionMutex.Lock()
			sessionActive = false
			sessionMutex.Unlock()
			return
		}
		// Set audio capture timer (10ms)
		procSetTimer.Call(uintptr(hwndMain), TIMER_AUDIO_CAP, 20, 0)
	} else {
		// Audio already set up, just restart client if stopped
		comCall(pAudioClient, 10) // Start
		procSetTimer.Call(uintptr(hwndMain), TIMER_AUDIO_CAP, 20, 0)
	}

	isListening = true
	fmt.Println("▶ Session started")
	if hwndMain != 0 {
		procInvalidateRect.Call(uintptr(hwndMain), 0, 1)
	}
}

func stopSession() {
	sessionMutex.Lock()
	if !sessionActive {
		sessionMutex.Unlock()
		return
	}
	sessionActive = false
	sessionMutex.Unlock()

	// Cancel any ongoing AI
	aiCancelMux.Lock()
	if cancelAI != nil {
		cancelAI()
		cancelAI = nil
	}
	aiCancelMux.Unlock()
	isProcessing = false
	isAIResponding = false
	isListening = false

	stopAudioDevice()
	disconnectAssemblyAI()

	fmt.Println("■ Session stopped")
	if hwndMain != 0 {
		procInvalidateRect.Call(uintptr(hwndMain), 0, 1)
	}
}

// ═══════════════════════════════════════════════════════════════════════════
//  AI Integration (with cancellation)
// ═══════════════════════════════════════════════════════════════════════════

func streamAIResponse(ctx context.Context, userText string) {

	procSetTimer.Call(uintptr(hwndMain), TIMER_AI_POLL, 50, 0) // 50ms refresh while streaming

	defer func() {
		procSetTimer.Call(uintptr(hwndMain), TIMER_AI_POLL, 500, 0) // back to normal
		// Cleanup when function exits (finished or cancelled)
		partialAIResponseMutex.Lock()
		partialAIResponse = ""
		partialAIResponseMutex.Unlock()
		isProcessing = false
		isAIResponding = false
		if hwndMain != 0 {
			procInvalidateRect.Call(uintptr(hwndMain), 0, 1)
		}
	}()

	// Build conversation array from global history
	convMutex.RLock()
	msgs := make([]openaiChatMessage, 0, len(conversation))
	for _, m := range conversation {
		role := "user"
		if m.Role == "assistant" {
			role = "assistant"
		}
		msgs = append(msgs, openaiChatMessage{Role: role, Content: m.Text})
	}
	convMutex.RUnlock()

	reqBody := openaiChatCompletionRequest{
		Model:    "gemini-2.5-flash-lite",
		Messages: msgs,
		Stream:   true,
	}
	jsonBody, err := json.Marshal(reqBody)
	if err != nil {
		fmt.Printf("❌ marshal request: %v\n", err)
		return
	}

	req, err := http.NewRequestWithContext(ctx, "POST",
		"https://generativelanguage.googleapis.com/v1beta/openai/chat/completions", bytes.NewReader(jsonBody))
	if err != nil {
		fmt.Printf("❌ create request: %v\n", err)
		return
	}
	req.Header.Set("Authorization", "Bearer "+geminiKey)
	req.Header.Set("Content-Type", "application/json")

	client := &http.Client{Timeout: 0} // no timeout, controlled by context
	resp, err := client.Do(req)
	if err != nil {
		if ctx.Err() != nil {
			fmt.Println("⏹️ AI stream cancelled (new speech started)")
		} else {
			fmt.Printf("❌ Gemini request failed: %v\n", err)
		}
		return
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(resp.Body)
		fmt.Printf("❌ Gemini API error %d: %s\n", resp.StatusCode, string(body))
		return
	}

	done := make(chan struct{})
	var fullResponse strings.Builder

	go func() {
		defer close(done)
		scanner := bufio.NewScanner(resp.Body)
		for scanner.Scan() {
			line := scanner.Text()
			if !strings.HasPrefix(line, "data: ") {
				continue
			}
			data := strings.TrimPrefix(line, "data: ")
			if data == "[DONE]" {
				break
			}

			// Parse SSE delta
			var chunk struct {
				Choices []struct {
					Delta struct {
						Content string `json:"content"`
					} `json:"delta"`
				} `json:"choices"`
			}
			if err := json.Unmarshal([]byte(data), &chunk); err != nil {
				continue
			}
			if len(chunk.Choices) > 0 {
				token := chunk.Choices[0].Delta.Content
				fullResponse.WriteString(token)

				// Update visible partial (UI will pick it up on next paint timer)
				partialAIResponseMutex.Lock()
				partialAIResponse = fullResponse.String()
				partialAIResponseMutex.Unlock()
			}

			// Stop early if context cancelled (new speech detected)
			if ctx.Err() != nil {
				return
			}
		}
		if scanner.Err() != nil {
			if ctx.Err() == nil {
				fmt.Printf("❌ streaming error: %v\n", scanner.Err())
			}
			return
		}
	}()

	select {
	case <-ctx.Done():
		resp.Body.Close() // force close to unblock scanner
	case <-done:
	}

	answer := fullResponse.String()
	if answer == "" {
		return
	}

	// Append assistant message to conversation
	convMutex.Lock()
	conversation = append(conversation, Message{
		Role: "assistant", Text: answer, Lines: wrapText(answer, 54), Timestamp: time.Now(),
	})
	convMutex.Unlock()

	updateTotalMessageHeight()
	scrollMutex.Lock()
	scrollOffset = 0
	scrollMutex.Unlock()
}

func sendToAI(userText string) {
	isProcessing = true
	isAIResponding = true

	convMutex.Lock()
	conversation = append(conversation, Message{
		Role: "user", Text: userText, Lines: wrapText(userText, 54), Timestamp: time.Now(),
	})
	convMutex.Unlock()

	updateTotalMessageHeight()
	scrollMutex.Lock()
	scrollOffset = 0 // jump to bottom
	scrollMutex.Unlock()

	ctx, cancel := context.WithCancel(context.Background())
	aiCancelMux.Lock()
	if cancelAI != nil {
		cancelAI() // kill any previous stream
	}
	cancelAI = cancel
	aiCancelMux.Unlock()

	// Clear any previous partial response
	partialAIResponseMutex.Lock()
	partialAIResponse = ""
	partialAIResponseMutex.Unlock()

	go streamAIResponse(ctx, userText)
}

func callAIAPI(ctx context.Context, prompt string) string {
	select {
	case <-ctx.Done():
		return ""
	case <-time.After(2 * time.Second):
	}
	return "Use the STAR method:\n1. Situation: Sprint planning conflict\n2. Task: Align PM & eng timelines\n3. Action: Facilitated async RFC doc\n4. Result: Shipped on time, 0 escalations"
}

// ═══════════════════════════════════════════════════════════════════════════
//  Audio format conversion (unchanged)
// ═══════════════════════════════════════════════════════════════════════════

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
		if inCloseBtn(cx, cy) || inStopBtn(cx, cy) || inPowerBtn(cx, cy) || inSliderArea(cx, cy) {
			return 1 // HTCLIENT – allow button clicks
		}
		if inChatArea(cx, cy) {
			return 1 // HTCLIENT – enable mouse scroll
		}
		return 2 // HTCAPTION – drag window everywhere else

	case WM_MOVE:
		windowX = int32(int16(lParam & 0xFFFF))
		windowY = int32(int16((lParam >> 16) & 0xFFFF))
		return 0

	case WM_ACTIVATEAPP:
		enforceTopmostOnly(uintptr(hwnd))
		return 0

	case WM_TIMER:
		switch wParam {
		case TIMER_REFRESH:
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
		case TIMER_AI_POLL:
			procInvalidateRect.Call(uintptr(hwnd), 0, 1)
		case TIMER_AUDIO_CAP:
			audioTick()
		}
		return 0

	case WM_DESTROY:
		stopSession()
		procPostQuitMessage.Call(0)
		return 0

	case WM_KEYDOWN:
		if wParam == VK_ESCAPE {
			stopSession()
			procPostQuitMessage.Call(0)
			return 0
		}
		if wParam == 0x26 { // VK_UP
			scrollMutex.Lock()
			scrollOffset += 40 // scroll up 40px
			clampScroll()
			scrollMutex.Unlock()
			procInvalidateRect.Call(uintptr(hwnd), 0, 1)
			return 0
		}
		if wParam == 0x28 { // VK_DOWN
			scrollMutex.Lock()
			scrollOffset -= 40
			if scrollOffset < 0 {
				scrollOffset = 0
			}
			scrollMutex.Unlock()
			procInvalidateRect.Call(uintptr(hwnd), 0, 1)
			return 0
		}
		return 0
	case WM_LBUTTONDOWN:
		x := int32(lParam & 0xFFFF)
		y := int32((lParam >> 16) & 0xFFFF)

		if inCloseBtn(x, y) {
			stopSession()
			procPostQuitMessage.Call(0)
			return 0
		}
		if inStopBtn(x, y) {
			sessionMutex.Lock()
			active := sessionActive
			sessionMutex.Unlock()
			if active {
				stopSession()
			} else {
				startSession()
			}
			procInvalidateRect.Call(uintptr(hwnd), 0, 1)
			return 0
		}
		if inPowerBtn(x, y) {
			sessionMutex.Lock()
			active := sessionActive
			sessionMutex.Unlock()
			if !active {
				startSession()
				procInvalidateRect.Call(uintptr(hwnd), 0, 1)
			}
			return 0
		}
		// Start slider drag
		if inSliderArea(x, y) {
			dragSlider = true
			procSetCapture.Call(uintptr(hwnd))
			updateAlphaFromMouse(x)
			return 0
		}
		// Start chat scroll drag
		if inChatArea(x, y) {
			scrollDragging = true
			dragStartY = y
			scrollMutex.Lock()
			dragStartOffset = scrollOffset
			scrollMutex.Unlock()
			procSetCapture.Call(uintptr(hwnd))
			return 0
		}
		return 0

	case WM_MOUSEMOVE:
		x := int32(lParam & 0xFFFF)
		y := int32((lParam >> 16) & 0xFFFF)
		if inSliderArea(x, y) && !dragSlider {
			// show hand cursor
			hCur, _, _ := procLoadCursorW.Call(0, uintptr(IDC_HAND))
			procSetCursor.Call(hCur)
		} else if !scrollDragging {
			// default arrow (IDC_ARROW = 32512)
			hCur, _, _ := procLoadCursorW.Call(0, 32512)
			procSetCursor.Call(hCur)
		}
		if dragSlider {
			updateAlphaFromMouse(int32(lParam & 0xFFFF))
		}
		if scrollDragging {
			y := int32((lParam >> 16) & 0xFFFF)
			delta := dragStartY - y
			scrollMutex.Lock()
			scrollOffset = dragStartOffset + float32(delta)
			clampScroll()
			scrollMutex.Unlock()
			procInvalidateRect.Call(uintptr(hwnd), 0, 1)
		}
		return 0

	case WM_LBUTTONUP:
		if dragSlider {
			dragSlider = false
			procReleaseCapture.Call()
		}
		if scrollDragging {
			scrollDragging = false
			procReleaseCapture.Call()
		}
		return 0

	case WM_MOUSEWHEEL:
		delta := int16(int32(wParam) >> 16)
		scrollMutex.Lock()
		scrollOffset -= float32(delta) / 120 * 40
		clampScroll()
		scrollMutex.Unlock()
		procInvalidateRect.Call(uintptr(hwnd), 0, 1)
		return 0
	}
	ret, _, _ := procDefWindowProcW.Call(uintptr(hwnd), uintptr(msg), wParam, lParam)
	return ret
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
func inPowerBtn(x, y int32) bool {
	return x >= POWER_BTN_X && x <= POWER_BTN_X+POWER_BTN_W &&
		y >= POWER_BTN_Y && y <= POWER_BTN_Y+POWER_BTN_H
}

func getSystemMetrics(n int32) int32 {
	ret, _, _ := procGetSystemMetrics.Call(uintptr(n))
	return int32(ret)
}

// ═══════════════════════════════════════════════════════════════════════════
//  Utility
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

// ═══════════════════════════════════════════════════════════════════════════
//  Main
// ═══════════════════════════════════════════════════════════════════════════

func main() {
	assemblyAIKey = os.Getenv("ASSEMBLYAI_API_KEY")
	geminiKey = os.Getenv("GEMINI_API_KEY")

	fmt.Println("═══════════════════════════════════════════")
	fmt.Println("  Candy AI — Interview Copilot (v2)")
	fmt.Println("═══════════════════════════════════════════")

	// COM init on main thread (required for audio & D2D)
	procCoInitializeEx.Call(0, 0)

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

	windowName, _ := syscall.UTF16PtrFromString("Candy AI")
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

	procSetLayeredWindowAttributes.Call(hwnd, 0, 215, LWA_ALPHA)

	procShowWindow.Call(hwnd, SW_SHOW)
	procUpdateWindow.Call(hwnd)
	enforceTopmost(hwnd)

	if ok, _, _ := procSetWindowDisplayAffinity.Call(hwnd, WDA_EXCLUDEFROMCAPTURE); ok != 0 {
		fmt.Println("✅ Screen capture protection active")
	} else {
		fmt.Println("⚠️  Needs Windows 10 2004+ for capture protection")
	}

	// UI refresh timers
	procSetTimer.Call(hwnd, TIMER_REFRESH, 150, 0)
	procSetTimer.Call(hwnd, TIMER_AI_POLL, 500, 0)

	fmt.Println("\nClick ▶ Start to begin, or use ■/▶ in top-right")
	fmt.Println("   × or ESC to quit\n")

	var msg MSG
	for {
		if r, _, _ := procGetMessageW.Call(uintptr(unsafe.Pointer(&msg)), 0, 0, 0); r == 0 {
			break
		}
		procTranslateMessage.Call(uintptr(unsafe.Pointer(&msg)))
		procDispatchMessageW.Call(uintptr(unsafe.Pointer(&msg)))
	}
}
