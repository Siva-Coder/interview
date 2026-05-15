package main

import (
	"bufio"
	"bytes"
	"os"
	"runtime"
	"strconv"

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

	WIN_W = 460
	WIN_H = 500

	// Close button
	CLOSE_X    int32 = WIN_W - 38
	CLOSE_Y    int32 = 14
	CLOSE_SIZE int32 = 22

	// Hand Cursor
	IDC_HAND = 32649

	// Stop/resume button (left of close)
	STOP_X    int32 = WIN_W - 68
	STOP_Y    int32 = 14
	STOP_SIZE int32 = 22

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
	TIMER_PING      = 4 // keepalive ping timer

	// Direct2D / DirectWrite
	DWRITE_FONT_STYLE_ITALIC                          = 1
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

	// WAVE format tags
	WAVE_FORMAT_PCM        = 0x0001
	WAVE_FORMAT_IEEE_FLOAT = 0x0003
	WAVE_FORMAT_EXTENSIBLE = 0xFFFE

	// AssemblyAI
	ASSEMBLYAI_WS_URL = "wss://streaming.assemblyai.com/v3/ws"
	SAMPLE_RATE       = 16000
)

// ═══════════════════════════════════════════════════════════════════════════
//
//	Segoe Fluent Icons (PUA codepoints) - preinstalled on Windows 11
//
// ═══════════════════════════════════════════════════════════════════════════
const (
	ICON_PLAY       = "\uE768" // Play
	ICON_PAUSE      = "\uE769" // Pause
	ICON_STOP       = "\uE71A" // Stop
	ICON_CLOSE      = "\uE711" // Cancel / X
	ICON_MIC        = "\uE720" // Microphone
	ICON_SEND       = "\uE724" // Send
	ICON_SETTINGS   = "\uE713" // Setting
	ICON_CHEVRON_UP = "\uE70E" // ChevronUp
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

// audioFormatInfo holds the decoded properties we actually need for conversion.
type audioFormatInfo struct {
	sampleRate    int
	channels      int
	bitsPerSample int  // container bits per sample
	validBits     int  // valid bits (may be less for 24-in-32 packed)
	isFloat       bool // true = IEEE float, false = integer PCM
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
	procKillTimer                  = user32.NewProc("KillTimer")
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
	clrBg         = rgba(0x0D, 0x0D, 0x18, 0.88) // semi-transparent for glass
	clrHeaderBg   = rgba(0x0D, 0x0D, 0x18, 0.96) // slightly more opaque header
	clrGreen      = rgba(0x34, 0xd3, 0x99, 1)
	clrGreenDim   = rgba(0x10, 0x99, 0x66, 1)
	clrGreenMuted = rgba(0x0a, 0x5c, 0x40, 1)
	clrTextMain   = rgba(0xf0, 0xf4, 0xff, 1)
	clrTextDim    = rgba(0x9a, 0xa2, 0xb8, 1)
	clrRed        = rgba(0xf8, 0x71, 0x71, 1)
	clrBtnBg      = rgba(0x23, 0x8b, 0x5a, 1)
	clrRedBright  = rgba(0xff, 0x4d, 0x4d, 1)
	clrUserBubble = rgba(0x1a, 0x3d, 0x6a, 0.9)
	clrAiLine     = rgba(0x34, 0xd3, 0x99, 0.25) // subtle left accent line for AI
	clrBorder     = rgba(0xff, 0xff, 0xff, 0.22) // more visible for glass edge
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
	brushAiLine                                         uintptr
	brushBorder                                         uintptr
	brushHeaderBg                                       uintptr

	fmtTitle, fmtSub, fmtBody, fmtBadge, fmtSmall, fmtIcon uintptr

	hwndMain   windows.Handle
	windowX    int32
	windowY    int32
	closeHover bool
	stopHover  bool

	fmtBodyBold   uintptr
	fmtBodyItalic uintptr
	fmtH1         uintptr
	fmtH2         uintptr
	fmtH3plus     uintptr

	aaiSessionReady    bool
	aaiSessionReadyMux sync.Mutex
)

// ═══════════════════════════════════════════════════════════════════════════
//  App State
// ═══════════════════════════════════════════════════════════════════════════

type Message struct {
	Role         string
	Text         string
	Lines        []string
	Spans        []RichTextSpan
	WrappedSpans []RichTextSpan
	Timestamp    time.Time
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
	audioInfo      audioFormatInfo // decoded format properties (handles EXTENSIBLE)
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
	alphaValue      uint8 = 215 // body transparency only
	dragSlider      bool

	pendingSendCancel    context.CancelFunc
	pendingSendCancelMux sync.Mutex

	// UX: track if user manually scrolled
	userScrolled bool

	// Track last speech activity for smarter debouncing
	lastSpeechTime  time.Time
	lastSpeechMutex sync.Mutex

	lastAudioActivity  time.Time
	audioActivityMutex sync.Mutex
	aaiReconnecting    bool
	aaiReconnectMux    sync.Mutex

	// Buffered channel for audio chunks — written on UI thread, drained in goroutine
	audioSendCh chan []byte
)

// safelySend sends to ch without panicking if ch was closed between the caller
// capturing the pointer and the actual send (possible during reconnect).
func safelySend(ch chan []byte, chunk []byte) {
	defer func() { recover() }()
	select {
	case ch <- chunk:
	default: // drop on backpressure
	}
}

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

// setBrushColor updates an existing ID2D1SolidColorBrush colour (vtable slot 8).
// Used to change background opacity at runtime without recreating brushes.
func setBrushColor(brush uintptr, c D2D1_COLOR_F) {
	if brush == 0 {
		return
	}
	comCall(brush, 8, uintptr(unsafe.Pointer(&c)))
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
	labelH := float32(18)
	lineH := float32(22)
	gap := float32(12)
	headingH := float32(26)
	codeLineH := float32(16)

	for _, msg := range conversation {
		// msgH must match renderFrame exactly: labelH + span heights + gap
		msgH := labelH
		if len(msg.Spans) > 0 {
			spans := msg.WrappedSpans
			if len(spans) == 0 {
				spans = msg.Spans
			}
			for _, span := range spans {
				if span.Heading > 0 {
					msgH += headingH
				} else if span.Code {
					codeLines := float32(strings.Count(span.Text, "\n") + 1)
					msgH += codeLines*codeLineH + 8
				} else {
					msgH += lineH
				}
			}
		} else {
			msgH += float32(len(msg.Lines)) * lineH
		}
		msgH += gap
		totalMsgHeight += msgH
	}

	// Include streaming AI message height so clamping works during generation
	partialAIResponseMutex.Lock()
	aiPartial := partialAIResponse
	streaming := isAIResponding && aiPartial != ""
	partialAIResponseMutex.Unlock()

	if streaming {
		lineCount := len(wrapText(aiPartial, 52))
		totalMsgHeight += labelH + float32(lineCount)*lineH + gap
	}
}

// clamp scrollOffset to valid range
func clampScroll() {
	availableH := float32(WIN_H) - 54 - 46 // chat view area
	maxScroll := totalMsgHeight - availableH
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
	return y >= 54 && y <= WIN_H-46
}

func inSliderArea(x, y int32) bool {
	return x >= SLIDER_X && x <= SLIDER_X+SLIDER_W && y >= SLIDER_Y && y <= SLIDER_Y+SLIDER_H
}

func updateAlphaFromMouse(mx int32) {
	thumbRange := float32(SLIDER_W - SLIDER_THUMB_W)
	newAlpha := uint8((float32(mx-SLIDER_X) / thumbRange) * 255)
	if newAlpha < 30 {
		newAlpha = 30 // minimum: nearly-transparent glass
	}
	alphaValue = newAlpha

	// Slider now controls the *background overlay* opacity only.
	// Text, borders, and all controls remain fully opaque (their brushes are unaffected).
	// This gives a "real glass" look: background tint fades out, text stays 100% readable.
	bgA := float32(alphaValue) / 255.0
	setBrushColor(brushBg, rgba(0x0D, 0x0D, 0x18, bgA*0.95))
	// Header is always at least 10% more opaque than body so it masks scrolling chat text
	headerA := bgA*0.95 + 0.10
	if headerA > 1.0 {
		headerA = 1.0
	}
	setBrushColor(brushHeaderBg, rgba(0x0D, 0x0D, 0x18, headerA))
	if hwndMain != 0 {
		procInvalidateRect.Call(uintptr(hwndMain), 0, 1)
	}
}

func renderSpans(rt uintptr, spans []RichTextSpan, x, y, maxW, viewBottom float32, isVirtual, isUser bool) {
	lineH := float32(22)
	headingH := float32(26)
	codeLineH := float32(16)
	currentY := y

	for si, span := range spans {
		if currentY > viewBottom {
			break
		}

		// Choose format based on style - USE CACHED FORMATS
		var format uintptr
		var brush uintptr

		switch {
		case span.Heading == 1:
			format = fmtH1
			brush = brushTextMain
		case span.Heading == 2:
			format = fmtH2
			brush = brushTextMain
		case span.Heading >= 3:
			format = fmtH3plus
			brush = brushTextMain
		case span.Code:
			format = fmtSmall // Use small font for code
			brush = brushTextMain
		case span.Bold:
			format = fmtBodyBold
			brush = brushTextMain
		case span.Italic:
			format = fmtBodyItalic
			brush = brushTextDim
		case span.Blockquote:
			format = fmtBody
			brush = brushTextDim
		default:
			format = fmtBody
			brush = brushTextMain
		}

		// Handle bullet points
		displayText := span.Text
		if span.Bullet {
			displayText = "• " + displayText
		} else if span.Numbered > 0 {
			displayText = fmt.Sprintf("%d. %s", span.Numbered, displayText)
		}

		// Cursor for streaming
		if isVirtual && si == len(spans)-1 {
			if (time.Now().UnixNano()/400000000)%2 == 0 {
				displayText += "▌"
			}
		}

		// Render code blocks with background
		if span.Code {
			codeH := float32(strings.Count(span.Text, "\n")+1)*codeLineH + 8
			// Dark background for code
			codeBg := rgba(0x1a, 0x1a, 0x2e, 0.9)
			var codeBrush uintptr
			comCall(rt, 8, uintptr(unsafe.Pointer(&codeBg)), 0, uintptr(unsafe.Pointer(&codeBrush)))
			d2dFillRoundedRect(rt, codeBrush, rectF(x, currentY, x+maxW, currentY+codeH), 6, 6)

			// Render code text line by line
			codeLines := strings.Split(span.Text, "\n")
			codeY := currentY + 4
			for _, cl := range codeLines {
				if codeY > viewBottom {
					break
				}
				d2dText(rt, cl, fmtSmall, brushTextMain, rectF(x+8, codeY, x+maxW-8, codeY+codeLineH))
				codeY += codeLineH
			}
			currentY += codeH
		} else if span.Heading > 0 {
			d2dText(rt, displayText, format, brush, rectF(x, currentY, x+maxW, currentY+headingH))
			currentY += headingH
		} else {
			// Blockquote indent
			indentX := x
			if span.Blockquote {
				indentX += 12
				// Quote line
				quoteBrush := brushGreenMuted
				d2dFillRect(rt, quoteBrush, rectF(x, currentY, x+3, currentY+lineH))
			}

			d2dText(rt, displayText, format, brush, rectF(indentX, currentY, x+maxW, currentY+lineH))
			currentY += lineH
		}
	}
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
	brushAiLine = createBrush(pRenderTarget, clrAiLine)
	brushBorder = createBrush(pRenderTarget, clrBorder)
	brushHeaderBg = createBrush(pRenderTarget, clrHeaderBg)

	// Larger font sizes for better readability
	fmtTitle = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 19)
	fmtSub = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 13)
	fmtBody = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 15)
	fmtBadge = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_SEMI_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 11)
	fmtSmall = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 11)
	// Segoe Fluent Icons font for icons - size 18 for header buttons
	fmtIcon = createTextFormat("Segoe Fluent Icons", DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 18)

	// Cached formats for markdown rendering
	fmtBodyBold = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 15)
	fmtBodyItalic = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_NORMAL, DWRITE_FONT_STYLE_ITALIC, DWRITE_FONT_STRETCH_NORMAL, 15)
	fmtH1 = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 20)
	fmtH2 = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 18)
	fmtH3plus = createTextFormat("Segoe UI Variable", DWRITE_FONT_WEIGHT_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, 16)
	return nil
}

// ═══════════════════════════════════════════════════════════════════════════
//  RENDERING
// ═══════════════════════════════════════════════════════════════════════════

type displayItem struct {
	msg   Message
	spans []RichTextSpan
}

func renderFrame() {
	if pRenderTarget == 0 {
		return
	}
	W := float32(WIN_W)

	comCall(pRenderTarget, 48) // BeginDraw

	// ── Glass effect: clear to transparent, let DWM blur show through ───────
	// Background opacity is controlled by the slider (alphaValue), NOT LWA_ALPHA.
	// Text and controls are drawn at full alpha so they are always fully readable.
	transparent := rgba(0, 0, 0, 0)
	comCall(pRenderTarget, 47, uintptr(unsafe.Pointer(&transparent))) // Clear transparent

	// Semi-transparent rounded window background (inset 1px so rounded corners clip cleanly)
	d2dFillRoundedRect(pRenderTarget, brushBg, rectF(1, 1, W-1, float32(WIN_H)-1), 19, 19)

	sessionMutex.Lock()
	active := sessionActive
	sessionMutex.Unlock()

	// ── Glass border (drawn on top of bg fill) ──────────────────────────────
	borderRR := D2D1_ROUNDED_RECT{Rect: rectF(0.5, 0.5, W-0.5, float32(WIN_H)-0.5), RadiusX: 20, RadiusY: 20}
	comCall(pRenderTarget, 18, uintptr(unsafe.Pointer(&borderRR)), brushBorder, uintptr(mathFloat32bits(1.5)), 0)

	// ── Chat content (drawn FIRST, will be covered by opaque header) ────
	// This ensures any chat text scrolling up gets hidden behind the header
	const (
		viewTop    = float32(54)
		viewBottom = float32(WIN_H) - 46
	)
	availableH := viewBottom - viewTop

	var items []displayItem
	lineH := float32(22)
	labelH := float32(18)
	gap := float32(12)
	headingH := float32(26)
	codeLineH := float32(16)

	convMutex.RLock()
	messages := make([]Message, len(conversation))
	copy(messages, conversation)
	convMutex.RUnlock()

	// Append virtual AI message if streaming so it scrolls naturally
	partialAIResponseMutex.Lock()
	aiPartial := partialAIResponse
	streaming := isAIResponding && aiPartial != ""
	partialAIResponseMutex.Unlock()

	if streaming {
		// Plain-text virtual message — no markdown parsing during stream
		lines := wrapText(aiPartial, 52)
		var plainSpans []RichTextSpan
		for _, line := range lines {
			plainSpans = append(plainSpans, RichTextSpan{Text: line})
		}
		virtualMsg := Message{
			Role: "assistant",
			Text: aiPartial,
			// Spans:     parseMarkdown(aiPartial), // parse markdown so it renders correctly while streaming
			Spans:     plainSpans,
			Timestamp: time.Now(),
		}
		messages = append(messages, virtualMsg)
	}

	scrollMutex.Lock()
	offset := scrollOffset
	scrollMutex.Unlock()

	// Pre‑compute heights from the local slice
	var localTotalH float32
	for _, msg := range messages {
		var spans []RichTextSpan
		if len(msg.Spans) > 0 {
			// Use cached wrapped spans if available and message is finalized
			if msg.Role == "assistant" && len(msg.WrappedSpans) > 0 && !streaming {
				spans = msg.WrappedSpans
			} else if msg.Role == "user" && len(msg.WrappedSpans) > 0 {
				spans = msg.WrappedSpans
			} else {
				spans = wrapMarkdownSpans(msg.Spans, 52)
			}
		} else {
			// Fallback for old messages without spans
			for _, line := range msg.Lines {
				spans = append(spans, RichTextSpan{Text: line})
			}
		}
		items = append(items, displayItem{msg: msg, spans: spans})

		// Calculate height
		msgH := labelH // "You" or "AI" label
		for _, span := range spans {
			if span.Heading > 0 {
				msgH += headingH
			} else if span.Code {
				codeLines := float32(strings.Count(span.Text, "\n") + 1)
				msgH += codeLines*codeLineH + 8
			} else {
				msgH += lineH
			}
		}
		msgH += gap
		localTotalH += msgH
	}

	// Top-pinned: content starts at viewTop, scrollOffset moves it up.
	// Never auto-scrolls — position is 100% user-controlled.
	firstY := viewTop - offset

	yPos := firstY
	for i := 0; i < len(items); i++ {
		item := items[i]
		msgH := float32(0)
		for _, span := range item.spans {
			if span.Heading > 0 {
				msgH += headingH
			} else if span.Code {
				codeLines := float32(strings.Count(span.Text, "\n") + 1)
				msgH += codeLines*codeLineH + 8
			} else {
				msgH += lineH
			}
		}
		msgH += labelH + gap

		if yPos+msgH < viewTop {
			yPos += msgH
			continue
		}
		if yPos > viewBottom {
			break
		}

		isVirtual := (i == len(items)-1 && streaming && item.msg.Role == "assistant")

		if item.msg.Role == "user" {
			// User bubble
			bubbleW := float32(300)
			bubbleX := W - bubbleW - 20
			bubbleY := yPos
			totalBubbleH := msgH - gap
			d2dFillRoundedRect(pRenderTarget, brushUserBubble, rectF(bubbleX, bubbleY, bubbleX+bubbleW, bubbleY+totalBubbleH), 12, 12)
			d2dText(pRenderTarget, "You", fmtBadge, brushGreenDim, rectF(bubbleX+10, bubbleY+6, bubbleX+50, bubbleY+labelH+6))

			renderSpans(pRenderTarget, item.spans, bubbleX+10, bubbleY+labelH+4, bubbleW-20, viewBottom, isVirtual, true)
		} else {
			// AI message
			textX := float32(18)
			textW := W - 18 - 8
			msgY := yPos

			d2dText(pRenderTarget, "AI", fmtBadge, brushGreen, rectF(textX, msgY+4, textX+40, msgY+labelH+4))

			renderSpans(pRenderTarget, item.spans, textX, msgY+labelH+4, textW, viewBottom, isVirtual, false)
		}
		yPos += msgH
	}

	// ── Scrollbar ───────────────────────────────────────────────────────
	if localTotalH > availableH {
		frac := availableH / localTotalH
		barH := frac * availableH
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
			if scrollFrac < 0 {
				scrollFrac = 0
			}
		}
		barY := viewBottom - barH - scrollFrac*(availableH-barH)
		d2dFillRoundedRect(pRenderTarget, brushGreenDim, rectF(W-5, barY, W-2, barY+barH), 3, 3)
	}

	// ── Footer: Live partial transcript & Typing indicator ──────────────
	footerY := float32(WIN_H) - 36

	partialMutex.RLock()
	partial := partialTranscript
	partialMutex.RUnlock()

	if partial != "" {
		d2dText(pRenderTarget, "You", fmtBadge, brushGreenMuted, rectF(18, footerY, 50, footerY+labelH))
		cursor := ""
		if (time.Now().UnixNano()/400000000)%2 == 0 {
			cursor = "▌"
		}
		d2dText(pRenderTarget, partial+cursor, fmtBody, brushTextDim,
			rectF(50, footerY, W-18, footerY+lineH))
	} else if isProcessing && !isAIResponding {
		// Animated pulsing dots
		t := float64(time.Now().UnixNano()) / 1e9
		dotBaseX := float32(40)
		dotY := footerY + 8
		for i := 0; i < 3; i++ {
			phase := t*4.0 + float64(i)*2.1
			scale := float32(0.4 + 0.6*math.Sin(phase))
			if scale < 0 {
				scale = 0
			}
			r := float32(2 + 2*scale)
			d2dFillEllipse(pRenderTarget, brushGreen, dotBaseX+float32(i)*14, dotY, r, r)
		}
		d2dText(pRenderTarget, "Thinking", fmtSmall, brushTextDim, rectF(82, footerY+2, 200, footerY+18))
	}

	// ═════════════════════════════════════════════════════════════════════
	//  HEADER (drawn LAST so it covers any chat content underneath)
	// ═════════════════════════════════════════════════════════════════════
	headerH := float32(54)
	// Opaque header background - covers everything underneath
	d2dFillRect(pRenderTarget, brushHeaderBg, rectF(0, 0, W, headerH))

	// Title
	d2dText(pRenderTarget, "Candy", fmtTitle, brushGreen, rectF(18, 12, 80, 44))
	d2dText(pRenderTarget, " Ai", fmtTitle, brushGreenDim, rectF(82, 12, 170, 44))

	// Live status dot
	aaiMutex.RLock()
	aaiConn := aaiConnected
	aaiMutex.RUnlock()
	dotBrush := brushGreenMuted
	if isListening && aaiConn && active {
		dotBrush = brushGreen
	}
	d2dFillEllipse(pRenderTarget, dotBrush, 118, 28, 5, 5)

	// Stop / resume button using Segoe Fluent Icons
	var stopIcon string
	if active {
		stopIcon = ICON_STOP // Stop square icon
	} else {
		stopIcon = ICON_PLAY // Play triangle icon
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
	d2dText(pRenderTarget, stopIcon, fmtIcon, stopBr,
		rectF(float32(STOP_X), float32(STOP_Y), float32(STOP_X+STOP_SIZE), float32(STOP_Y+STOP_SIZE)))

	// Close button using Segoe Fluent Icons (Cancel/X)
	closeBrush := brushTextDim
	if closeHover {
		closeBrush = brushRed
	}
	d2dText(pRenderTarget, ICON_CLOSE, fmtIcon, closeBrush,
		rectF(float32(CLOSE_X), float32(CLOSE_Y), float32(CLOSE_X+CLOSE_SIZE), float32(CLOSE_Y+CLOSE_SIZE)))

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

	// Divider line at bottom of header
	d2dFillRect(pRenderTarget, brushGreenDim, rectF(18, 48, W-18, 49))

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
	params.Add("min_turn_silence", "800")  // 800ms silence = end of sentence
	params.Add("max_turn_silence", "3000") // allow 3s pauses within a turn

	wsURL := fmt.Sprintf("%s?%s", ASSEMBLYAI_WS_URL, params.Encode())
	headers := http.Header{}
	headers.Add("Authorization", assemblyAIKey)

	dialer := websocket.Dialer{
		HandshakeTimeout:  10 * time.Second,
		EnableCompression: true,
	}
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

	// Configure connection for longevity
	conn.SetReadDeadline(time.Time{}) // Clear any deadline - we handle timeouts ourselves
	// FIX 2: Enable ping/pong keepalive to prevent i/o timeout
	conn.SetPingHandler(func(appData string) error {
		return conn.WriteControl(websocket.PongMessage, []byte(appData), time.Now().Add(5*time.Second))
	})

	// Reset session-ready flag so the drain goroutine won't forward audio
	// until AssemblyAI sends the Begin message for this new connection.
	aaiSessionReadyMux.Lock()
	aaiSessionReady = false
	aaiSessionReadyMux.Unlock()

	// Close the old send channel (if any) to stop the previous drain goroutine,
	// then create a fresh one. Do this BEFORE starting handleAssemblyAIResponses
	// so there is no window where the old channel is still live alongside the
	// new connection.
	aaiMutex.Lock()
	if audioSendCh != nil {
		close(audioSendCh) // signals the old drain goroutine to exit
		audioSendCh = nil
	}
	newCh := make(chan []byte, 100)
	audioSendCh = newCh
	aaiWSConn = conn
	aaiConnected = true
	aaiMutex.Unlock()

	fmt.Println("✅ Connected to AssemblyAI")

	// Drain goroutine: forwards audio to *this specific connection* only.
	// Using captured conn and newCh prevents races with future reconnects.
	go func(ch chan []byte, c *websocket.Conn) {
		for chunk := range ch {
			if err := c.WriteMessage(websocket.BinaryMessage, chunk); err != nil {
				// Connection gone; exit so the reconnect can start a fresh goroutine.
				return
			}
		}
	}(newCh, conn)

	go handleAssemblyAIResponses()

	// Re-assert isListening in case this is a reconnect (session was active but
	// the AssemblyAI session got terminated and we dialled a fresh one).
	sessionMutex.Lock()
	if sessionActive {
		isListening = true
	}
	sessionMutex.Unlock()

	return nil
}

func handleAssemblyAIResponses() {
	fmt.Println("👂 AssemblyAI response handler started")

	// Start ping ticker to keep connection alive
	pingTicker := time.NewTicker(10 * time.Second)
	defer pingTicker.Stop()

	pingDone := make(chan struct{})
	defer close(pingDone)

	// Run ping in separate goroutine
	go func() {
		for {
			select {
			case <-pingTicker.C:
				aaiMutex.RLock()
				conn := aaiWSConn
				aaiMutex.RUnlock()
				if conn == nil {
					return
				}
				// Send ping frame
				if err := conn.WriteControl(websocket.PingMessage, []byte("keepalive"), time.Now().Add(5*time.Second)); err != nil {
					fmt.Printf("⚠️ AssemblyAI ping failed: %v\n", err)
				}
			case <-pingDone:
				return
			}
		}
	}()

	for {
		aaiMutex.RLock()
		conn := aaiWSConn
		aaiMutex.RUnlock()
		if conn == nil {
			break
		}

		msgType, data, err := conn.ReadMessage()
		if err != nil {
			// Any error (timeout, close, network) – treat as disconnection.
			fmt.Printf("❌ AssemblyAI read error: %v\n", err)
			aaiMutex.Lock()
			if aaiWSConn == conn {
				aaiConnected = false
				aaiWSConn = nil
			}
			aaiMutex.Unlock()

			// Auto-reconnect if session is still active
			sessionMutex.Lock()
			shouldReconnect := sessionActive
			sessionMutex.Unlock()

			if shouldReconnect {
				go func() {
					for attempt := 1; attempt <= 10; attempt++ {
						sessionMutex.Lock()
						stillActive := sessionActive
						sessionMutex.Unlock()
						if !stillActive {
							return
						}
						fmt.Printf("🔄 Reconnecting to AssemblyAI (attempt %d)...\n", attempt)
						if err := connectAssemblyAI(); err == nil {
							fmt.Println("✅ Reconnected, listening again")
							return
						}
						sleep := time.Duration(attempt) * time.Second
						if sleep > 8*time.Second {
							sleep = 8 * time.Second
						}
						time.Sleep(sleep)
					}
					fmt.Println("❌ Failed to reconnect to AssemblyAI after 10 attempts")
				}()
			}
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

			aaiSessionReadyMux.Lock()
			aaiSessionReady = true
			aaiSessionReadyMux.Unlock()
		case "Turn":
			var msg AAITurnMessage
			json.Unmarshal(data, &msg)
			handleTurnMessage(msg)
		case "Termination":
			var msg AAITerminationMessage
			json.Unmarshal(data, &msg)
			fmt.Printf("🔴 Session terminated\n")
			aaiMutex.Lock()
			if aaiWSConn == conn {
				aaiConnected = false
				aaiWSConn = nil
			}
			aaiMutex.Unlock()

			aaiSessionReadyMux.Lock()
			aaiSessionReady = false
			aaiSessionReadyMux.Unlock()

			// Auto-reconnect on termination if session still active
			sessionMutex.Lock()
			shouldReconnect := sessionActive
			sessionMutex.Unlock()

			if shouldReconnect {
				go func() {
					for attempt := 1; attempt <= 10; attempt++ {
						sessionMutex.Lock()
						stillActive := sessionActive
						sessionMutex.Unlock()
						if !stillActive {
							return
						}
						fmt.Printf("🔄 Reconnecting to AssemblyAI (attempt %d)...\n", attempt)
						if err := connectAssemblyAI(); err == nil {
							fmt.Println("✅ Reconnected, listening again")
							return
						}
						sleep := time.Duration(attempt) * time.Second
						if sleep > 8*time.Second {
							sleep = 8 * time.Second
						}
						time.Sleep(sleep)
					}
					fmt.Println("❌ Failed to reconnect to AssemblyAI after 10 attempts")
				}()
			}
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

	if msg.EndOfTurn {
		// ── Final transcript received ─────────────────────────────────────
		transcriptMutex.Lock()
		currentTranscript = msg.Transcript
		transcriptMutex.Unlock()

		partialMutex.Lock()
		partialTranscript = ""
		partialMutex.Unlock()

		// Track when this end-of-turn arrived (used only for logging/debug now)
		lastSpeechMutex.Lock()
		lastSpeechTime = time.Now()
		lastSpeechMutex.Unlock()

		// Interrupt any in-progress AI stream — user has completed a new utterance
		aiCancelMux.Lock()
		if cancelAI != nil {
			cancelAI()
			cancelAI = nil
		}
		aiCancelMux.Unlock()
		isProcessing = false
		isAIResponding = false

		partialAIResponseMutex.Lock()
		partialAIResponse = ""
		partialAIResponseMutex.Unlock()

		// ── Debounced send: 200ms after end-of-turn ───────────────────────
		// pendingSendCancel guards against a newer end-of-turn arriving in
		// the window; it is the only cancellation mechanism needed here.
		finalTranscript := msg.Transcript

		pendingSendCancelMux.Lock()
		if pendingSendCancel != nil {
			pendingSendCancel()
		}
		ctx, cancel := context.WithCancel(context.Background())
		pendingSendCancel = cancel
		pendingSendCancelMux.Unlock()

		go func(transcript string, ctx context.Context) {
			select {
			case <-time.After(1000 * time.Millisecond):
				if transcript != "" {
					sendToAI(transcript)
				}
			case <-ctx.Done():
				// cancelled: a newer end-of-turn arrived, let it handle sending
			}
		}(finalTranscript, ctx)

	} else {
		// ── Partial transcript (user is actively speaking) ────────────────
		// Only update the display — do NOT cancel pending sends or AI streams.
		// Cancelling here caused premature mid-sentence sends whenever the user
		// took a breath, because each breath reset the whole pipeline.
		partialMutex.Lock()
		partialTranscript = msg.Transcript
		partialMutex.Unlock()

		// Cancel any pending debounced send — user is still speaking
		pendingSendCancelMux.Lock()
		if pendingSendCancel != nil {
			pendingSendCancel()
			pendingSendCancel = nil
		}
		pendingSendCancelMux.Unlock()
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
	// FIX: avoid deadlock by grabbing conn pointer under lock, then closing outside lock
	aaiMutex.Lock()
	conn := aaiWSConn
	aaiWSConn = nil
	aaiConnected = false

	// Close old channel to unblock sender goroutine
	if audioSendCh != nil {
		close(audioSendCh)
		audioSendCh = nil
	}

	aaiMutex.Unlock()

	aaiSessionReadyMux.Lock()
	aaiSessionReady = false
	aaiSessionReadyMux.Unlock()

	if conn != nil {
		terminate := AAITerminateMessage{Type: "Terminate"}
		data, _ := json.Marshal(terminate)
		conn.WriteMessage(websocket.TextMessage, data)
		conn.Close()
	}
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

	// Decode the format — WASAPI almost always returns WAVEFORMATEXTENSIBLE
	// (FormatTag == 0xFFFE). Reading WBitsPerSample from the base struct gives
	// the *container* size, not the valid bits, which causes garbled audio when
	// the format is e.g. 24-valid-bits-in-32-bit-containers.
	info := audioFormatInfo{
		sampleRate:    int(audioFormat.NSamplesPerSec),
		channels:      int(audioFormat.NChannels),
		bitsPerSample: int(audioFormat.WBitsPerSample),
		validBits:     int(audioFormat.WBitsPerSample),
		isFloat:       false,
	}
	if audioFormat.WFormatTag == WAVE_FORMAT_IEEE_FLOAT {
		info.isFloat = true
	} else if audioFormat.WFormatTag == WAVE_FORMAT_EXTENSIBLE && audioFormat.CBSize >= 22 {
		// Go pads WAVEFORMATEX to 20 bytes, so WAVEFORMATEXTENSIBLE overlay is misaligned.
		// Read extended fields at the raw Windows memory offsets directly.
		base := unsafe.Pointer(pFormat)
		info.validBits = int(*(*uint16)(unsafe.Pointer(uintptr(base) + 18)))
		// SubFormat GUID starts at offset 24; first 4 bytes identify PCM vs FLOAT
		subTag := binary.LittleEndian.Uint32((*[4]byte)(unsafe.Pointer(uintptr(base) + 24))[:])
		if subTag == WAVE_FORMAT_IEEE_FLOAT {
			info.isFloat = true
		}
	}
	audioInfo = info

	fmt.Printf("🎧 Loopback capture: %d Hz, %d ch, %d-bit container (%d valid), float=%v\n",
		info.sampleRate, info.channels, info.bitsPerSample, info.validBits, info.isFloat)

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
		data := unsafe.Slice((*byte)(unsafe.Pointer(pData)), bytesAvailable)
		converted := convertAudioFormat(data)

		if flags&AUDCLNT_BUFFERFLAGS_SILENT == 0 {
			audioActivityMutex.Lock()
			lastAudioActivity = time.Now()
			audioActivityMutex.Unlock()
		}

		// Always buffer converted audio (including silent frames) so AssemblyAI
		// receives a continuous stream for accurate pause/turn detection.
		audioMutex.Lock()
		audioBuffer = append(audioBuffer, converted...)
		audioMutex.Unlock()
		comCall(pCaptureClient, 4, uintptr(framesAvailable))
	}

	// AssemblyAI v3 requires audio chunks between 50ms and 1000ms.
	// We accumulate 100ms (3200 bytes at 16kHz 16-bit mono) before sending.
	// The capture timer fires every 20ms, so we only send once we have enough
	// buffered — this keeps transmission at real-time rate (100ms sent per 100ms).
	const chunkBytes = SAMPLE_RATE * 2 * 100 / 1000 // 100ms = 3200 bytes

	// Only send one chunk per audioTick call to avoid exceeding real-time rate.
	audioMutex.Lock()
	bufLen := len(audioBuffer)
	if bufLen < chunkBytes {
		audioMutex.Unlock()
		return
	}
	chunk := make([]byte, chunkBytes)
	copy(chunk, audioBuffer[:chunkBytes])
	audioBuffer = audioBuffer[chunkBytes:]
	audioMutex.Unlock()

	aaiMutex.RLock()
	ch := audioSendCh
	aaiMutex.RUnlock()
	if ch != nil {
		safelySend(ch, chunk)
	}
}

func stopAudioDevice() {
	if audioStarted {
		if pAudioClient != 0 {
			comCall(pAudioClient, 11) // Stop
		}
		audioStarted = false
		procKillTimer.Call(uintptr(hwndMain), TIMER_AUDIO_CAP) // actually kill the timer
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

	// Run blocking network + COM init off the UI thread so the message loop
	// never stalls. Once everything is ready we start the audio capture timer
	// via PostMessage-equivalent (SetTimer is safe to call cross-thread on Win32).
	go func() {
		if !audioStarted {
			if err := connectAssemblyAI(); err != nil {
				fmt.Printf("⚠️  AssemblyAI connection failed: %v\n", err)
			}
			if err := initAudioDevice(); err != nil {
				fmt.Printf("⚠️  Audio init failed: %v\n", err)
				sessionMutex.Lock()
				sessionActive = false
				sessionMutex.Unlock()
				if hwndMain != 0 {
					procInvalidateRect.Call(uintptr(hwndMain), 0, 1)
				}
				return
			}
		} else {
			// Audio already set up, just restart client if stopped
			comCall(pAudioClient, 10) // Start
		}

		// SetTimer with a window handle is safe to call from any thread.
		if hwndMain != 0 {
			procSetTimer.Call(uintptr(hwndMain), TIMER_AUDIO_CAP, 20, 0)
		}

		isListening = true

		audioActivityMutex.Lock()
		lastAudioActivity = time.Now()
		audioActivityMutex.Unlock()

		fmt.Println("▶ Session started")
		if hwndMain != 0 {
			procInvalidateRect.Call(uintptr(hwndMain), 0, 1)
		}

		// Continuous silence keepalive: always send 100ms of silence every 100ms
		// when no real audio is in the buffer. This guarantees AssemblyAI never
		// sees a dead connection, which was causing it to drop the session.
		const silenceChunkBytes = SAMPLE_RATE * 2 * 100 / 1000 // 100ms = 3200 bytes
		silenceChunk := make([]byte, silenceChunkBytes)        // zero = silence
		go func() {
			ticker := time.NewTicker(100 * time.Millisecond)
			defer ticker.Stop()
			for range ticker.C {
				sessionMutex.Lock()
				active := sessionActive
				sessionMutex.Unlock()
				if !active {
					return
				}
				// Only inject silence when the real audio buffer is empty —
				// real audio always takes priority.
				audioMutex.Lock()
				bufEmpty := len(audioBuffer) < silenceChunkBytes
				audioMutex.Unlock()
				if bufEmpty {
					aaiMutex.RLock()
					ch := audioSendCh
					aaiMutex.RUnlock()
					if ch != nil {
						safelySend(ch, silenceChunk)
					}
				}
			}
		}()
	}()
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
	go disconnectAssemblyAI() // run off UI thread — contains network I/O

	fmt.Println("■ Session stopped")
	if hwndMain != 0 {
		procInvalidateRect.Call(uintptr(hwndMain), 0, 1)
	}
}

// ═══════════════════════════════════════════════════════════════════════════
//  AI Integration (with cancellation)
// ═══════════════════════════════════════════════════════════════════════════

func streamAIResponse(ctx context.Context, userText string) {
	// 60fps refresh while streaming for smooth typing feel
	procSetTimer.Call(uintptr(hwndMain), TIMER_AI_POLL, 16, 0)

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

	msgs = append(msgs, openaiChatMessage{
		Role: "system",
		Content: `You are a real-time interview assistant helping a candidate answer interview questions naturally and confidently.

When the candidate hears a question, your job is to give them a clear, concise answer they can speak out loud immediately — not read from a screen.

Rules:
- Use simple, everyday words. No jargon, no buzzwords unless the question specifically requires them.
- Keep answers short: 3 to 5 sentences max for most questions. If it needs more, break it into clear points.
- Write the way a person talks, not the way a textbook reads. Contractions are fine. Short sentences are better.
- Always lead with the direct answer first, then briefly explain or give one example.
- Never use filler phrases like "That's a great question", "Certainly!", "Absolutely!", or "As an AI...".
- Never use bullet points or markdown formatting — the candidate needs to speak this, not read a list.
- If the question is behavioral (tell me about a time when...), give a short story structure: situation, what they did, what happened. Keep it under 60 seconds of speaking time.
- If the question is technical, explain it simply as if to someone smart but not in that field.
- If the question is vague or unclear, pick the most likely intent and answer that.
- Stay grounded and humble in tone — confident but not arrogant.

The candidate will speak the answer out loud, so every word you write should sound natural when read aloud.`,
	})

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

	// No overall Timeout — SSE streaming responses may run indefinitely.
	// Cancellation is handled via the request context (ctx).
	client := &http.Client{}
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
		// FIX 3: Limit scanner buffer to prevent memory explosion on malformed data
		const maxScanTokenSize = 64 * 1024 // 64KB max per SSE line
		buf := make([]byte, 4096)
		scanner.Buffer(buf, maxScanTokenSize)

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

				if hwndMain != 0 {
					procInvalidateRect.Call(uintptr(hwndMain), 0, 1)
				}
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
		// Close the body to unblock scanner.Scan(), then wait for the goroutine
		// to exit before we read fullResponse — otherwise it's a data race.
		doneClosing := make(chan struct{})
		go func() {
			resp.Body.Close()
			close(doneClosing)
		}()
		select {
		case <-doneClosing:
		case <-time.After(2 * time.Second):
			fmt.Println("⚠️ Force closing response body timed out")
		}
		// Wait for scanner goroutine to drain and close `done`
		select {
		case <-done:
		case <-time.After(1 * time.Second):
		}
	case <-done:
	}

	answer := fullResponse.String()
	if answer == "" {
		return
	}

	// Append assistant message to conversation
	aiSpans := parseMarkdown(answer)
	wrapped := wrapMarkdownSpans(aiSpans, 52)
	convMutex.Lock()
	conversation = append(conversation, Message{
		Role:         "assistant",
		Text:         answer,
		Spans:        aiSpans,
		WrappedSpans: wrapped,
		Timestamp:    time.Now(),
	})
	convMutex.Unlock()

	updateTotalMessageHeight()
	// Do not reset scrollOffset here — preserve the user's reading position.
	// The next user question (sendToAI) will scroll to show it.
}

func sendToAI(userText string) {
	isProcessing = true
	isAIResponding = true

	// Parse user text (minimal markdown)
	userSpans := parseMarkdown(userText)
	wrapped := wrapMarkdownSpans(userSpans, 52)
	convMutex.Lock()
	conversation = append(conversation, Message{
		Role:         "user",
		Text:         userText,
		Spans:        userSpans,
		WrappedSpans: wrapped,
		Timestamp:    time.Now(),
	})
	convMutex.Unlock()

	updateTotalMessageHeight()

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

// ═══════════════════════════════════════════════════════════════════════════
//  Audio format conversion (unchanged)
// ═══════════════════════════════════════════════════════════════════════════

func convertAudioFormat(data []byte) []byte {
	info := audioInfo
	if info.sampleRate == 0 {
		return data // not initialised yet
	}

	srcRate := info.sampleRate
	channels := info.channels
	bitsPerSample := info.bitsPerSample // container width
	isFloat := info.isFloat

	bytesPerFrame := channels * (bitsPerSample / 8)
	if bytesPerFrame == 0 || len(data)%bytesPerFrame != 0 {
		return data
	}
	frameCount := len(data) / bytesPerFrame

	monoFloat := make([]float32, frameCount)

	for i := 0; i < frameCount; i++ {
		frameOff := i * bytesPerFrame
		var sum float64
		for ch := 0; ch < channels; ch++ {
			sampleOff := frameOff + ch*(bitsPerSample/8)
			var s float32
			if isFloat && bitsPerSample == 32 {
				bits := binary.LittleEndian.Uint32(data[sampleOff:])
				s = math.Float32frombits(bits)
			} else if isFloat && bitsPerSample == 64 {
				bits := binary.LittleEndian.Uint64(data[sampleOff:])
				s = float32(math.Float64frombits(bits))
			} else if bitsPerSample == 32 {
				// 32-bit integer PCM (rare but exists)
				v := int32(binary.LittleEndian.Uint32(data[sampleOff:]))
				s = float32(v) / 2147483648.0
			} else if bitsPerSample == 24 {
				// 24-bit packed integer
				v := int32(uint32(data[sampleOff]) | uint32(data[sampleOff+1])<<8 | uint32(data[sampleOff+2])<<16)
				if v&0x800000 != 0 {
					v |= ^0xFFFFFF // sign-extend
				}
				s = float32(v) / 8388608.0
			} else if bitsPerSample == 16 {
				v := int16(binary.LittleEndian.Uint16(data[sampleOff:]))
				s = float32(v) / 32768.0
			}
			sum += float64(s)
		}
		monoFloat[i] = float32(sum / float64(channels))
	}

	// Resample to target rate using linear interpolation (better than box filter
	// for downsampling ratios that aren't integer multiples, e.g. 48000→16000).
	outRate := SAMPLE_RATE
	var resampled []float32
	if srcRate == outRate {
		resampled = monoFloat
	} else {
		ratio := float64(srcRate) / float64(outRate)
		outLen := int(float64(frameCount) / ratio)
		resampled = make([]float32, outLen)
		for i := 0; i < outLen; i++ {
			srcPos := float64(i) * ratio
			lo := int(srcPos)
			hi := lo + 1
			frac := float32(srcPos - float64(lo))
			if hi >= frameCount {
				hi = frameCount - 1
			}
			resampled[i] = monoFloat[lo]*(1-frac) + monoFloat[hi]*frac
		}
	}

	// Convert to 16-bit signed PCM little-endian (what AssemblyAI expects)
	out := make([]byte, len(resampled)*2)
	for i, f := range resampled {
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
			// enforceTopmostOnly removed from here — calling SetWindowPos 7×/sec
			// caused jank and reentrant WM_WINDOWPOSCHANGED storms on the UI thread.
			// Topmost is enforced once at startup and on WM_ACTIVATEAPP instead.
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
			userScrolled = (scrollOffset != 0)
			scrollMutex.Unlock()
			procInvalidateRect.Call(uintptr(hwnd), 0, 1)
			return 0
		}
		if wParam == 0x28 { // VK_DOWN
			scrollMutex.Lock()
			scrollOffset -= 40
			clampScroll()
			userScrolled = (scrollOffset != 0)
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
			userScrolled = (scrollOffset != 0)
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
		scrollOffset += float32(delta) / 120 * 40
		clampScroll()
		userScrolled = (scrollOffset != 0)
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

// wrapMarkdownSpans wraps parsed markdown spans into display lines
func wrapMarkdownSpans(spans []RichTextSpan, maxLen int) []RichTextSpan {
	var result []RichTextSpan
	for _, span := range spans {
		if span.Code || span.Heading > 0 {
			// Code blocks and headings don't wrap
			result = append(result, span)
			continue
		}

		lines := wrapText(span.Text, maxLen)
		for i, line := range lines {
			s := span
			s.Text = line
			if i > 0 {
				// Only first line keeps bullet/number
				s.Bullet = false
				s.Numbered = 0
			}
			result = append(result, s)
		}
	}
	return result
}

// ═══════════════════════════════════════════════════════════════════════════
//  Markdown rendering helpers
// ═══════════════════════════════════════════════════════════════════════════

// RichTextSpan represents a span of text with formatting
type RichTextSpan struct {
	Text       string
	Bold       bool
	Italic     bool
	Code       bool
	Heading    int  // 0 = normal, 1-6 = h1-h6
	Bullet     bool // unordered list item
	Numbered   int  // 0 = not numbered, >0 = list number
	LinkURL    string
	Blockquote bool
}

// parseMarkdown parses simple markdown into formatted spans
func parseMarkdown(text string) []RichTextSpan {
	var spans []RichTextSpan
	lines := strings.Split(text, "\n")

	inCodeBlock := false
	codeBlockContent := ""

	for _, line := range lines {
		trimmed := strings.TrimSpace(line)

		// Code blocks
		if strings.HasPrefix(trimmed, "```") {
			if inCodeBlock {
				// End code block
				spans = append(spans, RichTextSpan{
					Text: codeBlockContent,
					Code: true,
				})
				codeBlockContent = ""
				inCodeBlock = false
			} else {
				inCodeBlock = true
			}
			continue
		}

		if inCodeBlock {
			if codeBlockContent != "" {
				codeBlockContent += "\n"
			}
			codeBlockContent += line
			continue
		}

		// Empty line
		if trimmed == "" {
			continue
		}

		span := RichTextSpan{}

		// Headings
		if strings.HasPrefix(trimmed, "# ") {
			span.Heading = 1
			span.Text = strings.TrimPrefix(trimmed, "# ")
		} else if strings.HasPrefix(trimmed, "## ") {
			span.Heading = 2
			span.Text = strings.TrimPrefix(trimmed, "## ")
		} else if strings.HasPrefix(trimmed, "### ") {
			span.Heading = 3
			span.Text = strings.TrimPrefix(trimmed, "### ")
		} else if strings.HasPrefix(trimmed, "#### ") {
			span.Heading = 4
			span.Text = strings.TrimPrefix(trimmed, "#### ")
		} else if strings.HasPrefix(trimmed, "##### ") {
			span.Heading = 5
			span.Text = strings.TrimPrefix(trimmed, "##### ")
		} else if strings.HasPrefix(trimmed, "###### ") {
			span.Heading = 6
			span.Text = strings.TrimPrefix(trimmed, "###### ")
		} else if strings.HasPrefix(trimmed, "> ") {
			// Blockquote
			span.Blockquote = true
			span.Text = strings.TrimPrefix(trimmed, "> ")
		} else if strings.HasPrefix(trimmed, "- ") || strings.HasPrefix(trimmed, "* ") {
			// Unordered list
			span.Bullet = true
			span.Text = strings.TrimPrefix(strings.TrimPrefix(trimmed, "- "), "* ")
		} else if match := matchOrderedList(trimmed); match > 0 {
			// Ordered list
			span.Numbered = match
			span.Text = trimmed[strings.Index(trimmed, ". ")+2:]
		} else {
			span.Text = trimmed
		}

		// Inline formatting
		span.Text = parseInlineMarkdown(span.Text, &span)

		spans = append(spans, span)
	}

	return spans
}

func matchOrderedList(s string) int {
	parts := strings.SplitN(s, ". ", 2)
	if len(parts) == 2 {
		if num, err := strconv.Atoi(parts[0]); err == nil && num > 0 {
			return num
		}
	}
	return 0
}

// parseInlineMarkdown handles bold, italic, inline code, and links
func parseInlineMarkdown(text string, base *RichTextSpan) string {
	// For simplicity, we strip inline markers and set flags on the base span
	// A full implementation would return sub-spans, but this covers 90% of cases

	// Inline code `code`
	if strings.HasPrefix(text, "`") && strings.HasSuffix(text, "`") && len(text) > 2 {
		base.Code = true
		return text[1 : len(text)-1]
	}

	// Bold **text** or __text__
	text = parseBoldItalic(text, base)

	// Links [text](url)
	text = parseLinks(text, base)

	return text
}

func parseBoldItalic(text string, span *RichTextSpan) string {
	// **bold**
	if strings.Contains(text, "**") {
		parts := strings.Split(text, "**")
		if len(parts) >= 3 {
			span.Bold = true
			// Return first bold segment for simplicity
			return parts[1]
		}
	}
	// __bold__
	if strings.Contains(text, "__") {
		parts := strings.Split(text, "__")
		if len(parts) >= 3 {
			span.Bold = true
			return parts[1]
		}
	}
	// *italic*
	if strings.Contains(text, "*") && !strings.Contains(text, "**") {
		parts := strings.Split(text, "*")
		if len(parts) >= 3 {
			span.Italic = true
			return parts[1]
		}
	}
	// _italic_
	if strings.Contains(text, "_") && !strings.Contains(text, "__") {
		parts := strings.Split(text, "_")
		if len(parts) >= 3 {
			span.Italic = true
			return parts[1]
		}
	}
	return text
}

func parseLinks(text string, span *RichTextSpan) string {
	// Simple [text](url) extraction
	if idx := strings.Index(text, "["); idx != -1 {
		if endIdx := strings.Index(text[idx:], "]"); endIdx != -1 {
			if urlStart := strings.Index(text[idx+endIdx:], "("); urlStart != -1 {
				if urlEnd := strings.Index(text[idx+endIdx+urlStart:], ")"); urlEnd != -1 {
					linkText := text[idx+1 : idx+endIdx]
					url := text[idx+endIdx+urlStart+1 : idx+endIdx+urlStart+urlEnd]
					span.LinkURL = url
					return linkText
				}
			}
		}
	}
	return text
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
	// Pin the main goroutine to its OS thread for the lifetime of the process.
	// COM (CoInitializeEx), Direct2D, WASAPI, and the Win32 message loop all
	// require every call to happen on the exact same OS thread. Without this,
	// Go's scheduler is free to migrate the goroutine between threads, which
	// causes silent COM failures, hangs in GetMessage, and random crashes.
	runtime.LockOSThread()

	assemblyAIKey = os.Getenv("ASSEMBLYAI_API_KEY")
	geminiKey = os.Getenv("GEMINI_API_KEY")
	fmt.Println(assemblyAIKey, geminiKey)
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

	procSetLayeredWindowAttributes.Call(hwnd, 0, 255, LWA_ALPHA) // window is fully opaque; bg rect controls glass tint

	// Enable DWM blur-behind so transparent D2D pixels show a blurred desktop
	// (gives the real frosted-glass appearance controlled by the slider)
	blur := DWM_BLURBEHIND{DwFlags: 3, FEnable: 1}
	procDwmEnableBlurBehindWindow.Call(hwnd, uintptr(unsafe.Pointer(&blur)))

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
