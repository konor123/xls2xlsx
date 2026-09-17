//go:build windows

package main

import (
	"archive/zip"
	"bytes"
	_ "embed"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"syscall"
	"time"
	"unsafe"
)

//go:embed payload.zip
var payloadZip []byte

var (
	user32   = syscall.NewLazyDLL("user32.dll")
	kernel32 = syscall.NewLazyDLL("kernel32.dll")
	comdlg32 = syscall.NewLazyDLL("comdlg32.dll")
	shell32  = syscall.NewLazyDLL("shell32.dll")
	gdi32    = syscall.NewLazyDLL("gdi32.dll")

	procRegisterClassExW = user32.NewProc("RegisterClassExW")
	procCreateWindowExW  = user32.NewProc("CreateWindowExW")
	procDefWindowProcW   = user32.NewProc("DefWindowProcW")
	procShowWindow       = user32.NewProc("ShowWindow")
	procUpdateWindow     = user32.NewProc("UpdateWindow")
	procGetMessageW      = user32.NewProc("GetMessageW")
	procTranslateMessage = user32.NewProc("TranslateMessage")
	procDispatchMessageW = user32.NewProc("DispatchMessageW")
	procPostQuitMessage  = user32.NewProc("PostQuitMessage")
	procSendMessageW     = user32.NewProc("SendMessageW")
	procSetWindowTextW   = user32.NewProc("SetWindowTextW")
	procMessageBoxW      = user32.NewProc("MessageBoxW")
	procEnableWindow     = user32.NewProc("EnableWindow")
	procGetModuleHandleW = kernel32.NewProc("GetModuleHandleW")
	procGetOpenFileNameW = comdlg32.NewProc("GetOpenFileNameW")
	procDragAcceptFiles  = shell32.NewProc("DragAcceptFiles")
	procDragQueryFileW   = shell32.NewProc("DragQueryFileW")
	procDragFinish       = shell32.NewProc("DragFinish")
	procCreateFontW      = gdi32.NewProc("CreateFontW")
)

const (
	WS_OVERLAPPEDWINDOW    = 0x00CF0000
	WS_VISIBLE             = 0x10000000
	WS_CHILD               = 0x40000000
	WS_BORDER              = 0x00800000
	WS_VSCROLL             = 0x00200000
	WS_HSCROLL             = 0x00100000
	WS_TABSTOP             = 0x00010000
	LBS_NOTIFY             = 0x0001
	CW_USEDEFAULT          = 0x80000000
	SW_SHOW                = 5
	WM_DESTROY             = 0x0002
	WM_COMMAND             = 0x0111
	WM_DROPFILES           = 0x0233
	BN_CLICKED             = 0
	LB_ADDSTRING           = 0x0180
	LB_RESETCONTENT        = 0x0184
	LB_SETHORIZONTALEXTENT = 0x0194
	WM_SETFONT             = 0x0030
	MB_OK                  = 0x00000000
	MB_YESNO               = 0x00000004
	MB_ICONERROR           = 0x00000010
	MB_ICONQUESTION        = 0x00000020
	MB_ICONINFORMATION     = 0x00000040
	IDYES                  = 6
	OFN_EXPLORER           = 0x00080000
	OFN_FILEMUSTEXIST      = 0x00001000
	OFN_ALLOWMULTISELECT   = 0x00000200
	IDC_ADD                = 1001
	IDC_CLEAR              = 1002
	SS_ENDELLIPSIS          = 0x00004000
	IDC_CONVERT            = 1003
	WM_APP_STATUS          = 0x8001
	WM_APP_DONE            = 0x8002
)

type WNDCLASSEX struct {
	CbSize        uint32
	Style         uint32
	LpfnWndProc   uintptr
	CbClsExtra    int32
	CbWndExtra    int32
	HInstance     uintptr
	HIcon         uintptr
	HCursor       uintptr
	HbrBackground uintptr
	LpszMenuName  *uint16
	LpszClassName *uint16
	HIconSm       uintptr
}
type POINT struct{ X, Y int32 }
type MSG struct {
	Hwnd           uintptr
	Message        uint32
	WParam, LParam uintptr
	Time           uint32
	Pt             POINT
}
type OPENFILENAME struct {
	LStructSize                    uint32
	HwndOwner, HInstance           uintptr
	LpstrFilter, LpstrCustomFilter *uint16
	NMaxCustFilter, NFilterIndex   uint32
	LpstrFile                      *uint16
	NMaxFile                       uint32
	LpstrFileTitle                 *uint16
	NMaxFileTitle                  uint32
	LpstrInitialDir, LpstrTitle    *uint16
	Flags                          uint32
	NFileOffset, NFileExtension    uint16
	LpstrDefExt                    *uint16
	LCustData, LpfnHook            uintptr
	LpTemplateName                 *uint16
	PvReserved                     uintptr
	DwReserved, FlagsEx            uint32
}

var (
	hwndMain, hwndList, hwndStatus, hwndConvert, hwndAdd, hwndClear uintptr
	files                                                           []string
	guiFont                                                         uintptr
)

func wstr(s string) *uint16       { p, _ := syscall.UTF16PtrFromString(s); return p }
func loword(v uintptr) uint16     { return uint16(v & 0xffff) }
func hiword(v uintptr) uint16     { return uint16((v >> 16) & 0xffff) }
func setText(h uintptr, s string) { procSetWindowTextW.Call(h, uintptr(unsafe.Pointer(wstr(s)))) }
func messageBox(text, title string, flags uintptr) int {
	r, _, _ := procMessageBoxW.Call(hwndMain, uintptr(unsafe.Pointer(wstr(text))), uintptr(unsafe.Pointer(wstr(title))), flags)
	return int(r)
}

func addToList(path string) {
	abs, err := filepath.Abs(path)
	if err != nil {
		abs = path
	}
	if strings.ToLower(filepath.Ext(abs)) != ".xls" {
		return
	}
	for _, f := range files {
		if strings.EqualFold(f, abs) {
			return
		}
	}
	files = append(files, abs)
	procSendMessageW.Call(hwndList, LB_ADDSTRING, 0, uintptr(unsafe.Pointer(wstr(abs))))
	setText(hwndStatus, fmt.Sprintf("Files: %d", len(files)))
}
func clearList() {
	files = nil
	procSendMessageW.Call(hwndList, LB_RESETCONTENT, 0, 0)
	setText(hwndStatus, "Add .xls files or drag and drop them into this window.")
}
func makeMultiSZ(parts ...string) []uint16 {
	var out []uint16
	for _, p := range parts {
		u, _ := syscall.UTF16FromString(p)
		out = append(out, u...)
	}
	return append(out, 0)
}
func splitMultiSZ(buf []uint16) []string {
	var out []string
	start := 0
	for i := 0; i < len(buf); i++ {
		if buf[i] == 0 {
			if i == start {
				break
			}
			out = append(out, syscall.UTF16ToString(buf[start:i]))
			start = i + 1
		}
	}
	return out
}
func ptrUTF16String(p uintptr) string {
	if p == 0 {
		return ""
	}
	u := make([]uint16, 0, 1024)
	for i := uintptr(0); i < 32768; i++ {
		ch := *(*uint16)(unsafe.Pointer(p + i*2))
		if ch == 0 {
			break
		}
		u = append(u, ch)
	}
	return syscall.UTF16ToString(u)
}
func pickFiles() {
	buf := make([]uint16, 65536)
	filter := makeMultiSZ("Excel 97-2003 (*.xls)", "*.xls", "All files (*.*)", "*.*")
	title, _ := syscall.UTF16FromString("Select XLS files to convert")
	ofn := OPENFILENAME{LStructSize: uint32(unsafe.Sizeof(OPENFILENAME{})), HwndOwner: hwndMain, LpstrFilter: &filter[0], NFilterIndex: 1, LpstrFile: &buf[0], NMaxFile: uint32(len(buf)), LpstrTitle: &title[0], Flags: OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT}
	r, _, _ := procGetOpenFileNameW.Call(uintptr(unsafe.Pointer(&ofn)))
	if r == 0 {
		return
	}
	parts := splitMultiSZ(buf)
	if len(parts) == 1 {
		addToList(parts[0])
		return
	}
	if len(parts) > 1 {
		for _, n := range parts[1:] {
			addToList(filepath.Join(parts[0], n))
		}
	}
}

func findSoffice() string {
	if p, err := exec.LookPath("soffice.exe"); err == nil {
		return p
	}
	cands := []string{filepath.Join(os.Getenv("ProgramFiles"), "LibreOffice", "program", "soffice.exe"), filepath.Join(os.Getenv("ProgramFiles(x86)"), "LibreOffice", "program", "soffice.exe"), filepath.Join(os.Getenv("LOCALAPPDATA"), "Programs", "LibreOffice", "program", "soffice.exe")}
	for _, p := range cands {
		if p != "" {
			if st, err := os.Stat(p); err == nil && !st.IsDir() {
				return p
			}
		}
	}
	return ""
}
func uniqueTarget(src string) string {
	dir := filepath.Dir(src)
	base := strings.TrimSuffix(filepath.Base(src), filepath.Ext(src))
	out := filepath.Join(dir, base+".xlsx")
	if _, err := os.Stat(out); os.IsNotExist(err) {
		return out
	}
	for i := 1; ; i++ {
		s := "_converted"
		if i > 1 {
			s = fmt.Sprintf("_converted_%d", i)
		}
		p := filepath.Join(dir, base+s+".xlsx")
		if _, err := os.Stat(p); os.IsNotExist(err) {
			return p
		}
	}
}

func convertWithLibreOffice(soffice, src, target string) error {
	tmp, err := os.MkdirTemp("", "xls2xlsx-lo-")
	if err != nil {
		return err
	}
	defer os.RemoveAll(tmp)
	profile := filepath.Join(tmp, "profile")
	outdir := filepath.Join(tmp, "out")
	os.MkdirAll(outdir, 0755)
	profileURL := "file:///" + strings.ReplaceAll(filepath.ToSlash(profile), " ", "%20")
	args := []string{"-env:UserInstallation=" + profileURL, "--headless", "--nologo", "--nodefault", "--nolockcheck", "--norestore", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", outdir, src}
	cmd := exec.Command(soffice, args...)
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true}
	out, err := cmd.CombinedOutput()
	if err != nil {
		return fmt.Errorf("LibreOffice failed: %v (%s)", err, strings.TrimSpace(string(out)))
	}
	gen := filepath.Join(outdir, strings.TrimSuffix(filepath.Base(src), filepath.Ext(src))+".xlsx")
	if _, err := os.Stat(gen); err != nil {
		return fmt.Errorf("LibreOffice did not create an XLSX file: %s", strings.TrimSpace(string(out)))
	}
	if err = os.Rename(gen, target); err != nil {
		b, e := os.ReadFile(gen)
		if e != nil {
			return e
		}
		return os.WriteFile(target, b, 0644)
	}
	return nil
}

func extractPayload() (string, error) {
	dir, err := os.MkdirTemp("", "xls2xlsx-parser-")
	if err != nil {
		return "", err
	}
	zr, err := zip.NewReader(bytes.NewReader(payloadZip), int64(len(payloadZip)))
	if err != nil {
		os.RemoveAll(dir)
		return "", err
	}
	for _, f := range zr.File {
		p := filepath.Join(dir, filepath.FromSlash(f.Name))
		clean := filepath.Clean(p)
		if !strings.HasPrefix(strings.ToLower(clean), strings.ToLower(filepath.Clean(dir)+string(os.PathSeparator))) {
			continue
		}
		if f.FileInfo().IsDir() {
			os.MkdirAll(clean, 0755)
			continue
		}
		os.MkdirAll(filepath.Dir(clean), 0755)
		r, e := f.Open()
		if e != nil {
			os.RemoveAll(dir)
			return "", e
		}
		w, e := os.Create(clean)
		if e == nil {
			_, e = io.Copy(w, r)
			w.Close()
		}
		r.Close()
		if e != nil {
			os.RemoveAll(dir)
			return "", e
		}
	}
	return dir, nil
}
func findPython() string {
	cands := []string{}
	if p, err := exec.LookPath("python.exe"); err == nil {
		cands = append(cands, p)
	}
	if p, err := exec.LookPath("py.exe"); err == nil { // use py.exe via wrapper marker
		return p
	}
	la := os.Getenv("LOCALAPPDATA")
	pf := os.Getenv("ProgramFiles")
	for _, v := range []string{"313", "312", "311", "310"} {
		cands = append(cands, filepath.Join(la, "Programs", "Python", "Python"+v, "python.exe"), filepath.Join(pf, "Python"+v, "python.exe"))
	}
	for _, p := range cands {
		if st, err := os.Stat(p); err == nil && !st.IsDir() {
			return p
		}
	}
	return ""
}
func ensurePython() string {
	if p := findPython(); p != "" {
		return p
	}
	ans := messageBox("Python is required for recovery mode but was not found on this PC.\r\n\r\nInstall Python 3.12 with winget?", "Recovery parser setup", MB_YESNO|MB_ICONQUESTION)
	if ans != IDYES {
		return ""
	}
	wg, err := exec.LookPath("winget.exe")
	if err != nil {
		messageBox("winget was not found. Install Python manually and try again.", "Python setup failed", MB_OK|MB_ICONERROR)
		return ""
	}
	setText(hwndStatus, "Installing Python...")
	cmd := exec.Command(wg, "install", "--id", "Python.Python.3.12", "-e", "--silent", "--accept-source-agreements", "--accept-package-agreements")
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true}
	if err := cmd.Run(); err != nil {
		return ""
	}
	time.Sleep(2 * time.Second)
	return findPython()
}
func runPython(python string, args ...string) *exec.Cmd {
	if strings.EqualFold(filepath.Base(python), "py.exe") {
		args = append([]string{"-3"}, args...)
	}
	cmd := exec.Command(python, args...)
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true}
	return cmd
}
func isOle(path string) bool {
	f, err := os.Open(path)
	if err != nil {
		return false
	}
	defer f.Close()
	b := make([]byte, 8)
	if _, err = io.ReadFull(f, b); err != nil {
		return false
	}
	sig := []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}
	return bytes.Equal(b, sig)
}
func ensureXlrd(python, payloadDir string) error {
	if !isOle(currentInput) {
		return nil
	}
	runtimeDir := filepath.Join(payloadDir, "runtime_packages")
	code := fmt.Sprintf("import sys;sys.path.insert(0,r'%s');import xlrd", strings.ReplaceAll(runtimeDir, "'", "\\'"))
	if err := runPython(python, "-c", code).Run(); err == nil {
		return nil
	}
	setText(hwndStatus, "Preparing recovery module...")
	cmd := runPython(python, "-m", "pip", "install", "--disable-pip-version-check", "--no-warn-script-location", "--target", runtimeDir, "xlrd==2.0.2")
	out, err := cmd.CombinedOutput()
	if err != nil {
		return fmt.Errorf("Failed to install xlrd: %v (%s)", err, strings.TrimSpace(string(out)))
	}
	return nil
}

var currentInput string

func convertInternal(src, target string) error {
	currentInput = src
	python := ensurePython()
	if python == "" {
		return fmt.Errorf("Python is unavailable")
	}
	dir, err := extractPayload()
	if err != nil {
		return err
	}
	defer os.RemoveAll(dir)
	if err := ensureXlrd(python, dir); err != nil {
		return err
	}
	env := append(os.Environ(), "PYTHONPATH="+filepath.Join(dir, "runtime_packages"))
	cmd := runPython(python, filepath.Join(dir, "convert_worker.py"), src, target)
	cmd.Env = env
	out, err := cmd.CombinedOutput()
	if err != nil {
		return fmt.Errorf("Recovery failed: %v (%s)", err, strings.TrimSpace(string(out)))
	}
	if _, err := os.Stat(target); err != nil {
		return fmt.Errorf("Recovery did not create an output file")
	}
	return nil
}
func convertOne(src string) (string, string, error) {
	target := uniqueTarget(src)
	if lo := findSoffice(); lo != "" {
		if err := convertWithLibreOffice(lo, src, target); err == nil {
			return target, "LibreOffice", nil
		}
	}
	if err := convertInternal(src, target); err != nil {
		return "", "", err
	}
	return target, "Built-in parser", nil
}

func startConversion() {
	if len(files) == 0 {
		messageBox("Add at least one XLS file first.", "XLS to XLSX", MB_OK|MB_ICONINFORMATION)
		return
	}
	procEnableWindow.Call(hwndConvert, 0)
	procEnableWindow.Call(hwndAdd, 0)
	procEnableWindow.Call(hwndClear, 0)
	local := append([]string(nil), files...)
	go func() {
		ok := 0
		errs := []string{}
		for i, f := range local {
			postStatus(fmt.Sprintf("Converting %d/%d: %s", i+1, len(local), filepath.Base(f)))
			_, _, err := convertOne(f)
			if err != nil {
				errs = append(errs, filepath.Base(f)+": "+err.Error())
			} else {
				ok++
			}
		}
		summary := fmt.Sprintf("Completed: %d succeeded / %d failed", ok, len(errs))
		if len(errs) > 0 {
			if len(errs) > 5 {
				errs = errs[:5]
			}
			summary += "\r\n\r\n" + strings.Join(errs, "\r\n")
		}
		postDone(summary, len(errs) == 0)
	}()
}
func postStatus(s string) {
	p := wstr(s)
	procSendMessageW.Call(hwndMain, WM_APP_STATUS, 0, uintptr(unsafe.Pointer(p)))
	runtime.KeepAlive(p)
}
func postDone(s string, success bool) {
	p := wstr(s)
	var wp uintptr
	if success {
		wp = 1
	}
	procSendMessageW.Call(hwndMain, WM_APP_DONE, wp, uintptr(unsafe.Pointer(p)))
	runtime.KeepAlive(p)
}
func wndProc(hwnd uintptr, msg uint32, wparam, lparam uintptr) uintptr {
	switch msg {
	case WM_COMMAND:
		id := int(loword(wparam))
		code := hiword(wparam)
		if code == BN_CLICKED {
			switch id {
			case IDC_ADD:
				pickFiles()
			case IDC_CLEAR:
				clearList()
			case IDC_CONVERT:
				startConversion()
			}
		}
		return 0
	case WM_DROPFILES:
		h := wparam
		cnt, _, _ := procDragQueryFileW.Call(h, 0xFFFFFFFF, 0, 0)
		for i := uintptr(0); i < cnt; i++ {
			n, _, _ := procDragQueryFileW.Call(h, i, 0, 0)
			buf := make([]uint16, n+1)
			procDragQueryFileW.Call(h, i, uintptr(unsafe.Pointer(&buf[0])), n+1)
			addToList(syscall.UTF16ToString(buf))
		}
		procDragFinish.Call(h)
		return 0
	case WM_APP_STATUS:
		setText(hwndStatus, ptrUTF16String(lparam))
		return 0
	case WM_APP_DONE:
		text := ptrUTF16String(lparam)
		setText(hwndStatus, strings.Split(text, "\r\n")[0])
		procEnableWindow.Call(hwndConvert, 1)
		procEnableWindow.Call(hwndAdd, 1)
		procEnableWindow.Call(hwndClear, 1)
		icon := uintptr(MB_ICONINFORMATION)
		if wparam == 0 {
			icon = MB_ICONERROR
		}
		messageBox(text, "Conversion result", MB_OK|icon)
		return 0
	case WM_DESTROY:
		procPostQuitMessage.Call(0)
		return 0
	}
	r, _, _ := procDefWindowProcW.Call(hwnd, uintptr(msg), wparam, lparam)
	return r
}
func createControl(class, text string, style uint32, x, y, w, h int32, id int) uintptr {
	hinst, _, _ := procGetModuleHandleW.Call(0)
	hwnd, _, _ := procCreateWindowExW.Call(0, uintptr(unsafe.Pointer(wstr(class))), uintptr(unsafe.Pointer(wstr(text))), uintptr(style), uintptr(x), uintptr(y), uintptr(w), uintptr(h), hwndMain, uintptr(id), hinst, 0)
	if guiFont != 0 {
		procSendMessageW.Call(hwnd, WM_SETFONT, guiFont, 1)
	}
	return hwnd
}
func main() {
	runtime.LockOSThread()
	hinst, _, _ := procGetModuleHandleW.Call(0)
	// Segoe UI, 9 pt at the traditional 96-DPI logical scale.
	fontHeight := int32(-12)
	guiFont, _, _ = procCreateFontW.Call(uintptr(fontHeight), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 5, 0, uintptr(unsafe.Pointer(wstr("Segoe UI"))))
	className := wstr("XLS2XLSXRecoveryWindow3")
	wc := WNDCLASSEX{CbSize: uint32(unsafe.Sizeof(WNDCLASSEX{})), LpfnWndProc: syscall.NewCallback(wndProc), HInstance: hinst, HbrBackground: 6, LpszClassName: className}
	procRegisterClassExW.Call(uintptr(unsafe.Pointer(&wc)))
	hwndMain, _, _ = procCreateWindowExW.Call(0, uintptr(unsafe.Pointer(className)), uintptr(unsafe.Pointer(wstr("XLS to XLSX Recovery Converter"))), WS_OVERLAPPEDWINDOW|WS_VISIBLE, CW_USEDEFAULT, CW_USEDEFAULT, 940, 640, 0, 0, hinst, 0)
	if hwndMain == 0 {
		return
	}

	// Engine details belong in source comments / documentation, not in the GUI.
	hwndAdd = createControl("BUTTON", "Add files", WS_CHILD|WS_VISIBLE|WS_TABSTOP, 20, 24, 115, 34, IDC_ADD)
	hwndClear = createControl("BUTTON", "Clear list", WS_CHILD|WS_VISIBLE|WS_TABSTOP, 145, 24, 115, 34, IDC_CLEAR)
	hwndConvert = createControl("BUTTON", "Convert", WS_CHILD|WS_VISIBLE|WS_TABSTOP, 785, 24, 120, 34, IDC_CONVERT)

	hwndList = createControl("LISTBOX", "", WS_CHILD|WS_VISIBLE|WS_BORDER|WS_VSCROLL|WS_HSCROLL|LBS_NOTIFY, 20, 72, 885, 405, 2001)
	procSendMessageW.Call(hwndList, LB_SETHORIZONTALEXTENT, 2200, 0)

	hwndStatus = createControl("STATIC", "Add .xls files or drag and drop them into this window.", WS_CHILD|WS_VISIBLE|SS_ENDELLIPSIS, 20, 494, 885, 24, 2002)

	procDragAcceptFiles.Call(hwndMain, 1)
	procShowWindow.Call(hwndMain, SW_SHOW)
	procUpdateWindow.Call(hwndMain)
	var msg MSG
	for {
		r, _, _ := procGetMessageW.Call(uintptr(unsafe.Pointer(&msg)), 0, 0, 0)
		if int32(r) <= 0 {
			break
		}
		procTranslateMessage.Call(uintptr(unsafe.Pointer(&msg)))
		procDispatchMessageW.Call(uintptr(unsafe.Pointer(&msg)))
	}
}
