//go:build windows

package workbookuse

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"slices"
	"strings"
	"syscall"
	"unsafe"

	"golang.org/x/sys/windows"
)

const (
	rmSessionKeyLength = 32
	rmMaxAppName       = 255
	rmMaxServiceName   = 63
)

type rmUniqueProcess struct {
	ProcessID        uint32
	ProcessStartTime windows.Filetime
}

type rmProcessInfo struct {
	Process          rmUniqueProcess
	AppName          [rmMaxAppName + 1]uint16
	ServiceShortName [rmMaxServiceName + 1]uint16
	ApplicationType  uint32
	AppStatus        uint32
	TSSessionID      uint32
	Restartable      int32
}

var (
	restartManagerDLL   = windows.NewLazySystemDLL("rstrtmgr.dll")
	rmStartSessionProc  = restartManagerDLL.NewProc("RmStartSession")
	rmRegisterFilesProc = restartManagerDLL.NewProc("RmRegisterResources")
	rmGetListProc       = restartManagerDLL.NewProc("RmGetList")
	rmEndSessionProc    = restartManagerDLL.NewProc("RmEndSession")

	startSession      = startRestartManagerSession
	registerWorkbook  = registerRestartManagerWorkbook
	queryProcesses    = getProcessList
	endSession        = endRestartManagerSession
	identifyExcel     = processIsExcel
	canonicalWorkbook = canonicalPath
)

func Detect(path string) (State, error) {
	canonical, err := canonicalWorkbook(path)
	if err != nil {
		return State{}, err
	}
	handle, err := startSession()
	if err != nil {
		return State{}, err
	}
	defer endSession(handle)
	if err := registerWorkbook(handle, canonical); err != nil {
		return State{}, err
	}
	processes, err := queryProcesses(handle)
	if err != nil {
		return State{}, err
	}
	pids := make([]uint32, 0, len(processes))
	for _, process := range processes {
		isExcel, identifyErr := identifyExcel(process.Process.ProcessID)
		if identifyErr != nil {
			return State{}, fmt.Errorf("identify Restart Manager process %d: %w", process.Process.ProcessID, identifyErr)
		}
		if isExcel {
			pids = append(pids, process.Process.ProcessID)
		}
	}
	slices.Sort(pids)
	pids = slices.Compact(pids)
	return State{OpenInExcel: len(pids) > 0, ExcelPIDs: pids}, nil
}

func startRestartManagerSession() (uint32, error) {
	var handle uint32
	key := make([]uint16, rmSessionKeyLength+1)
	if code, _, callErr := rmStartSessionProc.Call(uintptr(unsafe.Pointer(&handle)), 0, uintptr(unsafe.Pointer(&key[0]))); code != 0 {
		return 0, windowsCallError("RmStartSession", code, callErr)
	}
	return handle, nil
}

func registerRestartManagerWorkbook(handle uint32, path string) error {
	path16, err := windows.UTF16PtrFromString(path)
	if err != nil {
		return fmt.Errorf("encode workbook path for Restart Manager: %w", err)
	}
	files := []*uint16{path16}
	if code, _, callErr := rmRegisterFilesProc.Call(
		uintptr(handle), 1, uintptr(unsafe.Pointer(&files[0])), 0, 0, 0, 0,
	); code != 0 {
		return windowsCallError("RmRegisterResources", code, callErr)
	}
	return nil
}

func endRestartManagerSession(handle uint32) {
	_, _, _ = rmEndSessionProc.Call(uintptr(handle))
}

func canonicalPath(path string) (string, error) {
	abs, err := filepath.Abs(path)
	if err != nil {
		return "", fmt.Errorf("resolve workbook path: %w", err)
	}
	if _, err := os.Stat(abs); err != nil {
		return "", err
	}
	resolved, err := filepath.EvalSymlinks(abs)
	if err != nil {
		return "", fmt.Errorf("resolve canonical workbook path: %w", err)
	}
	return filepath.Clean(resolved), nil
}

func getProcessList(handle uint32) ([]rmProcessInfo, error) {
	var needed uint32
	var count uint32
	var rebootReasons uint32
	code, _, callErr := rmGetListProc.Call(
		uintptr(handle), uintptr(unsafe.Pointer(&needed)), uintptr(unsafe.Pointer(&count)), 0,
		uintptr(unsafe.Pointer(&rebootReasons)),
	)
	if code == 0 && needed == 0 {
		return nil, nil
	}
	if code != uintptr(windows.ERROR_MORE_DATA) {
		return nil, windowsCallError("RmGetList", code, callErr)
	}
	for {
		processes := make([]rmProcessInfo, needed)
		count = needed
		code, _, callErr = rmGetListProc.Call(
			uintptr(handle), uintptr(unsafe.Pointer(&needed)), uintptr(unsafe.Pointer(&count)),
			uintptr(unsafe.Pointer(&processes[0])), uintptr(unsafe.Pointer(&rebootReasons)),
		)
		if code == 0 {
			return processes[:count], nil
		}
		if code != uintptr(windows.ERROR_MORE_DATA) {
			return nil, windowsCallError("RmGetList", code, callErr)
		}
	}
}

func processIsExcel(pid uint32) (bool, error) {
	process, err := windows.OpenProcess(windows.PROCESS_QUERY_LIMITED_INFORMATION, false, pid)
	if err != nil {
		if errors.Is(err, windows.ERROR_INVALID_PARAMETER) {
			return false, nil
		}
		return false, err
	}
	defer func() { _ = windows.CloseHandle(process) }()
	buffer := make([]uint16, windows.MAX_PATH)
	for {
		size := uint32(len(buffer))
		err = windows.QueryFullProcessImageName(process, 0, &buffer[0], &size)
		if err == nil {
			return strings.EqualFold(filepath.Base(windows.UTF16ToString(buffer[:size])), "EXCEL.EXE"), nil
		}
		if !errors.Is(err, windows.ERROR_INSUFFICIENT_BUFFER) {
			return false, err
		}
		buffer = make([]uint16, len(buffer)*2)
	}
}

func windowsCallError(name string, code uintptr, callErr error) error {
	if code == 0 {
		return nil
	}
	err := syscall.Errno(code)
	if callErr != nil && !errors.Is(callErr, syscall.Errno(0)) {
		return fmt.Errorf("%s: %w (%v)", name, err, callErr)
	}
	return fmt.Errorf("%s: %w", name, err)
}
