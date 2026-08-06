package oracle

import "errors"

const oracleAlreadyRunningMessage = "Another local VBE oracle process is already running"

var errOracleAlreadyRunning = errors.New("another local VBE oracle process is already running")
