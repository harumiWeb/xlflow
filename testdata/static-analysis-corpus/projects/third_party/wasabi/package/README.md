# Wasabi Core Package

This directory contains **Wasabi.bas**, the monolithic, self contained module that introduces a complete, high performance networking stack to any Microsoft Office VBA or VB6 project. The architecture encompasses WebSocket (RFC 6455), raw TCP with optional TLS, and MQTT over WebSockets, functioning entirely without external dependencies or compiled DLLs.

> When imported, the VBA runtime compiles the module natively in memory. There is no build pipeline, no external binaries, and no complex packaging. The source code acts as the execution engine.

## Integration Guide

1. Open the VBA Integrated Development Environment (IDE).
2. Click **File > Import File...**
3. Select `Wasabi.bas` from this directory.
4. No additional configuration is required. The module relies strictly on native Win32 APIs, eliminating the need for external references or registry modifications.

After importing, the entire API surface becomes globally available.

> [!TIP]
> The comprehensive technical documentation is available in [`docs/API_REFERENCE.md`](../docs/API_REFERENCE.md).

## Architectural Capabilities

Wasabi exposes three independent operational layers that share a unified underlying transport and TLS engine.

**WebSocket Protocol (RFC 6455):** The primary interface. It manages the complete connection lifecycle, including secure handshakes, masking, payload fragmentation, ping/pong keepalives, and graceful closures. Advanced features include `permessage-deflate` compression, custom subprotocols, automatic MTU aware framing, and robust exponential backoff reconnections. Both text and binary frames are supported, featuring zero copy memory allocation for performance sensitive data streams.

**Raw TCP Sockets:** Provides direct access to the raw socket layer. This is essential when communicating with remote endpoints that utilize line protocols or proprietary binary formats outside the WebSocket specification. This layer shares the exact same SChannel TLS engine, supporting direct client certificate authentication via thumbprint, subject name, or disk loaded PFX files.

**MQTT Telemetry (v3.1.1):** Operates on top of an established WebSocket transport, implementing the standard MQTT packet exchange. Fully supports broker connection handshakes, QoS 0/1/2 publishing, topic subscriptions, and manual keep alive pinging for persistent IoT or messaging connections.

### Unified Transport Features

All three modes benefit from a shared foundational architecture:
* TLS 1.2 and TLS 1.3 negotiated natively via the Windows SChannel provider.
* HTTP and SOCKS proxy traversal, featuring stable `CryptStringToBinaryW` decoding for safe NTLM proxy authentication without data corruption.
* Dual stack IPv4 and IPv6 routing with configurable resolution preference.
* Microsecond precision latency telemetry and deep internal buffer diagnostics.
* Asynchronous, non blocking execution utilizing a subclassed hidden message window (`WSAAsyncSelect`), protected by `EbMode` guards to prevent host application crashes during state interruptions.

## Connection Handle Management

Every successful connection generates a unique numeric handle. Passing this handle to subsequent method calls targets that specific socket, enabling multiplexed concurrent streams (such as listening to multiple exchange feeds simultaneously). If operating a single socket architecture, handle arguments can be omitted, allowing Wasabi to route calls to the default active handle.

## Telemetry and Error Handling

The module maintains stateful error telemetry for every active handle. Developers can invoke `WasabiGetErrorDescription` to retrieve parsed, human readable diagnostic messages. For granular debugging, `WebSocketGetTechnicalDetails` extracts deep subsystem state variables, isolating low level Winsock or SChannel fault codes.
