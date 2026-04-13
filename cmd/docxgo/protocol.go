// Package main implements the docxgo CLI binary with JSON-RPC over stdin/stdout.
package main

import "encoding/json"

// Request represents a JSON-RPC request.
type Request struct {
	ID     any     `json:"id"`
	Method string          `json:"method"`
	Params json.RawMessage `json:"params,omitempty"`
}

// Response represents a JSON-RPC response.
type Response struct {
	ID     any `json:"id"`
	Result any `json:"result,omitempty"`
	Error  *RPCError   `json:"error,omitempty"`
}

// RPCError represents an error in a JSON-RPC response.
type RPCError struct {
	Code      string `json:"code"`
	Message   string `json:"message"`
	Operation string `json:"operation,omitempty"`
}

// errorResponse is a convenience helper that builds an error Response.
func errorResponse(id any, code, message, operation string) Response {
	return Response{
		ID: id,
		Error: &RPCError{
			Code:      code,
			Message:   message,
			Operation: operation,
		},
	}
}
