# Error Handling RFC

This document defines the standardized error handling strategy for the PM Automation project. The goal is to create a consistent, explicit, and traceable approach to managing errors, reducing silent failures, and improving the overall resilience and maintainability of the codebase.

## 1. Error Taxonomy

Errors are classified into the following categories to ensure they are handled appropriately. All custom errors will extend a `BaseError` class that captures the stack trace and supports error chaining.

- **`ValidationError`**: Represents an error due to invalid input data, such as a missing required field or a malformed SFID. These errors are generally not retryable and should result in a fast failure.
- **`DependencyError`**: Occurs when an external service (e.g., Google Sheets, `LockService`) fails or returns an unexpected response. These errors may be transient and can be retried.
- **`TransientError`**: A sub-type of `DependencyError` that is known to be temporary, such as a lock timeout or a temporary network issue. These are ideal candidates for retries with exponential backoff.
- **`ConfigurationError`**: Represents an error in the application's configuration, such as a missing sheet name or an invalid column mapping. These errors should cause the application to fail fast at startup or when the configuration is loaded.
- **`BusinessLogicError`**: An unexpected error in the application's business logic that is not covered by the other categories. These should be treated as critical bugs and investigated immediately.

## 2. Error Propagation

- **Catch Specific Errors**: Avoid broad `try...catch (e)` blocks. Instead, catch specific error types (`ValidationError`, `TransientError`, etc.) to handle them appropriately.
- **Wrap and Rethrow**: When catching an error from a lower-level function, wrap it with additional context and rethrow it as a more specific error type. This preserves the original error while adding meaningful context for debugging.
- **Fail Fast**: For non-recoverable errors like `ValidationError` or `ConfigurationError`, the operation should be terminated immediately with a clear log message.

## 3. Structured Logging

All log entries, especially for errors, must be in a structured JSON format. This enables easier parsing, filtering, and analysis in a logging system.

The standard log format will be:

```json
{
  "timestamp": "2025-10-08T18:20:00.000Z",
  "severity": "ERROR",
  "message": "A brief, human-readable error message.",
  "correlationId": "unique-id-for-the-operation",
  "error": {
    "type": "TransientError",
    "message": "Lock not acquired after 3 retries.",
    "stack": "...",
    "cause": {
      "type": "Error",
      "message": "Lock timed out.",
      "stack": "..."
    }
  },
  "context": {
    "appName": "PM Automation",
    "functionName": "executeTransfer",
    "spreadsheetId": "...",
    "sheetName": "Upcoming",
    "row": 123
  }
}
```

- **`severity`**: Can be `DEBUG`, `INFO`, `WARN`, or `ERROR`.
- **`correlationId`**: A unique ID generated at the start of an operation (e.g., in the `onEdit` trigger) and passed through all subsequent function calls. This allows for easy tracing of an entire operation.

## 4. Operational Safeguards

- **Timeouts**: All I/O operations (e.g., calls to `SpreadsheetApp`, `UrlFetchApp`) must have explicit timeouts to prevent them from hanging indefinitely.
- **Retries with Exponential Backoff and Jitter**: For `TransientError`s, implement a retry mechanism with exponential backoff and jitter to handle temporary service disruptions gracefully. A central utility function will be created for this purpose.
- **Circuit Breaker**: For critical dependencies that are known to be flaky, a simple circuit breaker mechanism will be implemented to prevent repeated calls to a failing service.

## 5. User-Facing Messages

Error notifications sent to users (e.g., via email) should be clear, concise, and actionable. They should include:

- A summary of the error.
- The `correlationId` for support reference.
- A link to the affected spreadsheet.
- A brief explanation of the impact.

Secrets and PII must **never** be included in user-facing messages or logs.

## Implementation Plan

1. **Create `ErrorService.gs`**: This new service will centralize all error handling logic, including the custom error classes, the structured logger, and the retry utility.
2. **Refactor Core Components**: Update `TransferEngine.gs`, `Automations.gs`, and `Utilities.gs` to use the new `ErrorService`.
3. **Add Tests**: Create new tests to cover failure paths and ensure the new error handling works as expected.
4. **Update Documentation**: Update the main `README.md` to reflect the new error handling strategy.