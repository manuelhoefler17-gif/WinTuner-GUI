Background package retries now appear once and include the triggering HTTP status, such as HTTP 429 (Too Many Requests).

When HTTP 429 supplies no usable Retry-After header, fallback waits are now 15 and 30 seconds instead of 3 and 5 seconds. Valid server-specified delays remain authoritative. The three-attempt cap, 120-second total retry-wait budget, cancellation, permanent-error handling and protection against replaying complete deployments remain in place. External service rate limits may still cause a final failure.

Validation:
- All 56 PowerShell files parsed without errors; all 226 Pester tests passed; git diff --check passed.
- Regression tests exercise the actual package-build/batch/JSON/GUI-log path with a simulated 429, confirm one retry log entry and one deployment after recovery, and cover backoff, Retry-After precedence, budget limits and cancellation.
- User reported no recurrence in the final GUI test. A real 429 was not reproduced during that test; the retry path has automated coverage.

Release: v0.10.28.
