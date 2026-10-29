When an existing Intune Win32 app is updated, copied Available assignments on the newly deployed app now automatically enable AutoUpdate superseded apps. Existing assignment settings are preserved, and Required, Uninstall and other intents are skipped. Configuration warnings are logged separately from deployment success.

Validation:
- Manual tenant test: Greenshot 1.3.315 -> 1.3.328; the new app's Available assignment shows AutoUpdate = Yes.
- Transient HTTP 429 failed safely during package creation; manual retry succeeded.
- PowerShell syntax, complete Pester suite and git diff --check.

Release: v0.10.23. Unified interactive authentication is deferred to the next change.
