Interactive user login now requests the existing WinTuner and Discovery permissions together and uses the same WinTuner credential provider for direct Graph requests. This removes the separate Graph PowerShell sign-in while keeping token renewal with the upstream provider. The actual user and tenant are checked before session metadata is shared across background workflows; logout clears that metadata. Client-secret and certificate Graph SDK authentication retain their existing path.

Update scans now continue if writing a progress notification fails, preventing transient progress-file collisions from discarding valid scan results.

Validation:
- Manual GUI test: the user confirmed only one interactive sign-in.
- Automated coverage: identity rejection, tenant verification failure, provider token retrieval, PATCH settings, external URL rejection, runspace reuse, logout, app-only routing and progress reporting failures.
- All 45 PowerShell files passed syntax validation; all 170 Pester tests passed; git diff --check passed.
- GUI log confirms an update scan completed with 19 checked apps and 11 candidates.
- Long-lived token renewal has not yet been observed in a tenant session.

Release: v0.10.24.
