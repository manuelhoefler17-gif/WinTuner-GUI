Update rules now apply per tenant and package ID. Packages can be excluded from Update All, pinned to an exact target or held at the installed version, with an optional reason and expiry. Exclusions leave explicitly checked updates available; pins apply to both update actions and do not initiate downgrades.

The Updates tab provides a rule editor, eligibility display and rule/skipped-package details in deployment preview. The worker verifies pinned targets are available in WinGet and rereads current rules before preview/build and before upload. Changed rules that alter eligibility or the approved target block upload.

Rules are stored separately from credentials with tenant/schema validation, serialized atomic writes and fail-closed loading. Expired rules are ignored, and idle UI eligibility refreshes every 30 seconds. The rules govern the Updates workflow, not manual package uploads or Discovery deployment.

Validation:
- User confirmed the manual GUI rule workflow worked.
- All 53 PowerShell files passed syntax validation; all 204 Pester tests passed; git diff --check passed.
- Automated coverage includes tenant separation, malformed data, expiry, exact pins/holds, action state and actual worker behavior, including mid-build rule changes. Graph calls are simulated and perform no tenant writes.

Release: v0.10.26.
