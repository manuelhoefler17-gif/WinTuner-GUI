App updates now show a read-only deployment preview with tenant, versions, target group IDs, assignment intents, exclusions/filters and planned AutoUpdate settings. Package creation starts after approval, and the existing Intune version and assignments are rechecked before upload. Unapproved fallback versions and changed assignments block upload.

Completed app-update results are persisted in tenant-separated local history with a Settings viewer and CSV export. History stores a whitelist of operational fields; exports neutralize spreadsheet formulas. A history-write failure does not repeat a successful deployment.

Transient package-build and direct delegated Graph failures have bounded retries that honor Retry-After when available. Complete deployments are not automatically replayed. The preview reads displayVersion through the Intune beta model; assignment requests remain on v1.0.

Validation:
- Manual GUI preview showed Available, Required and no-assignment cases.
- Real 7-Zip update: 26.03 -> 26.05; history and CSV export confirmed Updated with the expected versions.
- All 51 PowerShell files passed syntax validation; all 190 Pester tests passed; git diff --check passed.
- Automated worker tests simulate Graph responses and perform no tenant writes.

Release: v0.10.25.
