Save named customer tenant profiles for interactive user, client-secret or certificate authentication. Profiles retain the expected tenant GUID, identity and package folder, can be selected while signed out and idle, and are restored after restarting the GUI. The header shows the profile, tenant and authentication mode.

Login checks the actual tenant against the selected profile before Intune inventory reads or app-only permission probes. Switching clears the previous identity and protected secret; missing or malformed selected profiles block login instead of silently falling back. Editing authentication settings returns to manual configuration.

Profile storage uses schema validation, serialized atomic writes and optional copying of an existing Windows DPAPI-protected client secret. Profile files contain no plaintext secrets, tokens or certificate private keys. DPAPI protects the stored secret for the Windows user; it is not protection against malicious code already running as that user. Profile deletion does not revoke credentials in Entra. Settings can scroll in smaller windows so profile management remains accessible.

Validation:
- User confirmed the interactive profile workflow and restoration after a GUI restart.
- All 55 PowerShell files passed syntax validation; all 221 Pester tests passed; git diff --check passed.
- Automated coverage includes profile validation/storage, DPAPI-only secret persistence, credential isolation, startup restoration and actual login-worker rejection of a wrong tenant in all three authentication modes. App-only profile login was not explicitly confirmed in the final manual test.

Release: v0.10.27.
