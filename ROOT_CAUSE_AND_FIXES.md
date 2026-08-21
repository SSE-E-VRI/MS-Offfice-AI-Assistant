# Root Cause Analysis: Office 2021 Add-in Failure

**Status**: RESOLVED  
**Verified Date**: 2026-08-21  
**Severity**: CRITICAL (blocks all Office 2021 & newer users)

---

## Executive Summary

The add-in fails to load on Office 2021 **not due to version incompatibility**, but because of **poisoned COM class registration** left over from previous builds. The COM activation engine (`mscoree.dll`) selects the highest-versioned subkey under `InprocServer32\<version>`, which points to a stale `1.0.0.0` assembly identity that no longer exists in the current DLL. This causes **manifest-mismatch failure (HRESULT 0x80131040)** on activation.

---

## Root Cause

### The Mechanism

The COM CLSID `{2F8D4B61-…}` in `HKCU\Software\Classes` accumulated two version subkeys:

| Subkey | Assembly Identity | Status |
|---|---|---|
| `InprocServer32\0.4.0.0` | `MSOfficeAIAssistant, Version=0.4.0.0` | ✅ Current |
| `InprocServer32\1.0.0.0` | `MistralOfficeAddin, Version=1.0.0.0` | ❌ **Stale** |

The CLR COM activation process **always selects the highest version number** (1.0.0.0 > 0.4.0.0). Because this version is no longer present in the current assembly (`MSOfficeAIAssistant`), the loader throws:

```
FileLoadException: … manifest definition does not match the assembly reference … 0x80131040
```

### Why It Happens

The `install.ps1` script registers the COM classes via `RegAsm /regfile`, which **merges** rather than replaces registry entries. Stale `1.0.0.0` subkeys persist indefinitely and are subsequently poisoned by a loop that blindly stamps the current CodeBase onto all child keys.

### Why "2010 Works, 2021 Doesn't"

The issue is **not version-specific**; it is **registration-history-specific**. A clean machine with no prior builds has no stale subkey. A machine upgraded from pre-0.4.0 builds carries the orphan.

---

## Fixes Applied

### 1. Clean Stale Registration (CRITICAL)
**File**: `install.ps1` (lines 83–98)  
Delete entire CLSID trees before importing new regfile.

### 2. Selective CodeBase Stamping (DEFENSIVE)
**File**: `install.ps1` (lines 151–167)  
Only stamp CodeBase on parent and current version (0.4.0.0).

### 3. Office 2021 Version Detection (QUALITY)
**File**: `src/Core/VersionDetector.cs` (lines 52–75)  
Use second build number part to discriminate versions.

### 4. Installer Pre-Cleanup (DEFENSIVE)
**Files**: `installer/setup-x64.iss`, `installer/setup-x86.iss`  
Run `regasm /unregister` before `/codebase`.

---

## Testing & Validation

All fixes verified and documented in `FIXES_SUMMARY.txt`.
