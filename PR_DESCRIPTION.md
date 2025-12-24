# Fix family loader crashes when switching modes and PropertyChanged errors

## Summary
Fixes multiple critical crashes in the Family Loader feature that occurred when switching between local and cloud modes, and when handling PropertyChanged events.

## Issues Fixed

### 1. Race Condition Crash When Switching Modes
**Symptoms:** Application crashed when switching from Cloud to Local mode or vice versa
**Root Cause:** Multiple scan threads running simultaneously, competing for UI updates
**Fix:** Added thread cancellation mechanism with `_cancel_current_scan()` method

### 2. Cloud API Configuration Errors
**Symptoms:** Cryptic "HTTP 404" error when clicking Cloud mode
**Root Cause:** Cloud API URL still set to placeholder, not deployed
**Fix:**
- Detect placeholder URLs and show helpful setup instructions
- Automatically switch back to Local mode
- Improved error messages with troubleshooting steps

### 3. PropertyChanged Subscription Crash
**Symptoms:** `AttributeError: 'FamilyItem' object has no attribute 'add_PropertyChanged'`
**Root Cause:** Attempted manual event subscription on a Python property, not a proper .NET event
**Fix:** Removed manual PropertyChanged subscriptions - WPF data binding handles this automatically

### 4. PropertyChanged Disposal Crash
**Symptoms:** `AttributeError: 'FamilyItem' object has no attribute 'remove_PropertyChanged'`
**Root Cause:** Setting `PropertyChanged = None` in Dispose() triggered event removal attempt
**Fix:** Removed PropertyChanged clearing from Dispose() method

## Technical Changes

**File Modified:** `T3Lab_Lite.extension/lib/GUI/FamilyLoaderDialog.py`

**Changes:**
1. Added `_cancel_current_scan()` method to safely terminate existing scan threads
2. Updated `data_source_changed()` to cancel operations before mode switch
3. Updated `scan_families()` and `load_cloud_families()` with thread safety
4. Added placeholder URL detection with helpful error messages
5. Removed all manual PropertyChanged event subscriptions
6. Simplified `update_family_display()` and `_cleanup()` methods
7. Removed PropertyChanged clearing from `FamilyItem.Dispose()`

## Testing
- ✅ Switching between Local and Cloud modes no longer crashes
- ✅ Cloud mode shows helpful message instead of HTTP 404
- ✅ No more PropertyChanged attribute errors
- ✅ Local mode works without any issues
- ✅ Thread safety prevents concurrent scan operations

## Commits
- `562098e` - Fix crash when switching between local and cloud modes
- `30c3b01` - Improve cloud API error handling with configuration messages
- `b2558e0` - Fix PropertyChanged event subscription crash
- `43a7d52` - Fix remove_PropertyChanged crash in Dispose method

## Impact
- **User Experience:** No more crashes when using family loader
- **Local Mode:** Fully functional for loading families from folders
- **Cloud Mode:** Shows clear instructions for setup (feature ready but needs deployment)
