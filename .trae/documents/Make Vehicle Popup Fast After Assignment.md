## Problem Focus
- After assigning a vehicle, the next Select Vehicle popup takes several seconds before vehicles appear.
- Cause: heavy refresh and cache invalidations, plus background probes and UI work happening before first paint.

## Goals
- Make the next popup open with vehicles preloaded within ~2s of assignment.
- Keep backend authoritative; use client cache only as fallback.

## Server Changes
1. Publish a ready snapshot right after assignment:
   - In the assignment success path, call `ensureVehcacheReady()` to rebuild `vehcache` immediately if needed.
   - Then call `getVehcacheMinimal()` to return a lean list for the popup.
2. Lightweight version stamp:
   - Update a script property (e.g., `VEH_POPUP_VERSION`) on success to indicate a fresh snapshot without forcing client invalidations.

## Client Changes
1. Post-assignment prefetch:
   - In `_VehiclePicker.html` success handler (4280–4314), immediately call `ensureVehcacheReady()` then `getVehcacheMinimal()`.
   - Persist the returned list to `localStorage` (fallback only). Do not clear caches on success.
2. Pause probes while modal is shown:
   - Stop the version monitor timer on modal open; restart it on close.
3. Server-first popup render:
   - On next popup open, call `getVehcacheMinimal()` directly and render immediately.
   - Defer beneficiary/users UI assembly until after vehicles are visible.
4. Debounce refresh:
   - Ensure we do not trigger multiple summary refreshes right after assignment; keep a single `ensureVehcacheReady()` call.

## UX & Monitoring
- Keep the loader visible only until server list arrives; reveal instantly on readiness.
- Log both fetch time and time-to-visible to confirm latency is gone.

## Expected Outcome
- Vehicles appear almost instantly on the next popup after assignment.
- Fewer moving parts on the critical path; cache becomes non-blocking fallback.

Do you approve this plan?