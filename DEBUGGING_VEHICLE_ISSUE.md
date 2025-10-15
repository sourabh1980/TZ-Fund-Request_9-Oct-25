# Debugging Vehicle Number Issue

## The Problem
Vehicle numbers from the `Vehicle_InUse` sheet are not being populated in the vehicle number field when the page loads.

## Root Cause
This is likely a **race condition** where:
1. Page loads and starts fetching data from `Vehicle_InUse` sheet (async operation)
2. User selects project/team → Rows are immediately created
3. Rows try to lookup vehicle assignments, but the data hasn't loaded yet
4. Vehicle fields remain empty

## How to Debug

### Step 1: Open Browser Console
1. Open your web app in Google Chrome/Firefox
2. Press `F12` or right-click → "Inspect" → "Console" tab

### Step 2: Check Current State
In the console, type:
```javascript
debugVehicleInUseState()
```

This will show you:
- ✅ Whether data has been loaded from the `Vehicle_InUse` sheet
- 📊 How many records are stored
- 🗂️ All the mappings (beneficiary → vehicle)
- 📋 Current rows in the table
- 🔍 Whether vehicle fields are populated

### Step 3: Analyze the Debug Output

#### Check 1: Is data loaded?
Look for:
```
📊 Statistics:
  • Records count: X
```
If `X` is 0, the data hasn't loaded from the sheet yet.

#### Check 2: Are beneficiary names matching?
Compare the beneficiary names in:
- **"All Records"** section (from Vehicle_InUse sheet)
- **"Current Rows in Table"** section (from your form)

They must match **EXACTLY** (including spaces, capitalization after normalization).

#### Check 3: Are teams/projects matching?
If beneficiaries have the same name but different teams, the system uses team+beneficiary matching.

Check if the team names match exactly between:
- The records (from Vehicle_InUse sheet)
- The rows (from your form)

### Step 4: Watch Real-Time Logs

When you load the page or select a project/team, watch the console for these logs:

**Data Loading:**
```
[VehicleInUse DEBUG] 🌐 Starting fetch from Vehicle_InUse sheet...
[VehicleInUse DEBUG] ✅ Fetch successful
[VehicleInUse DEBUG] 📥 storeVehicleInUseRecords called with X records
```

**Row Processing:**
```
[VehicleInUse DEBUG] 🔧 fillRowFromItem - Processing vehicle assignment
[VehicleInUse DEBUG]   Beneficiary: "Mohassin"
[VehicleInUse DEBUG]   Team: "HW_VOD_M_Moh"
[VehicleInUse DEBUG]   ✅ Found via exact map: T-1234
```

**Assignment Application:**
```
[VehicleInUse DEBUG] 🔍 Starting applyVehicleInUseAssignmentsToRows
[VehicleInUse DEBUG]   ✅ ASSIGNING vehicle "T-1234" via exact match
```

### Step 5: Manual Actions

If vehicles aren't showing up, try these commands in the console:

**Force refresh data from sheet:**
```javascript
window.refreshVehicleInUseData()
```

**Force reapply assignments to all rows:**
```javascript
window.applyVehicleInUseAssignmentsToRows(true)
```

## Common Issues & Solutions

### Issue 1: No records loaded
**Symptoms:** `Records count: 0`

**Possible causes:**
- `Vehicle_InUse` sheet doesn't exist
- Sheet is empty
- Sheet has no rows with status "IN USE"
- Column names don't match expected headers

**Solution:** Check your `Vehicle_InUse` sheet and ensure:
- Sheet exists and is named exactly "Vehicle_InUse"
- Has headers: `R.Beneficiary`, `Vehicle Number`, `Team`, `Project`, `Status`
- Has at least one row with `Status` = "IN USE"

### Issue 2: Names don't match
**Symptoms:** Records show "John Smith" but rows show "john smith"

**Solution:** The system normalizes names (lowercase, trim whitespace), but check for:
- Extra spaces
- Special characters
- Different spellings

### Issue 3: Team/Project mismatch
**Symptoms:** Beneficiary exists in records but vehicle not assigned

**Solution:** Check that team and project names match exactly between:
- The `Vehicle_InUse` sheet
- The beneficiary data (from project selection)

### Issue 4: Timing issue
**Symptoms:** Vehicles don't show on first load but appear after refresh

**Solution:** This confirms the race condition. The system should auto-apply after data loads, but if not:
1. Wait 2-3 seconds after page load
2. Select your project/team again
3. Or use `window.applyVehicleInUseAssignmentsToRows(true)` in console

## Example Debug Session

```javascript
// 1. Check current state
debugVehicleInUseState()

// Output shows:
// Records count: 3
// Exact Matches:
//   • "Mohassin" → T-1234
//   • "Ahmed" → T-5678
//
// Current Rows:
//   Row 1:
//     Beneficiary: "Mohassin"
//     Vehicle: ""  ← EMPTY!
//     Team: "HW_VOD_M_Moh"

// 2. Try manual apply
window.applyVehicleInUseAssignmentsToRows(true)

// 3. Check again
debugVehicleInUseState()

// Should now show:
//   Row 1:
//     Beneficiary: "Mohassin"
//     Vehicle: "T-1234"  ← FILLED!
```

## Next Steps

1. Load your page
2. Open console (F12)
3. Run `debugVehicleInUseState()`
4. Share the output with the developer
5. Look for mismatches in names/teams/projects
6. Try manual commands if needed

---

**Need more help?** The debug logs will show exactly where the lookup is failing.
