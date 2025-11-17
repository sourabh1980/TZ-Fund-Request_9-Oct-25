// Test script for New Car toggle functionality
// This script validates that the backend correctly loads data from the 'vehicle' tab
// when isNewCar parameter is true

function testNewCarToggle() {
  console.log('Testing New Car toggle functionality...');

  // Test 1: Verify getVehiclePickerData with isNewCar=true loads from 'vehicle' sheet
  console.log('Test 1: Testing getVehiclePickerData(true)...');
  try {
    const result = getVehiclePickerData(true);
    console.log('✅ getVehiclePickerData(true) returned:', result);

    if (result && result.vehicles && Array.isArray(result.vehicles)) {
      console.log(`✅ Found ${result.vehicles.length} vehicles from vehicle sheet`);

      // Check if vehicles have the expected columns (Z-AE mapping)
      const sampleVehicle = result.vehicles[0];
      if (sampleVehicle) {
        const expectedFields = ['make', 'model', 'category', 'usageType', 'owner', 'carNumber'];
        const missingFields = expectedFields.filter(field => !sampleVehicle.hasOwnProperty(field));

        if (missingFields.length === 0) {
          console.log('✅ Vehicle object has all expected fields:', expectedFields.join(', '));
        } else {
          console.log('⚠️ Missing fields in vehicle object:', missingFields.join(', '));
        }
      }
    } else {
      console.log('❌ getVehiclePickerData(true) did not return expected structure');
    }
  } catch (error) {
    console.log('❌ Error testing getVehiclePickerData(true):', error);
  }

  // Test 2: Verify getVehiclePickerData with isNewCar=false returns available vehicles derived from CarT_P
  console.log('Test 2: Testing getVehiclePickerData(false)...');
  try {
    const result = getVehiclePickerData(false);
    console.log('✅ getVehiclePickerData(false) returned:', result);

    if (result && result.vehicles && Array.isArray(result.vehicles)) {
      console.log(`✅ Found ${result.vehicles.length} available vehicles directly from CarT_P summary data`);
    } else {
      console.log('❌ getVehiclePickerData(false) did not return expected structure');
    }
  } catch (error) {
    console.log('❌ Error testing getVehiclePickerData(false):', error);
  }

  // Test 3: Verify cache invalidation works
  console.log('Test 3: Testing cache invalidation...');
  try {
    const cacheKey = 'vehicle_sheet_cache_vehicle';
    // First, populate cache
    const testData = [{ carNumber: 'TEST123', make: 'Test Make' }];
    const cache = CacheService.getScriptCache();
    cache.put(cacheKey, JSON.stringify(testData), 300); // 5 minutes

    // Verify cache has data
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      console.log('✅ Cache populated successfully');

      // Now test invalidation
      invalidateVehicleSheetCache('vehicle');
      const clearedData = cache.get(cacheKey);

      if (!clearedData) {
        console.log('✅ Cache invalidation successful');
      } else {
        console.log('❌ Cache invalidation failed - data still exists');
      }
    } else {
      console.log('❌ Failed to populate cache for testing');
    }
  } catch (error) {
    console.log('❌ Error testing cache invalidation:', error);
  }

  // Test 4: Verify Vehicle_InUse filtering heuristic for active assignments
  console.log('Test 4: Verifying Vehicle_InUse active status heuristics...');
  const statusCases = [
    { status: 'IN USE', beneficiary: 'Alice', expected: true },
    { status: 'Release', beneficiary: '', expected: false },
    { status: 'AVAILABLE', beneficiary: '', expected: false },
    { status: 'Assigned', beneficiary: 'Bob', expected: true },
    { status: '', beneficiary: '', expected: false },
    { status: '', beneficiary: 'Charlie', expected: true }
  ];
  statusCases.forEach(function(testCase, index) {
    try {
      const result = _isActiveVehicleInUseEntry_({
        assignmentStatus: testCase.status,
        beneficiary: testCase.beneficiary
      });
      if (result === testCase.expected) {
        console.log(`✅ Case ${index + 1} passed for status "${testCase.status}" with beneficiary "${testCase.beneficiary}".`);
      } else {
        console.log(`❌ Case ${index + 1} failed for status "${testCase.status}" with beneficiary "${testCase.beneficiary}". Expected ${testCase.expected} but got ${result}.`);
      }
    } catch (error) {
      console.log(`❌ Error evaluating case ${index + 1}:`, error);
    }
  });

  console.log('Test completed. Check the logs above for results.');
}