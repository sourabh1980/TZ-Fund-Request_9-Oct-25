function testGetVehicleInUseSummary() {
  try {
    console.log('Testing getVehicleInUseSummary...');
    const result = getVehicleInUseSummary();
    console.log('Result OK:', result.ok);
    console.log('Assignments count:', result.assignments ? result.assignments.length : 'N/A');
    if (!result.ok) {
      console.error('Error:', result.error);
    } else if (result.assignments.length > 0) {
      console.log('Sample assignment:', JSON.stringify(result.assignments[0]));
    }
    return result;
  } catch (e) {
    console.error('Test failed:', e);
    return e.toString();
  }
}
