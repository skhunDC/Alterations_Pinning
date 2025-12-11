// Minimal sanity tests for Alterations Pinning Certification logic.
const modules = ['M1', 'M2', 'M3', 'M4', 'M5'];

function passedAllModules(status) {
  return modules.every(id => status.completedModules.includes(id));
}

function calculateScore(correct, total) {
  return Math.round((correct / total) * 100);
}

(function run() {
  const mockStatus = { completedModules: ['M1', 'M2', 'M3', 'M4', 'M5'] };
  if (!passedAllModules(mockStatus)) {
    throw new Error('Certification check failed');
  }
  if (calculateScore(4, 5) !== 80) {
    throw new Error('Score calculation failed');
  }
  console.log('All tests passed');
})();
