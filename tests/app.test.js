// Minimal sanity tests for Alterations Pinning Certification helper logic.
const modules = ['M1', 'M2', 'M3', 'M4', 'M5'];

function isCertifiedFromModules(modulesPassed) {
  return modules.every(id => modulesPassed.includes(id));
}

function progressPercentage(completedCount, total = modules.length) {
  if (total === 0) return 0;
  return Math.round((completedCount / total) * 100);
}

(function run() {
  const certifiedSet = ['M1', 'M2', 'M3', 'M4', 'M5'];
  const partialSet = ['M1', 'M3'];

  if (!isCertifiedFromModules(certifiedSet)) {
    throw new Error('Certification check failed for full completion');
  }

  if (isCertifiedFromModules(partialSet)) {
    throw new Error('Certification check failed for partial completion');
  }

  if (progressPercentage(3, 5) !== 60) {
    throw new Error('Progress percentage should be 60% for 3/5 complete');
  }

  if (progressPercentage(0, 5) !== 0) {
    throw new Error('Progress percentage should be 0% for 0 complete');
  }

  console.log('All tests passed');
})();
