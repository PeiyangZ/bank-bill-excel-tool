const { createSmokeContext } = require('./smoke/support');
const { runSmokeScenarios } = require('./smoke/scenarios');

function run() {
  const context = createSmokeContext();
  runSmokeScenarios(context);
  console.log('smoke test passed');
}

run();
