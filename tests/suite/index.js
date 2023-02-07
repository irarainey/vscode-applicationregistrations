"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.jestTestRunnerForVSCodeE2E = void 0;
const jest_1 = require("jest");
const path = require("path");
exports.jestTestRunnerForVSCodeE2E = {
    run(testsRoot, reportTestResults) {
        const projectRootPath = process.cwd();
        const config = path.join(projectRootPath, "jest.e2e.config.js");
        (0, jest_1.runCLI)({ config }, [projectRootPath])
            .then((jestCliCallResult) => {
            jestCliCallResult.results.testResults.forEach((testResult) => {
                testResult.testResults
                    .filter((assertionResult) => assertionResult.status === "passed")
                    .forEach(({ ancestorTitles, title, status }) => {
                    console.info(`  ● ${ancestorTitles} > ${title} (${status})`);
                });
            });
            jestCliCallResult.results.testResults.forEach((testResult) => {
                if (testResult.failureMessage) {
                    console.error(testResult.failureMessage);
                }
            });
            reportTestResults(undefined, jestCliCallResult.results.numFailedTests);
        })
            .catch((errorCaughtByJestRunner) => {
            reportTestResults(errorCaughtByJestRunner, 0);
        });
    }
};
//# sourceMappingURL=index.js.map