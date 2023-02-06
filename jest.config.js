/* eslint-disable @typescript-eslint/naming-convention */
module.exports = {
  preset: "ts-jest",
  testEnvironment: "node",
  preset: "ts-jest/presets/default-esm",
  extensionsToTreatAsEsm: [".ts"],
  moduleFileExtensions: [ "ts", "js"],
  transform: {
    '^.+\\.ts?$': [
      'ts-jest',
      {
        useESM: true
      }
    ]
  },
  testMatch: ["<rootDir>/test/**/*.test.ts"],
  coverageProvider: "v8",
  collectCoverage: true,
  collectCoverageFrom: [
    "<rootDir>/src/services/*.ts",
    "<rootDir>/src/utils/status-bar.ts",
    "<rootDir>/src/utils/copy-value.ts",
    "<rootDir>/src/utils/cli-authentication.ts",
    "<rootDir>/src/types/*.ts",
    "<rootDir>/src/data/*.ts",
    "<rootDir>/src/error-handler.ts",
    "<rootDir>/src/constants.ts"
  ]
};
