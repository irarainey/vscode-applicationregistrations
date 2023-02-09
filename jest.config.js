/* eslint-disable @typescript-eslint/naming-convention */
module.exports = {
  preset: "ts-jest",
  testEnvironment: "node",
  preset: "ts-jest/presets/default-esm",
  extensionsToTreatAsEsm: [".ts"],
  moduleFileExtensions: [ "ts", "js"],
  modulePathIgnorePatterns: [
    "<rootDir>/out/",
    "<rootDir>/dist/",
    "<rootDir>/packages/"
  ],
  transform: {
    '^.+\\.ts?$': [
      'ts-jest',
      {
        useESM: true
      }
    ]
  },
  testMatch: ["<rootDir>/tests/**/*.test.ts"],
  coverageProvider: "v8",
  collectCoverage: true,
  collectCoverageFrom: [
    "<rootDir>/src/services/*.ts",
    "<rootDir>/src/utils/validation.ts",
    "<rootDir>/src/utils/status-bar.ts",
    "<rootDir>/src/utils/copy-value.ts",
    "<rootDir>/src/utils/escape-string.ts",
    "<rootDir>/src/utils/cli-authentication.ts",
    "<rootDir>/src/data/*.ts",
    "<rootDir>/src/models/*.ts",
    "<rootDir>/src/error-handler.ts",
    "<rootDir>/src/constants.ts"
  ]
};
