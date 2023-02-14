/* eslint-disable @typescript-eslint/naming-convention */
module.exports = {
  preset: "ts-jest",
  testEnvironment: "node",
  preset: "ts-jest/presets/default-esm",
  extensionsToTreatAsEsm: [".ts"],
  moduleFileExtensions: [ "ts", "js"],
  modulePathIgnorePatterns: [
    "<rootDir>/dist/",
    "<rootDir>/out/",
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
  testMatch: ["<rootDir>/src/tests/*.test.ts"],
  coverageProvider: "v8",
  collectCoverage: true,
  collectCoverageFrom: [
    "<rootDir>/src/services/*.ts",
    "<rootDir>/src/data/*.ts",
    "<rootDir>/src/models/*.ts",
    "<rootDir>/src/utils/*.ts",
    "<rootDir>/src/error-handler.ts",
    "<rootDir>/src/constants.ts",
    "!<rootDir>/src/utils/exec-shell-cmd.ts"
  ]
};
