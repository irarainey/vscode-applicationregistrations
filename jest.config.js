/* eslint-disable @typescript-eslint/naming-convention */
module.exports = {
  preset: "ts-jest",
  testEnvironment: "node",
  preset: "ts-jest/presets/default-esm",
  extensionsToTreatAsEsm: ['.ts'],
  moduleDirectories: [
    'node_modules',
    'src'
  ],
  moduleFileExtensions: [
    "ts", 
    "js"
  ],
  modulePaths: [
    "<rootDir>",
    "<rootDir>/node_modules"
  ],
  moduleNameMapper: {
    '^@/(.*)$': '<rootDir>/src/$1'
  },
  transform: {
    '^.+\\.ts?$': [
      'ts-jest',
      {
        useESM: true
      }
    ]
  },
  testMatch: ["<rootDir>/test/**/*.test.ts"]
};
