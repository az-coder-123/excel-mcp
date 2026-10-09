/** @type {import('jest').Config} */
module.exports = {
  testEnvironment: 'node',
  roots: ['<rootDir>/tests'],
  // Source imports use ESM-style ".js" suffixes; strip them for CommonJS ts-jest transforms
  moduleNameMapper: {
    '^(\\.{1,2}/.*)\\.js$': '$1',
  },
  transform: {
    '^.+\\.tsx?$': [
      'ts-jest',
      {
        tsconfig: {
          module: 'commonjs',
          moduleResolution: 'node',
          esModuleInterop: true,
          target: 'ES2022',
          strict: true,
          noUnusedLocals: false,
          noUnusedParameters: false,
          types: ['node', 'jest'],
        },
      },
    ],
  },
  clearMocks: true,
};