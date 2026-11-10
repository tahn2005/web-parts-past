// Type declarations for static asset imports used in gulp-based SPFx projects.
// In heft-based projects, these typings are generated at build time by
// @rushstack/heft-static-asset-typings-plugin. Gulp projects require them
// to be declared statically.
declare module '*.png' {
  const value: string;
  export default value;
}

declare module '*.jpg' {
  const value: string;
  export default value;
}

declare module '*.svg' {
  const value: string;
  export default value;
}

