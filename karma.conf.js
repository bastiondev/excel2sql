module.exports = function(config) {
  config.set({
    // base path that will be used to resolve all patterns (eg. files, exclude)
    basePath: '',

    // frameworks to use
    frameworks: ['mocha', 'chai'],
    
    // client configuration
    client: {
      mocha: {
        timeout: 10000 // 10 seconds
      }
    },

    // list of files / patterns to load in the browser
    files: [
      // Load ExcelJS from CDN
      { pattern: 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js', included: true },
      // Load our browser bundle
      { pattern: 'dist/excel2sql.browser.js', included: true },
      // Load test data
      { pattern: 'test/browser/test-data.js', included: true, served: true, watched: true },
      // Load test files
      { pattern: 'test/browser/**/*.spec.js', included: true }
    ],

    // list of files / patterns to exclude
    exclude: [],

    // preprocess matching files before serving them to the browser
    preprocessors: {
      'test/browser/**/*.spec.js': ['webpack']
    },

    webpack: {
      mode: 'development',
      module: {
        rules: [
          {
            test: /\.tsx?$/,
            use: 'ts-loader',
            exclude: /node_modules/,
          }
        ]
      },
      resolve: {
        extensions: ['.tsx', '.ts', '.js']
      }
    },

    // test results reporter to use
    reporters: ['progress'],

    // web server port
    port: 9876,

    // enable / disable colors in the output (reporters and logs)
    colors: true,

    // level of logging
    logLevel: config.LOG_INFO,

    // enable / disable watching file and executing tests whenever any file changes
    autoWatch: false,

    // start these browsers
    browsers: ['ChromeHeadless'],

    // Continuous Integration mode
    // if true, Karma captures browsers, runs the tests and exits
    singleRun: true,

    // Concurrency level
    // how many browser should be started simultaneous
    concurrency: Infinity
  });
};
