const path = require('path');

module.exports = {
  mode: 'production',
  entry: './src/index.ts',
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/,
      },
    ],
  },
  resolve: {
    extensions: ['.tsx', '.ts', '.js'],
  },
  output: {
    filename: 'excel2sql.browser.js',
    path: path.resolve(__dirname, 'dist'),
    library: {
      name: 'excel2sql',
      type: 'umd',
    },
    globalObject: 'this'
  },
  externals: {
    'exceljs': {
      commonjs: 'exceljs',
      commonjs2: 'exceljs',
      amd: 'exceljs',
      root: 'ExcelJS'
    }
  }
};
