const path = require('path');
const UglifyJSPlugin = require('uglifyjs-webpack-plugin');

module.exports = {
 entry: './src/CTMemberEditFormCSR.ts',
 output: {
   filename: 'bundle.js',
   path: path.resolve(__dirname, 'dist'),
 },
 module: {
   rules: [
     {
      use: 'ts-loader',
      test: /\.ts$/,
      exclude: ["/node_modules/", "/dist/"]
     },
   ]
 },
 // plugins: [new UglifyJSPlugin({sourceMap:true})],
 // resolve: {
 //  extensions: [".ts", ".js"]
// },
devtool: 'source-map'
};