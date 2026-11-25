const path = require('path');
const webpack = require('webpack');

module.exports = {
  mode: 'production',

  entry: './dev/signin-source.js',

  output: {
    path: path.resolve(__dirname, 'dev'),
    filename: 'signin.js',
    clean: false,
  },

  optimization: {
    splitChunks: false,
    runtimeChunk: false,
  },

  plugins: [
    new webpack.optimize.LimitChunkCountPlugin({
      maxChunks: 1,
    }),
  ],
};
