const path = require('path');
const webpack = require('webpack');

module.exports = function createConfig({ root = __dirname, outputPath, entry, config } = {}) {
  if (!['graph-runtime', 'signin'].includes(entry) || !outputPath || !config) {
    throw new Error('Use npm run build:dev or npm run build:prod to build a complete package.');
  }

  return {
    context: root,
    mode: 'production',

    entry: path.join(root, 'src', `${entry}-source.js`),

    output: {
      path: outputPath,
      filename: `${entry}.js`,
      clean: false,
    },

    optimization: {
      splitChunks: false,
      runtimeChunk: false,
    },

    plugins: [
      new webpack.DefinePlugin({
        __AUTH_CLIENT_ID__: JSON.stringify(config.clientId),
        __AUTH_AUTHORITY__: JSON.stringify(config.authority),
      }),
      new webpack.optimize.LimitChunkCountPlugin({ maxChunks: 1 }),
    ],
  };
};
