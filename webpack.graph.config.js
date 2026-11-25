const baseConfig = require('./webpack.config.js');

module.exports = {
  ...baseConfig,

  entry: './dev/graph-runtime-source.js',

  output: {
    ...baseConfig.output,
    filename: 'graph-runtime.js',
  },
};
