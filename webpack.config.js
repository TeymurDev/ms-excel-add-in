const path = require('path');
const CustomFunctionsMetadataPlugin = require('custom-functions-metadata-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const devCerts = require('office-addin-dev-certs');

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return {
    ca: httpsOptions.ca,
    key: httpsOptions.key,
    cert: httpsOptions.cert,
  };
}

module.exports = async (env, options) => {
  const dev = options.mode === 'development';

  return {
    devtool: dev ? 'source-map' : false,
    entry: {
      functions: './src/functions/functions.ts',
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: '[name].js',
      clean: true,
    },
    resolve: {
      extensions: ['.ts', '.js'],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: 'ts-loader',
        },
      ],
    },
    plugins: [
      // Auto-generates functions.json from the @customfunction JSDoc tags
      new CustomFunctionsMetadataPlugin({
        output: 'functions.json',
        input: './src/functions/functions.ts',
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: 'manifest.xml', to: '[name][ext]' },
          { from: 'src/functions/functions.html', to: 'functions.html' },
        ],
      }),
    ],
    devServer: {
      hot: true,
      server: {
        type: 'https',
        options: dev ? await getHttpsOptions() : {},
      },
      port: 3000,
    },
  };
};
