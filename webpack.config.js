const path = require('node:path');
const TerserPlugin = require('terser-webpack-plugin');

module.exports = {
    context: path.resolve(__dirname, 'src'),
    devtool: 'inline-source-map',
    entry: './index.ts',
    mode: 'development',
    target: ['web', 'es5'],
    module: {
        rules: [
            {
                test: /\.tsx?$/,
                use: 'ts-loader',
                exclude: /node_modules/,
            },
        ],
    },
    output: {
        filename: 'build.min.js',
        path: path.resolve(__dirname, 'build'),
    },
    resolve: {
        extensions: ['.tsx', '.ts', '.jsx', '.js'],
    },
    optimization: {
        minimize: true,
        minimizer: [
            new TerserPlugin({
                extractComments: false,
                terserOptions: {
                    ecma: 5,
                    ie8: true,
                    safari10: true,
                    compress: {
                        ecma: 5,
                    },
                    mangle: {
                        safari10: true,
                    },
                },
            }),
        ],
    },
};
