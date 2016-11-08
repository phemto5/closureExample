var webpack = require('webpack');

module.exports = {
    entry: {
        Dynamics: './Dynamics/goToSPFolder.jsx',
        SharePoint: './Sharepoint/handleSPFolderRedirection.jsx'
    },
    output: {
        path: './bin',
        filename: '[name].bundle.js'
    },
    module: {
        loaders: [{
            test: /\.jsx?$/,
            exclude: /node_modules/,
            loader: 'babel-loader',
            query: {
                presets: ['es2015', 'react']
            }
        },
        {
            test: /.html/ | /.css/,
            exclude: /.jsx?$/,
            loader: "file?name=[name].[ext]&context=./bin"
        }
        ]
    }
}