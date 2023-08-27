const path = require("path");

module.exports = {
  mode: "production",
  devtool: "source-map",
  entry: {
    xlwings: "./src/xlwings.ts",
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: "ts-loader",
        exclude: /node_modules/,
      },
    ],
  },
  resolve: {
    extensions: [".tsx", ".ts", ".js"],
  },
  output: {
    filename: "xlwings.min.js",
    path: path.resolve(__dirname, "dist"),
    umdNamedDefine: true,
    library: "xlwings",
  },
};
