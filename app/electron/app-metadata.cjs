const fs = require("fs");
const path = require("path");

function getAppMetadata() {
  const packageJsonPath = path.join(__dirname, "../../package.json");

  try {
    const raw = fs.readFileSync(packageJsonPath, "utf8");
    const pkg = JSON.parse(raw);
    const build = pkg && typeof pkg.build === "object" ? pkg.build : {};

    return {
      appName: String(build.productName || pkg.productName || "PAD Maintenance Routière"),
      appVersion: String(pkg.version || "0.0.0")
    };
  } catch (_error) {
    return {
      appName: "PAD Maintenance Routière",
      appVersion: "0.0.0"
    };
  }
}

module.exports = { getAppMetadata };
