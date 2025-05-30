const { execSync } = require("child_process");
const fs = require("fs");

const tag = execSync("git describe --tags --abbrev=0").toString().trim();
const version = tag.replace(/^v/, "");

const pkg = JSON.parse(fs.readFileSync("package.json", "utf8"));
pkg.version = version;
fs.writeFileSync("package.json", JSON.stringify(pkg, null, 2) + "\n");

console.log(`📦 Synced version to ${version}`);
