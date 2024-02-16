const path = require("path");

const configPath = path.join(process.cwd(), "./slideshow_config.json");

module.exports = require(configPath);