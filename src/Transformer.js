const { exec, spawn } = require("child_process");
const config = require("./Config.js");
const path = require("path");

module.exports = {};

const RunPowerpoint = async (file) => {

  return new Promise((resolve, reject) => {
    const p = spawn(config.powerpointPath, ["/S", file]);
    if (p.error) {
      reject(p.error);
    }

    resolve(p.pid);
  });

};

const RunTransformer = async (file) => {
  return new Promise((resolve, reject) => {
    let transformerPath = path.resolve(__dirname, "PowerPointTransformer.exe") + " " + file;
    exec(transformerPath, (error, stdout, stderr) => {
      if (error) {
        reject(error);
      }
      
      resolve(stdout? stdout : stderr);
    });
  });
};

module.exports.RunTransformer = RunTransformer;
module.exports.RunPowerpoint = RunPowerpoint;