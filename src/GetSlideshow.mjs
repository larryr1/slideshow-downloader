#!/usr/bin/env node
import config from "./Config.js";
import { GetAuthorizationCode, GetGraphToken } from "./Authorization.mjs";
import { GetDownloadUrl, DownloadFileFromUrl, GetLatestSharedSlideshow, GetDriveItemLastModifiedDateTime } from "./OneDrive.mjs";
import { RunPowerpoint, RunTransformer } from "./Transformer.js";
import { existsSync, unlinkSync } from "fs";
import { fileTypeFromFile } from "file-type";
import { kill } from "process";
import { exec } from "child_process";
import { resolve } from "path";

var graphToken = "";
var driveItemId = "";
var driveItemDriveId = "";
const startTime = new Date();


function authorizationCodeError(e) {
  throw new Error(e);
}

function graphTokenError(e) {
  throw new Error(e);
}

(async () => {

  // Delete leftover files.
  console.log("Deleting old files.");
  if (existsSync("slideshow.pptx")) { unlinkSync("slideshow.pptx"); }

  // Obtain Microsoft authorization code.
  console.log("Obtaining Graph authorization code.");
  const authorizationCode = await GetAuthorizationCode().catch(authorizationCodeError);

  // Exchange authorization code for Microsoft Graph API access token.
  console.log("Exchanging authorization code for Graph access token.");
  graphToken = await GetGraphToken(authorizationCode, config.clientSecret).catch(graphTokenError);

  // Locate the slideshow resource
  console.log("Getting latest shared slideshow.");
  const latestSharedFile = await GetLatestSharedSlideshow(graphToken);

  driveItemId = latestSharedFile.remoteItem.id;
  driveItemDriveId = latestSharedFile.remoteItem.parentReference.driveId;

  // Get the slideshow URL
  console.log("Getting download url for shared slideshow.");
  const fileDownloadUrl = await GetDownloadUrl(
    graphToken,
    latestSharedFile.remoteItem.parentReference.driveId,
    latestSharedFile.remoteItem.id
  );

  // Download the slideshow
  console.log("Downloading shared slideshow.");
  await DownloadFileFromUrl(fileDownloadUrl, "slideshow.pptx");

  // Check to make sure downloaded file is a PowerPoint. The transformer throws cryptic errors if it's passed a non-pptx file.
  const fileType = await fileTypeFromFile("slideshow.pptx");
  const requiredMimeType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  const requiredExtension = "pptx";

  try {
    if (fileType.mime != requiredMimeType) {
      throw new Error(`File MIME type ${fileType.mime} does not match required MIME type of ${requiredMimeType}.`);
    } else if (fileType.ext.toLowerCase() != requiredExtension) {
      throw new Error(`File extension .${fileType.ext} does not match required extension .${requiredExtension}.`);
    }
  } catch (error) {
    console.log("Error: " + error);
    
    // Check for an existing transformed slideshow and use it
    if (existsSync("slideshow.pptx-transformed.pptx")) {
      console.log("Starting a cached slideshow.");
      await RunPowerpoint("slideshow.pptx-transformed.pptx");
    } else {
      console.log("No slideshow is cached. Aborting.");
    }

    return;
  }
  
  // Remove file before transformer
  if (existsSync("slideshow.pptx-transformed.pptx")) { unlinkSync("slideshow.pptx-transformed.pptx"); }

  // The transformer uses PowerPoint Interop DLLs to apply an automatic transition to every slide and sets the slideshow to loop.
  console.log("Running transformer.");
  await RunTransformer("slideshow.pptx");

  console.log("Waiting for transformer to close file.");  
  setTimeout(async () => {
    // Starts the PowerPoint in slideshow mode.
    console.log("Starting PowerPoint.");
    await RunPowerpoint("slideshow.pptx-transformed.pptx");
    
  }, 2000);

  console.log(`Checking for updates every ${config.updateCheckInterval/60000} minute(s).`)

  // Periodically check for updates
  setInterval(updateCheck, config.updateCheckInterval);

})();

async function updateCheck() {
  // Locate the slideshow resource
  console.log("Checking for updates.");
  const lastModifiedDateTime = await GetDriveItemLastModifiedDateTime(graphToken);

  // return it hasnt been modified
  if (new Date(lastModifiedDateTime) < startTime) {
    console.log("No update available.");
    return;
  };

  console.log("Update detected. Exiting in 3 seconds.");

  // It was modified, restart the slideshow.
  exec(config.updateTriggerCommand);
  
  setTimeout(() => {
    console.log("Exiting...");
    process.exit(0);
  }, 3000);
  
}