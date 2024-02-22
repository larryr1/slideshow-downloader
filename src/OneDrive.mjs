import axios from "axios";
import { createWriteStream } from "fs";
import config from "./Config.js";

export async function GetLatestSharedSlideshow(graphToken) {

  // Get recent shared items from OneDrive
  const response = await axios.get("https://graph.microsoft.com/v1.0/me/drive/sharedwithme", { headers: { Authorization: `Bearer ${graphToken}`}});


  const sharedSlideshow = response.data.value.filter(item => {
    return (item.createdBy.user.email.toLowerCase() == config.slideshowSender && item.name == config.slideshowName);
  })[0];

  // Check for none
  if (!sharedSlideshow) {
    throw new Error("Could not locate the specified shared slideshow.");
  }
    
  return sharedSlideshow;
}

// Used to get the lastModifiedDateTime of a DriveItem.
export async function GetDriveItemLastModifiedDateTime(graphToken, driveId, fileId) {

  const item = await GetLatestSharedSlideshow(graphToken);

  const requestUrl = `https://graph.microsoft.com/v1.0/drives/${item.remoteItem.parentReference.driveId}/items/${item.id}?select=lastModifiedDateTime`;
  const itemData = await axios.get(requestUrl, { headers: { Authorization: `Bearer ${graphToken}`}});
  if (itemData.data["error"]) {
    throw new Error(itemData.data.error);
  }

  return itemData.data.lastModifiedDateTime;
}

// Used to get the unique one-time download URL for a drive item.
export async function GetDownloadUrl(graphToken, driveId, fileId) {
  const response = await axios.get(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}?select=id,@microsoft.graph.downloadUrl`, { headers: { Authorization: `Bearer ${graphToken}`}});
  return response.data["@microsoft.graph.downloadUrl"];
}

// Download and save a OneDrive file, given the url.
export async function DownloadFileFromUrl(url, filename) {

  const response = await axios.get(url, { responseType: "stream"});
  return new Promise((resolve, reject) => {
    
    response.data.pipe(createWriteStream(filename));

    response.data.on("end", () => { resolve(true); });
    response.data.on("error", () => { reject("Error while piping response to file."); });
  });
  
  
}