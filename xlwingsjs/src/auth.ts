// Office.auth.getAccessToken claims that it does everything that this module does,
// only it doesn't: https://github.com/OfficeDev/office-js/issues/3298

let accessToken = null;
let isRenewingToken = false;
let tokenLock = false;
let tokenExpiry = null;

function hasKeyExpired() {
  if (!tokenExpiry) {
    return true;
  }
  const currentTime = Math.floor(Date.now() / 1000); // Convert to seconds
  // Renew 15 minutes before expiry
  return currentTime >= tokenExpiry - 15 * 60;
}

async function renewAccessToken() {
  console.log("Renewing access token");
  try {
    accessToken = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
    });

    // Read exp
    let payload = accessToken.split(".")[1];
    // Add padding to base64Url string and then use atob() to decode it
    let base64 = payload.replace(/-/g, "+").replace(/_/g, "/");
    while (base64.length % 4) {
      base64 += "=";
    }
    let decodedPayload = JSON.parse(window.atob(base64));
    tokenExpiry = decodedPayload.exp;

    accessToken = "Bearer " + accessToken;
  } catch (error) {
    let token_error = `Error ${error.code}: ${error.message}`;
    console.log(token_error);
    // return token error so it can be logged on backend
    accessToken = token_error;
  } finally {
    tokenLock = false;
  }
}

export async function getAccessToken() {
  if (!accessToken || hasKeyExpired()) {
    if (!tokenLock) {
      tokenLock = true;
      isRenewingToken = true;
      await renewAccessToken();

      isRenewingToken = false;
    } else {
      while (isRenewingToken) {
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
    }
  }
  return accessToken;
}
