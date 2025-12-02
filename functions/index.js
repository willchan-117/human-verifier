const functions = require("firebase-functions");
const crypto = require("crypto");

exports.verify = functions.https.onRequest(async (req, res) => {
  try {
    const { json, fileHash, token } = req.body;

    // 1. Recreate the hash
    const sessionString = JSON.stringify(json.sessions);
    const calculatedHash = crypto
      .createHash("sha256")
      .update(sessionString)
      .digest("hex");

    if (calculatedHash !== fileHash) {
      return res.json({ success: false, message: "❌ Hash mismatch (file tampered)" });
    }

    if (json.token !== token) {
      return res.json({ success: false, message: "❌ Invalid token" });
    }

    return res.json({ success: true, message: "✅ File is authentic" });

  } catch (e) {
    return res.json({ success: false, message: "Error verifying file." });
  }
});
