const functions = require("firebase-functions");
const crypto = require("crypto");

// The verification function is named 'verify'
exports.verify = functions.https.onRequest(async (req, res) => {
    // Set CORS headers to allow your web checker to communicate with this function
    res.set('Access-Control-Allow-Origin', '*');
    if (req.method === 'OPTIONS') {
        // Handle CORS preflight request
        res.set('Access-Control-Allow-Methods', 'POST');
        res.set('Access-Control-Allow-Headers', 'Content-Type');
        res.status(204).send('');
        return;
    }

    try {
        // Extract the required data sent from the website
        const { jsonReport, calculatedDocHash, token } = req.body;
        const checks = [];
        let success = true;

        if (!jsonReport || !calculatedDocHash || !token) {
            return res.status(400).json({ success: false, message: "Missing required fields (jsonReport, calculatedDocHash, token)." });
        }
        
        // --- DATA EXTRACTION (TELEMETRY for Display) ---
        const sessions = jsonReport.session?.sessions || [];
        
        const flags = {
            largePaste: 0,
            speedSpike: 0,
            sustainedSpeed: 0,
            totalSessions: sessions.length
        };

        sessions.forEach(s => {
            if (s.flags?.largePaste) flags.largePaste++;
            if (s.flags?.speedSpike) flags.speedSpike++;
            if (s.flags?.sustainedSpeed) flags.sustainedSpeed++;
        });
        
        const exportTimestamp = jsonReport.session?.summary?.exportTimestamp || 'N/A (Add-in version may be old)';
        // ------------------------------------

        // --- 1. REPORT INTEGRITY CHECK (Checks if the JSON report itself was manually modified) ---
        const sessionString = JSON.stringify(jsonReport.session);
        const calculatedReportHash = crypto
            .createHash("sha256")
            .update(sessionString)
            .digest("hex");

        if (calculatedReportHash !== jsonReport.hash) {
            checks.push("❌ Report Data Tampering Detected! The report's content hash is invalid.");
            success = false;
        } else {
            checks.push("✅ Report Integrity Verified: Report content is authentic.");
        }

        // --- 2. TOKEN MATCH CHECK ---
        if (jsonReport.token !== token) {
            checks.push("❌ Token Mismatch: The token provided does not match the token stored in the report.");
            success = false;
        } else {
            checks.push("✅ Token Match: The token is valid.");
        }

        // --- 3. DOCUMENT INTEGRITY CHECK (Checks if the DOCX was modified after the report) ---
        const storedDocumentHash = jsonReport.documentHash;

        if (!storedDocumentHash) {
            checks.push("⚠️ Report is missing Document Hash. Cannot verify document integrity. (Requires updated add-in)");
        } else if (calculatedDocHash !== storedDocumentHash) {
            checks.push("❌ Document Tampering Detected! Submitted document's hash does not match the report's stored hash.");
            checks.push(`    > Hash in Report: ${storedDocumentHash.substring(0, 12)}...`);
            checks.push(`    > Calculated Hash: ${calculatedDocHash.substring(0, 12)}...`);
            success = false; // CRITICAL FAILURE
        } else {
            checks.push("✅ Document Integrity Verified: Submitted document matches the reported version.");
        }

        // --- FINAL RESPONSE ---
        return res.json({
            success: success,
            message: success ? "✅ Verification Successful" : "❌ Verification Failed",
            details: checks,
            telemetry: {
                flags: flags,
                exportTimestamp: exportTimestamp
            }
        });

    } catch (e) {
        console.error("Verification error:", e);
        // Handle unexpected server-side errors
        return res.status(500).json({ success: false, message: "Error verifying file.", details: [`Server Error: ${e.message}`] });
    }
});