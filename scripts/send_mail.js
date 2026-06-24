// Envoie le .pptx par mail via l'API Gmail (OAuth), sans dépendance externe.
// Réutilise le token OAuth de la chaîne LKA (compte joselonm11@gmail.com).
//
// Variables d'environnement :
//   GMAIL_TOKEN_JSON  contenu de connections/token.json (client_id, client_secret, refresh_token, token_uri…)
//   MAIL_FILE         chemin du .pptx à joindre
//   MAIL_WEEK         libellé de la semaine (objet du mail)
//   MAIL_TO           destinataire (défaut g.fondzefe@lkaservices.com)
const fs = require("fs");
const path = require("path");

async function main() {
  const tok = JSON.parse(process.env.GMAIL_TOKEN_JSON || "{}");
  const file = process.env.MAIL_FILE;
  const week = process.env.MAIL_WEEK || "";
  const to = process.env.MAIL_TO || "g.fondzefe@lkaservices.com";
  if (!tok.refresh_token || !tok.client_id || !tok.client_secret) throw new Error("GMAIL_TOKEN_JSON incomplet (client_id/secret/refresh_token)");
  if (!file || !fs.existsSync(file)) throw new Error("MAIL_FILE introuvable : " + file);

  // 1) rafraîchir l'access token
  const tokenRes = await fetch(tok.token_uri || "https://oauth2.googleapis.com/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: tok.client_id, client_secret: tok.client_secret, refresh_token: tok.refresh_token, grant_type: "refresh_token" }),
  });
  const tokenJson = await tokenRes.json();
  if (!tokenJson.access_token) throw new Error("Refresh token échoué : " + JSON.stringify(tokenJson));

  // 2) construire le message MIME (multipart/mixed avec pièce jointe)
  const fname = path.basename(file);
  const b64file = fs.readFileSync(file).toString("base64").replace(/(.{76})/g, "$1\r\n");
  const subject = "Rapport hebdomadaire DTC Assisted" + (week ? " — " + week : "");
  const from = tok.account || "me";
  const body =
    "Bonjour,\r\n\r\n" +
    "Le rapport DTC Assisted" + (week ? " (" + week + ")" : "") + " est généré automatiquement — la semaine est bouclée.\r\n" +
    "Le fichier .pptx est en pièce jointe.\r\n\r\n— Chaîne automatisée LKA × MTN";
  const boundary = "lka_" + Date.now();
  const mime = [
    "From: LKA DTC <" + from + ">",
    "To: " + to,
    "Subject: =?UTF-8?B?" + Buffer.from(subject, "utf8").toString("base64") + "?=",
    "MIME-Version: 1.0",
    'Content-Type: multipart/mixed; boundary="' + boundary + '"',
    "",
    "--" + boundary,
    'Content-Type: text/plain; charset="UTF-8"',
    "Content-Transfer-Encoding: base64",
    "",
    Buffer.from(body, "utf8").toString("base64").replace(/(.{76})/g, "$1\r\n"),
    "",
    "--" + boundary,
    'Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation; name="' + fname + '"',
    "Content-Transfer-Encoding: base64",
    'Content-Disposition: attachment; filename="' + fname + '"',
    "",
    b64file,
    "",
    "--" + boundary + "--",
    "",
  ].join("\r\n");
  const raw = Buffer.from(mime, "utf8").toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");

  // 3) envoyer via Gmail API
  const sendRes = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/messages/send", {
    method: "POST", headers: { Authorization: "Bearer " + tokenJson.access_token, "Content-Type": "application/json" },
    body: JSON.stringify({ raw }),
  });
  const sendJson = await sendRes.json();
  if (!sendRes.ok) throw new Error("Gmail send échoué (" + sendRes.status + ") : " + JSON.stringify(sendJson));
  console.log("✅ Mail envoyé à " + to + " (id " + sendJson.id + ")");
}

main().catch(e => { console.error("❌ " + e.message); process.exit(1); });
