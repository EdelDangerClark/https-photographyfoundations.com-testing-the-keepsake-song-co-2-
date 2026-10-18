// Stripe webhook: when someone pays for a song, add them to the
// "Keepsake – Customers" group in Sender, which starts the order emails.
//
// Needs two environment variables in Netlify:
//   STRIPE_WEBHOOK_SECRET  the signing secret of the Stripe webhook (starts whsec_)
//   SENDER_API_TOKEN       a Sender API access token
//
// Stripe calls this at /.netlify/functions/stripe-to-sender for the
// checkout.session.completed event.

import crypto from "node:crypto";

const CUSTOMERS_GROUP = "azNrj5"; // Sender: Keepsake – Customers

// Only the song links start the order emails, not the add-on links.
const SONG_LINKS = new Set([
  "plink_1UQqimDKougkOytzPklep0RD", // Intro Bundle £34.99
  "plink_1UQrGDDKougkOytz6wFhSxco", // Song £34.99
  "plink_1UQrGDDKougkOytzA0nA26ax", // Bundle £44.99
]);

const TOLERANCE_SECONDS = 300;

function verifyStripeSignature(body, header, secret) {
  if (!header) return false;
  const parts = header.split(",").map((p) => p.split("="));
  const t = parts.find(([k]) => k === "t")?.[1];
  const signatures = parts.filter(([k]) => k === "v1").map(([, v]) => v);
  if (!t || signatures.length === 0) return false;
  if (Math.abs(Date.now() / 1000 - Number(t)) > TOLERANCE_SECONDS) return false;

  const expected = crypto
    .createHmac("sha256", secret)
    .update(`${t}.${body}`, "utf8")
    .digest("hex");
  const expectedBuf = Buffer.from(expected, "hex");
  return signatures.some((sig) => {
    const sigBuf = Buffer.from(sig, "hex");
    return sigBuf.length === expectedBuf.length && crypto.timingSafeEqual(sigBuf, expectedBuf);
  });
}

function splitName(full) {
  const words = (full || "").trim().split(/\s+/).filter(Boolean);
  return { firstname: words[0] || "", lastname: words.slice(1).join(" ") };
}

export default async (req) => {
  if (req.method !== "POST") return new Response("Method not allowed", { status: 405 });

  const { STRIPE_WEBHOOK_SECRET, SENDER_API_TOKEN } = process.env;
  if (!STRIPE_WEBHOOK_SECRET || !SENDER_API_TOKEN) {
    console.error("Missing STRIPE_WEBHOOK_SECRET or SENDER_API_TOKEN");
    return new Response("Not configured", { status: 500 });
  }

  const body = await req.text();
  if (!verifyStripeSignature(body, req.headers.get("stripe-signature"), STRIPE_WEBHOOK_SECRET)) {
    return new Response("Bad signature", { status: 400 });
  }

  const event = JSON.parse(body);
  if (event.type !== "checkout.session.completed") {
    return new Response("Ignored event", { status: 200 });
  }

  const session = event.data.object;
  if (!SONG_LINKS.has(session.payment_link)) {
    return new Response("Not a song order", { status: 200 });
  }
  if (session.payment_status !== "paid") {
    return new Response("Not paid", { status: 200 });
  }

  const email = session.customer_details?.email || session.customer_email;
  if (!email) {
    console.error("No email on session", session.id);
    return new Response("No email", { status: 200 });
  }

  const { firstname, lastname } = splitName(session.customer_details?.name);

  const res = await fetch("https://api.sender.net/v2/subscribers", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${SENDER_API_TOKEN}`,
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    body: JSON.stringify({
      email,
      firstname,
      lastname,
      groups: [CUSTOMERS_GROUP],
      trigger_automation: true,
    }),
  });

  if (!res.ok) {
    // A non-2xx reply makes Stripe retry the webhook later.
    console.error("Sender error", res.status, await res.text());
    return new Response("Sender error", { status: 502 });
  }

  console.log("Added to Keepsake – Customers:", session.id);
  return new Response("OK", { status: 200 });
};
