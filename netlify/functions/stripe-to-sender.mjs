// Stripe webhook: when someone pays for a song, add them to the
// "Keepsake – Customers" group in Sender, which starts the order emails,
// and add a row to the Keepsake order tracker Google Sheet.
// Add-on payments (rush, extra style) only go to the Sheet.
//
// Needs these environment variables in Netlify:
//   STRIPE_WEBHOOK_SECRET  the signing secret of the Stripe webhook (starts whsec_)
//   SENDER_API_TOKEN       a Sender API access token
//   SHEET_WEBHOOK_URL      (optional) the order tracker's web app URL
//   SHEET_SECRET           (optional) the secret shown by the Sheet's set-up menu
//
// Stripe calls this at /.netlify/functions/stripe-to-sender for the
// checkout.session.completed event.

import crypto from "node:crypto";

const CUSTOMERS_GROUP = "azNrj5"; // Sender: Keepsake – Customers

// Only the song links start the order emails, not the add-on links.
const SONG_LINKS = new Map([
  ["plink_1UQqimDKougkOytzPklep0RD", "Intro Bundle"], // £34.99, first 15
  ["plink_1UQrGDDKougkOytz6wFhSxco", "Song"], // £34.99
  ["plink_1UQrGDDKougkOytzA0nA26ax", "Bundle"], // £44.99
]);

const ADDON_LINKS = new Map([
  ["plink_1UQrBxDKougkOytz5VQwtrcg", "Rush"], // £9.99
  ["plink_1UQrBxDKougkOytzs8zaRcoJ", "Extra style"], // £9.99
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

// Reads a Stripe custom field as readable text (dropdowns give their label).
function customField(session, key) {
  const field = (session.custom_fields || []).find((f) => f.key === key);
  if (!field) return "";
  if (field.type === "dropdown") {
    const value = field.dropdown?.value;
    const option = (field.dropdown?.options || []).find((o) => o.value === value);
    return option?.label || value || "";
  }
  return field[field.type]?.value || "";
}

// Sends the order to the Google Sheet. A failure here is logged but never
// fails the webhook, so the Sender step is not repeated by a Stripe retry.
async function sendToSheet(order) {
  const { SHEET_WEBHOOK_URL, SHEET_SECRET } = process.env;
  if (!SHEET_WEBHOOK_URL || !SHEET_SECRET) return;
  try {
    const res = await fetch(SHEET_WEBHOOK_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ ...order, secret: SHEET_SECRET }),
      redirect: "follow",
      signal: AbortSignal.timeout(8000),
    });
    const text = await res.text();
    if (!res.ok || !text.startsWith("OK")) {
      console.error("Sheet error", res.status, text.slice(0, 200));
    }
  } catch (err) {
    console.error("Sheet error", err.message);
  }
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
  const songPackage = SONG_LINKS.get(session.payment_link);
  const addon = ADDON_LINKS.get(session.payment_link);
  if (!songPackage && !addon) {
    return new Response("Not a Keepsake order", { status: 200 });
  }
  // "no_payment_required" covers £0 orders made with a 100% discount code.
  if (session.payment_status !== "paid" && session.payment_status !== "no_payment_required") {
    return new Response("Not paid", { status: 200 });
  }

  const email = session.customer_details?.email || session.customer_email;
  if (!email) {
    console.error("No email on session", session.id);
    return new Response("No email", { status: 200 });
  }

  const order = {
    type: songPackage ? "order" : "addon",
    package: songPackage || addon,
    session: session.id,
    created: session.created,
    name: session.customer_details?.name || "",
    email,
    phone: session.customer_details?.phone || "",
    songFor: customField(session, "songfor"),
    occasion: customField(session, "occasion"),
    neededBy: customField(session, "neededby"),
    amount: (session.amount_total ?? 0) / 100,
  };

  if (addon) {
    await sendToSheet(order);
    console.log("Add-on sent to Sheet:", session.id);
    return new Response("OK", { status: 200 });
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
  await sendToSheet(order);
  return new Response("OK", { status: 200 });
};
