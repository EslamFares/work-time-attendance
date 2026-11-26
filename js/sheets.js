import { APP_SIGNATURE } from "./calc.js";

export async function pingWebhook(url) {
  const pingUrl = new URL(url);
  pingUrl.searchParams.set("ping", "1");
  const response = await fetch(pingUrl.toString(), { method: "GET" });
  const data = JSON.parse(await response.text());
  if (!data || data.ok !== true || data.app !== APP_SIGNATURE) {
    throw new Error("webhook");
  }
  return data;
}

export async function pullWebhook(url) {
  const pullUrl = new URL(url);
  pullUrl.searchParams.set("pull", "1");
  const response = await fetch(pullUrl.toString(), { method: "GET" });
  const data = JSON.parse(await response.text());
  if (!data || data.ok !== true || data.app !== APP_SIGNATURE) {
    throw new Error(data?.error || "pull");
  }
  return data;
}

export async function pushWebhook(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });
  const data = JSON.parse(await response.text());
  if (!data || data.ok !== true) throw new Error(data?.error || "sync");
  return data;
}
