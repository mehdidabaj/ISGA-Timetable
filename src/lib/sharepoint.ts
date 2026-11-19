const DEFAULT_SHARE_URL =
  "https://isgacloud-my.sharepoint.com/:x:/r/personal/rachid_ettaibi_isga_ma/_layouts/15/Doc.aspx?sourcedoc=%7B64769b21-cb6f-4e08-a6ec-7542e5e21f0d%7D&action=default";

function encodeShareId(shareUrl: string): string {
  const b64 = Buffer.from(shareUrl, "utf8")
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
  return `u!${b64}`;
}

async function getGraphToken(): Promise<string | null> {
  const tenant = process.env.MS_TENANT_ID?.trim();
  const clientId = process.env.MS_CLIENT_ID?.trim();
  const clientSecret = process.env.MS_CLIENT_SECRET?.trim();
  if (!tenant || !clientId || !clientSecret) return null;

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });

  const res = await fetch(
    `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
      cache: "no-store",
    },
  );

  if (!res.ok) {
    const detail = await res.text();
    throw new Error(`Échec du token Microsoft Graph (${res.status}): ${detail.slice(0, 200)}`);
  }

  const json = (await res.json()) as { access_token?: string };
  if (!json.access_token) throw new Error("Token Microsoft Graph manquant dans la réponse.");
  return json.access_token;
}

async function downloadViaGraph(shareUrl: string, token: string): Promise<Buffer> {
  const shareId = encodeShareId(shareUrl);
  const res = await fetch(
    `https://graph.microsoft.com/v1.0/shares/${encodeURIComponent(shareId)}/driveItem/content`,
    {
      headers: { Authorization: `Bearer ${token}` },
      redirect: "follow",
      cache: "no-store",
    },
  );

  if (!res.ok) {
    const detail = await res.text();
    throw new Error(`Téléchargement Graph échoué (${res.status}): ${detail.slice(0, 240)}`);
  }

  return Buffer.from(await res.arrayBuffer());
}

async function downloadViaAnonymous(shareUrl: string): Promise<Buffer | null> {
  const candidates = [
    shareUrl.includes("download=") ? shareUrl : `${shareUrl}${shareUrl.includes("?") ? "&" : "?"}download=1`,
    shareUrl.replace("action=default", "action=download"),
  ];

  for (const url of candidates) {
    try {
      const res = await fetch(url, {
        redirect: "follow",
        cache: "no-store",
        headers: {
          Accept:
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
        },
      });
      if (!res.ok) continue;
      const contentType = res.headers.get("content-type") || "";
      const buf = Buffer.from(await res.arrayBuffer());
      // Reject HTML login pages.
      if (contentType.includes("text/html")) continue;
      if (buf.length < 100) continue;
      // XLSX files start with PK (zip).
      if (buf[0] === 0x50 && buf[1] === 0x4b) return buf;
    } catch {
      // try next candidate
    }
  }
  return null;
}

export function resolveShareUrl(): string {
  return process.env.SHAREPOINT_SHARE_URL?.trim() || DEFAULT_SHARE_URL;
}

export async function fetchSharePointExcel(): Promise<{
  buffer: Buffer;
  mode: "graph" | "anonymous";
}> {
  const shareUrl = resolveShareUrl();
  const token = await getGraphToken();

  if (token) {
    const buffer = await downloadViaGraph(shareUrl, token);
    return { buffer, mode: "graph" };
  }

  const anonymous = await downloadViaAnonymous(shareUrl);
  if (anonymous) return { buffer: anonymous, mode: "anonymous" };

  throw new Error(
    "Impossible d'accéder au fichier SharePoint. Ajoutez un lien de partage public (Anyone with the link) ou les identifiants Microsoft Graph.",
  );
}
