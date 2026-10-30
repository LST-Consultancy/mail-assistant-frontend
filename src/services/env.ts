// Values are inlined at build time by webpack's DefinePlugin.
export const env = {
  clientId: process.env.VITE_ENTRA_CLIENT_ID ?? "",
  apiBaseUrl: (process.env.VITE_API_BASE_URL || "https://localhost:8080").replace(/\/$/, ""),
  maxEmails: Number(process.env.VITE_MAX_EMAILS || 300),
  maxBodyChars: Number(process.env.VITE_MAX_BODY_CHARS || 6000),
  mock: process.env.VITE_MOCK === "1",
};
