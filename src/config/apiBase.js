const rawAiExtractBaseUrl = import.meta.env.VITE_AI_EXTRACT?.trim() ?? "";

const normalizedAiExtractBaseUrl = rawAiExtractBaseUrl.replace(/\/+$/, "");

export const buildAiExtractUrl = (path = "") => {
  if (!normalizedAiExtractBaseUrl) {
    return path;
  }

  if (!path) {
    return normalizedAiExtractBaseUrl;
  }

  return `${normalizedAiExtractBaseUrl}${path.startsWith("/") ? path : `/${path}`}`;
};
