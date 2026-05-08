function parseTeamsUrl(url) {
  if (url.includes("/meet/") || url.includes("/meetup-join/")) {
    return url;
  }
  if (url.toLowerCase().includes("meetingrecap")) {
    const params = Object.fromEntries(
      [...url.matchAll(/([a-zA-Z]+)=([^&#]+)/g)].map((m) => [m[1], m[2]])
    );
    const threadId = decodeURIComponent(params.threadId || "");
    const tenantId = params.tenantId || "";
    const organizerId = params.organizerId || "";
    if (!threadId || !tenantId || !organizerId) {
      throw new Error("Invalid recap URL: missing threadId, tenantId, or organizerId parameter");
    }
    const threadEnc = encodeURIComponent(threadId).replace(/%3A/gi, "%3a").replace(/%40/gi, "%40");
    const ctx = JSON.stringify({ Tid: tenantId, Oid: organizerId });
    const ctxEnc = encodeURIComponent(ctx);
    return `https://teams.microsoft.com/l/meetup-join/${threadEnc}/0?context=${ctxEnc}`;
  }
  return url;
}
export {
  parseTeamsUrl
};
