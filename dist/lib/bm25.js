function tokenize(text) {
  if (!text) return [];
  return text.replace(/([a-z0-9])([A-Z])/g, "$1 $2").toLowerCase().split(/[\s\-_/.,;:(){}[\]'"!?@#$]+/).filter((t) => t.length > 0);
}
function buildBM25Index(documents, k1 = 1.2, b = 0.75) {
  const docs = /* @__PURE__ */ new Map();
  const df = /* @__PURE__ */ new Map();
  let totalLen = 0;
  for (const { id, tokens } of documents) {
    const termFreq = /* @__PURE__ */ new Map();
    for (const tok of tokens) {
      termFreq.set(tok, (termFreq.get(tok) ?? 0) + 1);
    }
    for (const tok of termFreq.keys()) {
      df.set(tok, (df.get(tok) ?? 0) + 1);
    }
    docs.set(id, { id, length: tokens.length, termFreq });
    totalLen += tokens.length;
  }
  const N = docs.size;
  const avgdl = N > 0 ? totalLen / N : 0;
  const idf = /* @__PURE__ */ new Map();
  for (const [term, n] of df) {
    idf.set(term, Math.log((N - n + 0.5) / (n + 0.5) + 1));
  }
  return { docs, idf, avgdl, k1, b };
}
function scoreQuery(query, index) {
  const queryTokens = [...new Set(tokenize(query))];
  if (queryTokens.length === 0) return [];
  const results = [];
  for (const [id, doc] of index.docs) {
    let score = 0;
    let matched = false;
    for (const qt of queryTokens) {
      const tf = doc.termFreq.get(qt);
      if (!tf) continue;
      matched = true;
      const idf = index.idf.get(qt) ?? 0;
      const num = tf * (index.k1 + 1);
      const den = tf + index.k1 * (1 - index.b + index.b * doc.length / (index.avgdl || 1));
      score += idf * (num / den);
    }
    if (matched) results.push({ id, score });
  }
  results.sort((a, b) => b.score - a.score);
  return results;
}
export {
  buildBM25Index,
  scoreQuery,
  tokenize
};
