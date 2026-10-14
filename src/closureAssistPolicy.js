export function groupClosureCandidates(candidates) {
  const groups = new Map();

  for (const [index, candidate] of (Array.isArray(candidates) ? candidates : []).entries()) {
    if (!candidate || typeof candidate !== 'object') {
      const malformedKey = `malformed:${index}`;
      groups.set(malformedKey, { key: malformedKey, workOrderId: '', candidates: [] });
      continue;
    }
    const workOrderId = String(candidate.work_order_id || '').trim();
    const candidateId = String(candidate.id || '').trim();
    const key = workOrderId ? `wo:${workOrderId}` : `candidate:${candidateId || index}`;
    let group = groups.get(key);
    if (!group) {
      group = { key, workOrderId, candidates: [] };
      groups.set(key, group);
    }
    group.candidates.push(candidate);
  }

  return Array.from(groups.values());
}
