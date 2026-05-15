const DEPRECATED_POLICY_KEYS = new Set(['Electricity Duty - Wind farm']);

const filterActivePolicyItems = (items) => {
  if (!Array.isArray(items)) return [];
  return items.filter((item) => item?.key && !DEPRECATED_POLICY_KEYS.has(item.key));
};

const sanitizePolicyDocument = (policy) => {
  if (!policy) return policy;
  const doc = typeof policy.toObject === 'function' ? policy.toObject() : { ...policy };
  return {
    ...doc,
    policies: filterActivePolicyItems(doc.policies),
  };
};

const sanitizeClientPolicyDocument = (clientPolicy) => {
  if (!clientPolicy) return clientPolicy;
  const doc =
    typeof clientPolicy.toObject === 'function' ? clientPolicy.toObject() : { ...clientPolicy };
  if (doc.policyId) {
    return {
      ...doc,
      policyId: sanitizePolicyDocument(doc.policyId),
    };
  }
  return doc;
};

module.exports = {
  DEPRECATED_POLICY_KEYS,
  filterActivePolicyItems,
  sanitizePolicyDocument,
  sanitizeClientPolicyDocument,
};
