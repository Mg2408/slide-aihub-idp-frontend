const DIRECT_VALUE_KEYS = [
  "valueString",
  "valueNumber",
  "valueInteger",
  "valueBoolean",
  "valueDate",
  "valueTime",
  "valuePhoneNumber",
  "valueCountryRegion",
  "valueSelectionMark",
  "valueCurrency",
  "valueSignature",
  "content",
];

const STATUS_PRIORITY = {
  high_confidence: 4,
  high: 3,
  medium_confidence: 2,
  medium: 2,
  low_confidence: 1,
  low: 1,
};

const isRecord = (value) =>
  value !== null && typeof value === "object" && !Array.isArray(value);

const toArray = (value) => {
  if (Array.isArray(value)) return value;
  if (value === null || value === undefined) return [];
  return [value];
};

const stringifyValue = (value) => {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

  try {
    return JSON.stringify(value);
  } catch (error) {
    return "";
  }
};

const pickIndexed = (value, index = 0) =>
  Array.isArray(value) ? value[index] ?? value[0] : value;

const normalizeReferenceValue = (value) => {
  if (Array.isArray(value)) {
    return value
      .map((item) => stringifyValue(item).trim())
      .filter(Boolean)
      .join(", ");
  }

  return stringifyValue(value).trim();
};

const parseNumericConfidence = (value) => {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === "string" && value.trim()) {
    const parsed = Number.parseFloat(value);
    if (!Number.isNaN(parsed)) {
      return parsed;
    }
  }

  return null;
};

const normalizeStatus = (value) => {
  if (typeof value !== "string" || !value.trim()) {
    return null;
  }

  return value.trim().toLowerCase().replace(/_/g, " ");
};

const getStatusRank = (status) =>
  STATUS_PRIORITY[status?.replace(/\s+/g, "_")] ?? 0;

const pickBestStatus = (statuses = []) =>
  statuses.reduce((best, current) => {
    if (!current) return best;
    if (!best || getStatusRank(current) > getStatusRank(best)) {
      return current;
    }
    return best;
  }, null);

const getNodeReference = (node, index = 0) => {
  if (!isRecord(node)) {
    return null;
  }

  const pageValue = normalizeReferenceValue(pickIndexed(node.page, index));
  if (pageValue) {
    return { label: "Page", value: pageValue };
  }

  const sourceValue = normalizeReferenceValue(pickIndexed(node.source, index));
  if (sourceValue) {
    return { label: "Source", value: sourceValue };
  }

  return null;
};

const getNodeConfidenceScore = (node, index = 0) => {
  if (!isRecord(node)) {
    return null;
  }

  return (
    parseNumericConfidence(pickIndexed(node.confidence_score, index)) ??
    parseNumericConfidence(pickIndexed(node.confidence, index)) ??
    parseNumericConfidence(pickIndexed(node.original_confidence, index))
  );
};

const getNodeStatus = (node, index = 0) => {
  if (!isRecord(node)) {
    return null;
  }

  return (
    normalizeStatus(pickIndexed(node.confidence_level, index)) ??
    normalizeStatus(pickIndexed(node.status, index))
  );
};

const getNodeObject = (node) => {
  if (isRecord(node?.valueObject)) {
    return node.valueObject;
  }

  if (isRecord(node?.value)) {
    return node.value;
  }

  if (isRecord(node)) {
    return node;
  }

  return {};
};

const getNodeArray = (node) => {
  if (Array.isArray(node?.valueArray)) {
    return node.valueArray;
  }

  if (Array.isArray(node?.value)) {
    return node.value;
  }

  if (Array.isArray(node)) {
    return node;
  }

  return [];
};

const extractDisplayValue = (node) => {
  if (node === null || node === undefined) {
    return "";
  }

  if (typeof node === "string") {
    return node;
  }

  if (typeof node === "number" || typeof node === "boolean") {
    return String(node);
  }

  if (Array.isArray(node)) {
    return node.map(extractDisplayValue).filter(Boolean).join("\n\n");
  }

  if (!isRecord(node)) {
    return stringifyValue(node);
  }

  if (Object.prototype.hasOwnProperty.call(node, "value")) {
    return extractDisplayValue(node.value);
  }

  for (const key of DIRECT_VALUE_KEYS) {
    if (
      Object.prototype.hasOwnProperty.call(node, key) &&
      !isRecord(node[key]) &&
      !Array.isArray(node[key])
    ) {
      return stringifyValue(node[key]);
    }
  }

  if (Array.isArray(node.valueArray)) {
    return node.valueArray.map(extractDisplayValue).filter(Boolean).join("\n\n");
  }

  if (isRecord(node.valueObject)) {
    return Object.entries(node.valueObject)
      .map(([key, value]) => {
        const displayValue = extractDisplayValue(value);
        return displayValue ? `${formatFieldName(key)}: ${displayValue}` : "";
      })
      .filter(Boolean)
      .join(" | ");
  }

  return stringifyValue(node);
};

const hasMeaningfulEntry = (entry) =>
  Boolean(
    entry.value ||
      entry.reference ||
      entry.confidenceScore !== null ||
      entry.confidenceStatus
  );

const createEntry = (node, overrides = {}) => {
  const derivedReference = getNodeReference(node);

  return {
    value: overrides.value ?? extractDisplayValue(node),
    confidenceScore:
      overrides.confidenceScore ?? getNodeConfidenceScore(node),
    confidenceStatus:
      overrides.confidenceStatus ?? getNodeStatus(node),
    reference: overrides.reference ?? derivedReference?.value ?? "",
    referenceLabel:
      overrides.referenceLabel ?? derivedReference?.label ?? null,
  };
};

const normalizeLegacyFieldEntries = (fieldNode) => {
  const rawValues = Array.isArray(fieldNode?.value)
    ? fieldNode.value
    : [fieldNode?.value];

  return rawValues
    .map((value, index) => {
      const reference = getNodeReference(fieldNode, index);

      return {
        value: extractDisplayValue(value),
        confidenceScore: getNodeConfidenceScore(fieldNode, index),
        confidenceStatus: getNodeStatus(fieldNode, index),
        reference: reference?.value ?? "",
        referenceLabel: reference?.label ?? null,
      };
    })
    .filter(hasMeaningfulEntry);
};

const aggregateNestedEntries = (nodes, value) => {
  const validNodes = nodes.filter(Boolean);
  const scores = validNodes
    .map((node) => getNodeConfidenceScore(node))
    .filter((score) => score !== null);
  const statuses = validNodes
    .map((node) => getNodeStatus(node))
    .filter(Boolean);
  const reference = validNodes.map((node) => getNodeReference(node)).find(Boolean);

  return {
    value,
    confidenceScore: scores.length ? Math.max(...scores) : null,
    confidenceStatus: pickBestStatus(statuses),
    reference: reference?.value ?? "",
    referenceLabel: reference?.label ?? null,
  };
};

const normalizePolicyEntries = (fieldNode) => {
  const policyNodes = getNodeArray(fieldNode);
  const policies = policyNodes.length ? policyNodes : toArray(fieldNode);

  return policies
    .flatMap((policyNode) => {
      const policyObject = getNodeObject(policyNode);
      const policyNumberNode = policyObject?.policy_number;
      const borrowers = getNodeArray(policyObject?.borrowers);
      const policyNumber = extractDisplayValue(policyNumberNode);

      if (!borrowers.length) {
        const fallbackValue =
          policyNumber ? `Policy Number: ${policyNumber}` : extractDisplayValue(policyNode);

        return [aggregateNestedEntries([policyNumberNode || policyNode], fallbackValue)];
      }

      return borrowers.map((borrowerNode) => {
        const borrowerObject = getNodeObject(borrowerNode);
        const nameNode = borrowerObject?.name;
        const addressNode = borrowerObject?.address;

        const combinedValue = [
          policyNumber ? `Policy Number: ${policyNumber}` : "",
          extractDisplayValue(nameNode)
            ? `Borrower: ${extractDisplayValue(nameNode)}`
            : "",
          extractDisplayValue(addressNode)
            ? `Address: ${extractDisplayValue(addressNode)}`
            : "",
        ]
          .filter(Boolean)
          .join(" | ");

        return aggregateNestedEntries(
          [policyNumberNode, nameNode, addressNode],
          combinedValue || extractDisplayValue(borrowerNode)
        );
      });
    })
    .filter(hasMeaningfulEntry);
};

const normalizeFieldEntries = (fieldName, fieldNode) => {
  if (formatFieldName(fieldName).toLowerCase() === "policies") {
    return normalizePolicyEntries(fieldNode);
  }

  const isLegacyField =
    isRecord(fieldNode) &&
    ("value" in fieldNode ||
      "confidence_score" in fieldNode ||
      "confidence_level" in fieldNode ||
      "page" in fieldNode);

  if (isLegacyField) {
    return normalizeLegacyFieldEntries(fieldNode);
  }

  if (Array.isArray(fieldNode)) {
    return fieldNode.map((entry) => createEntry(entry)).filter(hasMeaningfulEntry);
  }

  if (Array.isArray(fieldNode?.valueArray)) {
    return fieldNode.valueArray
      .map((entry) => createEntry(entry))
      .filter(hasMeaningfulEntry);
  }

  return [createEntry(fieldNode)].filter(hasMeaningfulEntry);
};

const getRawMetadata = (rawResponse) =>
  isRecord(rawResponse?.metadata) ? rawResponse.metadata : {};

const getRawFields = (rawResponse) => {
  if (!isRecord(rawResponse)) {
    return {};
  }

  if (isRecord(rawResponse.fields)) {
    return rawResponse.fields;
  }

  return Object.fromEntries(
    Object.entries(rawResponse).filter(([key]) => key !== "metadata")
  );
};

export const formatFieldName = (name) =>
  String(name ?? "").replace(/_/g, " ").trim();

export const getExtractionViewModel = (rawResponse) => {
  const metadata = getRawMetadata(rawResponse);
  const fields = getRawFields(rawResponse);

  return {
    metadata,
    metadataRows: Object.entries(metadata).map(([key, value], index) => ({
      key: index,
      keyName: formatFieldName(key).toUpperCase(),
      value: stringifyValue(value),
    })),
    fieldGroups: Object.entries(fields).map(([fieldName, fieldNode], index) => ({
      id: `field-${index}`,
      rawName: fieldName,
      fieldName: formatFieldName(fieldName),
      entries: normalizeFieldEntries(fieldName, fieldNode),
    })),
  };
};

export const buildFlatFieldRows = (rawResponse) => {
  const { fieldGroups } = getExtractionViewModel(rawResponse);

  return fieldGroups.flatMap((group) => {
    if (!group.entries.length) {
      return [
        {
          Field: group.fieldName,
          Value: "",
          Reference: "",
        },
      ];
    }

    return group.entries.map((entry, index) => ({
      Field: index === 0 ? group.fieldName : "",
      Value: entry.value,
      Reference: entry.reference,
    }));
  });
};

export const buildMatrixPreviewRows = (rawResponse, documentName) => {
  const { fieldGroups } = getExtractionViewModel(rawResponse);
  const maxRows = Math.max(
    1,
    ...fieldGroups.map((group) => Math.max(group.entries.length, 1))
  );

  return Array.from({ length: maxRows }, (_, rowIndex) => {
    const row = {
      "Document Name": rowIndex === 0 ? documentName : "",
    };

    fieldGroups.forEach((group) => {
      row[group.fieldName] = group.entries[rowIndex]?.value ?? "";
    });

    row.Reference =
      fieldGroups
        .map((group) => group.entries[rowIndex]?.reference)
        .find(Boolean) ?? "";

    return row;
  });
};

export const getDocumentName = (rawResponse, fallbacks = []) => {
  const metadata = getRawMetadata(rawResponse);

  return (
    metadata.document_name ||
    metadata.file_name ||
    fallbacks.find((value) => typeof value === "string" && value.trim()) ||
    "extracted"
  );
};

export const getOwnerName = (rawResponse, fallbacks = []) => {
  const metadata = getRawMetadata(rawResponse);

  return (
    metadata.owner_name ||
    metadata.submitted_by ||
    metadata.owner ||
    fallbacks.find((value) => typeof value === "string" && value.trim()) ||
    "—"
  );
};

export const getConfidenceDisplay = (entry = {}) => {
  const numericScore = parseNumericConfidence(entry.confidenceScore);
  if (numericScore !== null) {
    return {
      value: `${Math.round(numericScore * 100)}%`,
      color: numericScore > 0.8 ? "green" : numericScore > 0.5 ? "orange" : "red",
    };
  }

  const status = normalizeStatus(entry.confidenceStatus);
  if (!status) {
    return null;
  }

  return {
    value: status.charAt(0).toUpperCase() + status.slice(1),
    color:
      status.includes("high")
        ? "green"
        : status.includes("medium")
          ? "orange"
          : status.includes("low")
            ? "red"
            : "blue",
  };
};

export const getReferenceDisplay = (entry = {}) => {
  if (!entry.reference) {
    return null;
  }

  return {
    label: entry.referenceLabel || "Reference",
    value: entry.reference,
  };
};
