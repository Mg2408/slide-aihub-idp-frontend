import { getConfidenceDisplay, getExtractionViewModel } from "./extractionResponse";

const MONTH_FORMATTER = new Intl.DateTimeFormat("en-US", {
  month: "short",
  year: "2-digit",
});

const toDate = (value) => {
  if (!value) return null;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const isHighConfidence = (entry) => {
  const score =
    typeof entry?.confidenceScore === "number" && Number.isFinite(entry.confidenceScore)
      ? entry.confidenceScore
      : null;

  if (score !== null) {
    return score >= 0.8;
  }

  const confidence = getConfidenceDisplay(entry);
  return confidence?.color === "green";
};

const getSubmissionDate = (submission, fallbackDate) => {
  const candidate =
    submission?.last_modified ||
    submission?.created_at ||
    submission?.updated_at ||
    submission?.llm_response?.metadata?.processed_at ||
    submission?.llm_response?.metadata?.created_at;

  return toDate(candidate) || fallbackDate;
};

const aggregateSubmissionMetrics = (submission) => {
  const extractionResponse = submission?.llm_response || submission?.data || {};
  const { fieldGroups } = getExtractionViewModel(extractionResponse);

  let fieldCount = 0;
  let highConfidenceFields = 0;
  let confidenceSum = 0;
  let confidenceCount = 0;

  const fieldBreakdown = {};

  fieldGroups.forEach((group) => {
    if (!fieldBreakdown[group.fieldName]) {
      fieldBreakdown[group.fieldName] = {
        high: 0,
        review: 0,
      };
    }

    group.entries.forEach((entry) => {
      fieldCount += 1;

      if (
        typeof entry?.confidenceScore === "number" &&
        Number.isFinite(entry.confidenceScore)
      ) {
        confidenceSum += entry.confidenceScore;
        confidenceCount += 1;
      }

      if (isHighConfidence(entry)) {
        highConfidenceFields += 1;
        fieldBreakdown[group.fieldName].high += 1;
      } else {
        fieldBreakdown[group.fieldName].review += 1;
      }
    });
  });

  const averageConfidence = confidenceCount
    ? confidenceSum / confidenceCount
    : null;

  return {
    fieldCount,
    highConfidenceFields,
    averageConfidence,
    fieldBreakdown,
  };
};

export const buildDashboardMetrics = (submissions = []) => {
  const now = new Date();
  const monthlyMap = new Map();
  const fieldMap = new Map();

  let totalFields = 0;
  let totalHighConfidenceFields = 0;
  let confidenceWeightedSum = 0;
  let confidenceWeight = 0;

  submissions.forEach((submission) => {
    const submissionDate = getSubmissionDate(submission, now);
    const monthKey = `${submissionDate.getFullYear()}-${String(
      submissionDate.getMonth() + 1
    ).padStart(2, "0")}`;

    if (!monthlyMap.has(monthKey)) {
      monthlyMap.set(monthKey, {
        date: new Date(submissionDate.getFullYear(), submissionDate.getMonth(), 1),
        submitted: 0,
        converted: 0,
      });
    }

    const bucket = monthlyMap.get(monthKey);
    bucket.submitted += 1;

    const submissionMetrics = aggregateSubmissionMetrics(submission);
    totalFields += submissionMetrics.fieldCount;
    totalHighConfidenceFields += submissionMetrics.highConfidenceFields;

    if (submissionMetrics.averageConfidence !== null) {
      confidenceWeightedSum +=
        submissionMetrics.averageConfidence * submissionMetrics.fieldCount;
      confidenceWeight += submissionMetrics.fieldCount;

      if (submissionMetrics.averageConfidence >= 0.8) {
        bucket.converted += 1;
      }
    }

    Object.entries(submissionMetrics.fieldBreakdown).forEach(([fieldName, counts]) => {
      if (!fieldMap.has(fieldName)) {
        fieldMap.set(fieldName, { high: 0, review: 0, total: 0 });
      }

      const existing = fieldMap.get(fieldName);
      existing.high += counts.high;
      existing.review += counts.review;
      existing.total += counts.high + counts.review;
    });
  });

  const monthlyData = Array.from(monthlyMap.values())
    .sort((a, b) => a.date - b.date)
    .map((item) => ({
      month: MONTH_FORMATTER.format(item.date),
      submitted: item.submitted,
      converted: item.converted,
    }));

  const topFields = Array.from(fieldMap.entries())
    .sort((a, b) => b[1].total - a[1].total)
    .slice(0, 8);

  const categories = topFields.map(([fieldName]) => fieldName);
  const highConfidenceSeries = topFields.map(([, counts]) => counts.high);
  const reviewSeries = topFields.map(([, counts]) => counts.review);

  const averageConfidence =
    confidenceWeight > 0 ? (confidenceWeightedSum / confidenceWeight) * 100 : 0;

  const highConfidenceRate =
    totalFields > 0 ? (totalHighConfidenceFields / totalFields) * 100 : 0;

  return {
    summary: {
      totalDocuments: submissions.length,
      totalFields,
      averageConfidence,
      highConfidenceRate,
    },
    monthlyData,
    fieldChart: {
      categories,
      series: [
        { name: "High confidence", data: highConfidenceSeries },
        { name: "Needs review", data: reviewSeries },
      ],
    },
  };
};
