import { buildDatasetFromUpload } from "./datasetUpload.js";

self.onmessage = (event) => {
  const { text, name } = event.data || {};
  try {
    const data = buildDatasetFromUpload(text, { name });
    self.postMessage({ ok: true, data });
  } catch (error) {
    self.postMessage({ ok: false, error: error.message || "Could not analyze that dataset." });
  }
};
