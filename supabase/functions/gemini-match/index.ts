const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

const DEFAULT_MODEL = "gemini-2.5-flash";
const FALLBACK_MODELS = ["gemini-2.5-flash", "gemini-2.0-flash"];

const RESPONSE_SCHEMA = {
  type: "object",
  properties: {
    documents: {
      type: "array",
      items: {
        type: "object",
        properties: {
          id: { type: "string" },
          fileName: { type: "string" },
          type: { type: "string", enum: ["invoice", "payment"] },
          paymentType: { type: "string" },
          supplier: { type: "string" },
          supplierVat: { type: "string" },
          invoiceNumber: { type: "string" },
          invoiceDate: { type: "string" },
          dueDate: { type: "string" },
          taxable: { type: "number" },
          vat: { type: "number" },
          total: { type: "number" },
          currency: { type: "string" },
          bank: { type: "string" },
          payer: { type: "string" },
          payerIban: { type: "string" },
          beneficiary: { type: "string" },
          beneficiaryVat: { type: "string" },
          beneficiaryIban: { type: "string" },
          swift: { type: "string" },
          documentDate: { type: "string" },
          executionDate: { type: "string" },
          reason: { type: "string" },
          noticeNumber: { type: "string" },
          flowName: { type: "string" },
          confidence: { type: "number" },
          notes: { type: "string" },
        },
        required: ["id", "fileName", "type"],
      },
    },
    invoiceMatches: {
      type: "array",
      items: {
        type: "object",
        properties: {
          invoiceDocumentId: { type: "string" },
          paymentAllocations: {
            type: "array",
            items: {
              type: "object",
              properties: {
                paymentDocumentId: { type: "string" },
                allocatedAmount: { type: "number" },
              },
              required: ["paymentDocumentId", "allocatedAmount"],
            },
          },
          confidence: { type: "number" },
          rationale: { type: "string" },
        },
        required: ["invoiceDocumentId", "paymentAllocations", "confidence", "rationale"],
      },
    },
    existingRecordMatches: {
      type: "array",
      items: {
        type: "object",
        properties: {
          recordId: { type: "string" },
          paymentAllocations: {
            type: "array",
            items: {
              type: "object",
              properties: {
                paymentDocumentId: { type: "string" },
                allocatedAmount: { type: "number" },
              },
              required: ["paymentDocumentId", "allocatedAmount"],
            },
          },
          confidence: { type: "number" },
          rationale: { type: "string" },
        },
        required: ["recordId", "paymentAllocations", "confidence", "rationale"],
      },
    },
    duplicateInvoices: {
      type: "array",
      items: {
        type: "object",
        properties: {
          invoiceDocumentId: { type: "string" },
          existingRecordId: { type: "string" },
          rationale: { type: "string" },
        },
        required: ["invoiceDocumentId", "existingRecordId", "rationale"],
      },
    },
  },
  required: ["documents", "invoiceMatches", "existingRecordMatches", "duplicateInvoices"],
};

function shouldFallback(status: number, text = "") {
  return [429, 500, 503, 504].includes(status)
    || /high demand|try again later|overloaded|unavailable|quota/i.test(text);
}

function permissionProblem(status: number, text = "") {
  return status === 403
    && /api key was reported as leaked|permission_denied|api key not valid|forbidden/i.test(text);
}

function sleep(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

Deno.serve(async (request) => {
  if (request.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    const geminiApiKey = Deno.env.get("GEMINI_API_KEY");
    if (!geminiApiKey) {
      return new Response(JSON.stringify({ error: "Missing GEMINI_API_KEY secret" }), {
        status: 500,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const body = await request.json();
    const prompt = String(body?.prompt || "").trim();
    const documents = Array.isArray(body?.documents) ? body.documents : [];
    const requestedModel = String(body?.model || DEFAULT_MODEL).trim() || DEFAULT_MODEL;
    const models = [...new Set([requestedModel, ...FALLBACK_MODELS])];
    const parts = documents.map((document: { mimeType: string; data: string }) => ({
      inline_data: {
        mime_type: document.mimeType || "application/pdf",
        data: document.data,
      },
    }));
    parts.push({ text: prompt });

    let lastFailure: { status: number; text: string; model: string } | null = null;
    for (const model of models) {
      for (let attempt = 0; attempt < 2; attempt += 1) {
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(geminiApiKey)}`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            system_instruction: {
              parts: [{
                text: "Agisci come motore ufficiale di matching documentale. Decidi tu gli abbinamenti ufficiali tra fatture e pagamenti e restituisci solo JSON valido.",
              }],
            },
            contents: [{ role: "user", parts }],
            generationConfig: {
              responseMimeType: "application/json",
              responseSchema: RESPONSE_SCHEMA,
              temperature: 0.1,
            },
          }),
        });

        if (response.ok) {
          const json = await response.json();
          const text = json?.candidates?.[0]?.content?.parts?.map((part: { text?: string }) => part.text || "").join("").trim() || "";
          if (!text) {
            return new Response(JSON.stringify({ error: "Gemini returned empty content" }), {
              status: 502,
              headers: { ...corsHeaders, "Content-Type": "application/json" },
            });
          }
          const payload = JSON.parse(text);
          return new Response(JSON.stringify({ ...payload, modelUsed: model }), {
            status: 200,
            headers: { ...corsHeaders, "Content-Type": "application/json" },
          });
        }

        const text = await response.text();
        lastFailure = { status: response.status, text, model };
        if (permissionProblem(response.status, text)) {
          return new Response(text, {
            status: response.status,
            headers: { ...corsHeaders, "Content-Type": "application/json" },
          });
        }
        if (shouldFallback(response.status, text)) {
          if (attempt === 0) {
            await sleep(1200);
            continue;
          }
          break;
        }
        return new Response(text, {
          status: response.status,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        });
      }
    }

    return new Response(lastFailure?.text || JSON.stringify({ error: "Gemini temporarily unavailable" }), {
      status: lastFailure?.status || 503,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (error) {
    return new Response(JSON.stringify({ error: String(error?.message || error || "Unknown error") }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
