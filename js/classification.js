// js/classification.js
// Wraps the Azure Function that proxies Gemini for email classification +
// quote-amount + AI summary calls. The Function holds the Gemini API key
// server-side; the browser only ever talks to our Function URL.
//
// CONFIG.classifierUrl points to the Function. If it's empty, we fall back
// to a local stub that returns "Question" for every email — keeps the rest
// of the app working until the Function is configured.

import { CONFIG } from './config.js';

// The Function expects POST { task, payload } and returns { result }.
async function callClassifier(task, payload) {
  if (!CONFIG.classifierUrl) {
    // Stub: until the Azure Function is configured, return safe fallbacks.
    return stubFallback(task, payload);
  }
  const res = await fetch(CONFIG.classifierUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ task, payload })
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Classifier ${res.status}: ${text.slice(0, 200)}`);
  }
  const json = await res.json();
  return json.result;
}

function stubFallback(task, payload) {
  // Conservative fallbacks so absence of classifier doesn't hide messages.
  if (task === 'classify') return { classification: 'Question', confidence: 0 };
  if (task === 'extractAmount') return { amount: null, currency: 'AUD', notes: 'classifier offline' };
  if (task === 'summarizeFilename') {
    const name = (payload && payload.originalName) || 'document';
    return { summary: name.replace(/\.[^.]+$/, '').slice(0, 30) };
  }
  return null;
}

// Classify an email. Returns one of:
//   Quote / Question / Suspicious / Out-of-Office / Decline / Unrelated
export async function classifyEmail({ subject, fromName, fromEmail, bodyText }) {
  const result = await callClassifier('classify', {
    subject: subject || '',
    fromName: fromName || '',
    fromEmail: fromEmail || '',
    bodyText: (bodyText || '').slice(0, 5000)  // cap to keep token cost bounded
  });
  return result || { classification: 'Question', confidence: 0 };
}

// Extract a quote amount from email body text. Returns
//   { amount: 58500, currency: 'AUD' }   on success
//   { amount: null }                     when no amount found
export async function extractQuoteAmount({ subject, bodyText, attachmentText }) {
  const result = await callClassifier('extractAmount', {
    subject: subject || '',
    bodyText: (bodyText || '').slice(0, 5000),
    attachmentText: (attachmentText || '').slice(0, 5000)
  });
  return result || { amount: null };
}

// Generate a 2-3 word filename summary from a PDF attachment's content.
// Returns { summary: "Site-Plan" }
export async function summarizeFilename({ originalName, attachmentText }) {
  const result = await callClassifier('summarizeFilename', {
    originalName: originalName || '',
    attachmentText: (attachmentText || '').slice(0, 3000)
  });
  return result || { summary: 'document' };
}
