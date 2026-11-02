import { Client } from "@microsoft/microsoft-graph-client";

export interface ListOptions {
  select?: string;
  top?: number;
  orderby?: string;
}

/**
 * GET a Graph collection and follow every @odata.nextLink so callers get the
 * whole collection, not just the first page (Graph returns 20 items by default).
 */
export async function getAll<T = any>(
  client: Client,
  uri: string,
  opts: ListOptions = {},
): Promise<T[]> {
  let request = client.api(uri);
  if (opts.select) request = request.select(opts.select);
  if (opts.top) request = request.top(opts.top);
  if (opts.orderby) request = request.orderby(opts.orderby);

  const results: T[] = [];
  let response = await request.get();
  results.push(...((response?.value ?? []) as T[]));

  let next: string | undefined = response?.["@odata.nextLink"];
  while (next) {
    response = await client.api(next).get();
    results.push(...((response?.value ?? []) as T[]));
    next = response?.["@odata.nextLink"];
  }
  return results;
}

/** Run fn over items with at most `limit` calls in flight (keeps Graph throttling at bay). */
export async function mapWithConcurrency<T, R>(
  items: T[],
  limit: number,
  fn: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const results: R[] = new Array(items.length);
  let nextIndex = 0;
  const workers = Array.from(
    { length: Math.max(1, Math.min(limit, items.length)) },
    async () => {
      while (nextIndex < items.length) {
        const i = nextIndex++;
        results[i] = await fn(items[i], i);
      }
    },
  );
  await Promise.all(workers);
  return results;
}

/** Graph returns page content as a string, a web ReadableStream, or binary; normalize to text. */
export async function readBodyAsText(body: unknown): Promise<string> {
  if (typeof body === "string") return body;
  if (body instanceof ReadableStream) {
    const reader = body.getReader();
    const decoder = new TextDecoder("utf-8");
    let result = "";
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      result += decoder.decode(value, { stream: true });
    }
    return result + decoder.decode();
  }
  if (body instanceof ArrayBuffer) return new TextDecoder("utf-8").decode(body);
  if (typeof Blob !== "undefined" && body instanceof Blob) return body.text();
  return `Unknown content type: ${typeof body}`;
}

const ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
  nbsp: " ",
};

/** Rough HTML -> plain text, good enough for searching and snippets. */
export function htmlToText(html: string): string {
  return html
    .replace(/<(script|style|head)[\s\S]*?<\/\1>/gi, " ")
    .replace(/<img[^>]*\balt="([^"]*)"[^>]*>/gi, " $1 ") // keep OCR'd image text
    .replace(/<(br|\/p|\/li|\/h\d|\/div|\/tr)[^>]*>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&(#x?[0-9a-f]+|\w+);/gi, (match, code: string) => {
      if (code[0] === "#") {
        const n =
          code[1] === "x" || code[1] === "X"
            ? parseInt(code.slice(2), 16)
            : parseInt(code.slice(1), 10);
        return Number.isFinite(n) ? String.fromCodePoint(n) : match;
      }
      return ENTITIES[code.toLowerCase()] ?? match;
    })
    .replace(/[ \t ]+/g, " ")
    .replace(/\s*\n\s*/g, "\n")
    .trim();
}

/** Up to `max` snippets of text around each case-insensitive match of query. */
export function findSnippets(
  text: string,
  query: string,
  max = 3,
  radius = 80,
): string[] {
  const haystack = text.toLowerCase();
  const needle = query.toLowerCase();
  const snippets: string[] = [];
  let from = 0;
  while (snippets.length < max) {
    const at = haystack.indexOf(needle, from);
    if (at < 0) break;
    const start = Math.max(0, at - radius);
    const end = Math.min(text.length, at + needle.length + radius);
    snippets.push(
      (start > 0 ? "…" : "") +
        text.slice(start, end).replace(/\s+/g, " ").trim() +
        (end < text.length ? "…" : ""),
    );
    from = end;
  }
  return snippets;
}
