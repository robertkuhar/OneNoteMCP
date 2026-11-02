import { Client } from "@microsoft/microsoft-graph-client";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import z from "zod";
import { Page } from "../types";
import { getErrorMessage } from "../utils/error";
import {
  findSnippets,
  getAll,
  htmlToText,
  mapWithConcurrency,
  readBodyAsText,
} from "../utils/graph";

const PAGE_FIELDS = "id,title,createdDateTime,lastModifiedDateTime,contentUrl";
const CONCURRENCY = 4;
const CONTENT_CACHE_TTL_MS = 10 * 60 * 1000;

interface SectionRef {
  id: string;
  name: string;
  notebook: string;
  path: string; // e.g. "Robert's Notebook / Group / 2026 Notes"
}

interface SearchHit {
  id: string;
  title: string;
  createdTime: string;
  lastModifiedTime: string;
  notebook: string;
  section: string;
  sectionPath: string;
  matchedIn: ("title" | "content")[];
  snippets?: string[];
}

// Page text cache so repeated content searches don't refetch every page.
// Keyed by page id only, with a short TTL, because OneNote's
// lastModifiedDateTime is not reliably bumped when a page is edited.
const contentCache = new Map<string, { text: string; at: number }>();

async function fetchPageText(azureClient: Client, id: string): Promise<string> {
  const cached = contentCache.get(id);
  if (cached && Date.now() - cached.at < CONTENT_CACHE_TTL_MS) return cached.text;
  const body = await azureClient.api(`/me/onenote/pages/${id}/content`).get();
  const text = htmlToText(await readBodyAsText(body));
  contentCache.set(id, { text, at: Date.now() });
  return text;
}

/** All sections in a notebook, including ones nested in section groups (any depth). */
async function collectNotebookSections(
  azureClient: Client,
  notebookId: string,
): Promise<SectionRef[]> {
  const notebook = await azureClient
    .api(`/me/onenote/notebooks/${notebookId}`)
    .select("id,displayName")
    .get();
  const notebookName: string = notebook.displayName;
  const found: SectionRef[] = [];

  async function walk(base: string, path: string): Promise<void> {
    const [sections, groups] = await Promise.all([
      getAll(azureClient, `${base}/sections`, { select: "id,displayName" }),
      getAll(azureClient, `${base}/sectionGroups`, { select: "id,displayName" }),
    ]);
    for (const s of sections) {
      found.push({
        id: s.id,
        name: s.displayName,
        notebook: notebookName,
        path: `${path} / ${s.displayName}`,
      });
    }
    for (const g of groups) {
      await walk(`/me/onenote/sectionGroups/${g.id}`, `${path} / ${g.displayName}`);
    }
  }

  await walk(`/me/onenote/notebooks/${notebookId}`, notebookName);
  return found;
}

async function resolveScope(
  azureClient: Client,
  notebookId?: string,
  sectionId?: string,
): Promise<SectionRef[]> {
  if (sectionId) {
    const s = await azureClient
      .api(`/me/onenote/sections/${sectionId}`)
      .expand("parentNotebook($select=displayName)")
      .select("id,displayName")
      .get();
    const notebook = s.parentNotebook?.displayName ?? "";
    return [
      { id: s.id, name: s.displayName, notebook, path: `${notebook} / ${s.displayName}` },
    ];
  }
  if (notebookId) return collectNotebookSections(azureClient, notebookId);

  const notebooks = await getAll(azureClient, "/me/onenote/notebooks", {
    select: "id,displayName",
  });
  const perNotebook = await mapWithConcurrency(notebooks, 2, (nb: any) =>
    collectNotebookSections(azureClient, nb.id),
  );
  return perNotebook.flat();
}

export function registerPageTools(mcpServer: McpServer, azureClient: Client) {
  mcpServer.tool(
    "listPages",
    "List pages within a specific OneNote section",
    {
      sectionId: z
        .string()
        .describe("The ID of the section containing the pages"),
    },
    async (args, extra): Promise<{
      content: {
        type: "resource";
        resource: { mimeType: string; text: string; uri: string }; // Added uri
      }[];
    }> => {
      const { sectionId } = args;
      const uri = `/me/onenote/sections/${sectionId}/pages`; // Define uri
      try {
        // Follow @odata.nextLink so sections with more than 20 pages come back complete
        const allPages = await getAll(azureClient, uri, {
          select: PAGE_FIELDS,
          top: 100,
        });

        const pages: Page[] = allPages.map((page: any) => ({
          id: page.id,
          title: page.title,
          createdTime: page.createdDateTime,
          lastModifiedTime: page.lastModifiedDateTime,
          contentUrl: page.contentUrl,
        }));
        return {
          content: [
            {
              type: "resource",
              resource: {
                mimeType: "application/json",
                text: JSON.stringify(pages),
                uri: uri, // Add uri to response
              },
            },
          ],
        };
      } catch (error) {
        throw new Error(
          `Failed to list pages in section ${sectionId}: ${getErrorMessage(
            error,
          )}`, // Use getErrorMessage
        );
      }
    },
  );

  mcpServer.tool(
    "searchPages",
    "Search OneNote pages by title and, optionally, by page text. Scope is one section, one whole " +
      "notebook (every section, including those inside section groups), or all notebooks when " +
      "neither ID is given. Matching is case-insensitive substring. Title-only search is fast; " +
      "searchContent=true reads every page in scope (cached for 10 minutes), so prefer narrowing " +
      "to a notebook or section for content searches.",
    {
      query: z.string().min(1).describe("Text to look for (case-insensitive substring)"),
      notebookId: z
        .string()
        .optional()
        .describe("Search every section of this notebook, including section groups"),
      sectionId: z
        .string()
        .optional()
        .describe("Search only this section (takes precedence over notebookId)"),
      searchContent: z
        .boolean()
        .optional()
        .describe("Also search the body text of each page and return snippets (default false)"),
      maxResults: z
        .number()
        .int()
        .positive()
        .optional()
        .describe("Maximum number of matching pages to return (default 100)"),
    },
    async (args, extra): Promise<{
      content: {
        type: "resource";
        resource: { mimeType: string; text: string; uri: string };
      }[];
    }> => {
      const { query, notebookId, sectionId } = args;
      const searchContent = args.searchContent ?? false;
      const maxResults = args.maxResults ?? 100;
      const needle = query.toLowerCase();
      const uri = sectionId
        ? `/me/onenote/sections/${sectionId}/pages`
        : notebookId
          ? `/me/onenote/notebooks/${notebookId}`
          : "/me/onenote/notebooks";

      try {
        const sections = await resolveScope(azureClient, notebookId, sectionId);
        const errors: string[] = [];

        // Every page in every section in scope, paged past Graph's 20-item default
        const pagesBySection = await mapWithConcurrency(
          sections,
          CONCURRENCY,
          async (section) => {
            try {
              const pages = await getAll(
                azureClient,
                `/me/onenote/sections/${section.id}/pages`,
                { select: PAGE_FIELDS, top: 100 },
              );
              return pages.map((page: any) => ({ page, section }));
            } catch (error) {
              errors.push(`Section "${section.path}": ${getErrorMessage(error)}`);
              return [];
            }
          },
        );
        const candidates = pagesBySection.flat();

        let pagesRead = 0;
        const hits = await mapWithConcurrency(
          candidates,
          CONCURRENCY,
          async ({ page, section }): Promise<SearchHit | undefined> => {
            const matchedIn: SearchHit["matchedIn"] = [];
            let snippets: string[] | undefined;
            if ((page.title ?? "").toLowerCase().includes(needle)) matchedIn.push("title");
            if (searchContent) {
              try {
                const text = await fetchPageText(azureClient, page.id);
                pagesRead++;
                const found = findSnippets(text, query);
                if (found.length > 0) {
                  matchedIn.push("content");
                  snippets = found;
                }
              } catch (error) {
                errors.push(`Page "${page.title}": ${getErrorMessage(error)}`);
              }
            }
            if (matchedIn.length === 0) return undefined;
            return {
              id: page.id,
              title: page.title,
              createdTime: page.createdDateTime,
              lastModifiedTime: page.lastModifiedDateTime,
              notebook: section.notebook,
              section: section.name,
              sectionPath: section.path,
              matchedIn,
              snippets,
            };
          },
        );

        const matches = hits
          .filter((h): h is SearchHit => h !== undefined)
          .sort((a, b) => (b.createdTime ?? "").localeCompare(a.createdTime ?? ""));

        const result = {
          query,
          searchContent,
          stats: {
            sectionsSearched: sections.length,
            pagesScanned: candidates.length,
            pagesRead,
            matches: matches.length,
            returned: Math.min(matches.length, maxResults),
          },
          results: matches.slice(0, maxResults),
          errors: errors.length > 0 ? errors.slice(0, 20) : undefined,
        };

        return {
          content: [
            {
              type: "resource",
              resource: {
                mimeType: "application/json",
                text: JSON.stringify(result),
                uri,
              },
            },
          ],
        };
      } catch (error) {
        throw new Error(`Failed to search pages: ${getErrorMessage(error)}`);
      }
    },
  );

  mcpServer.tool(
    "createPage",
    "Create a new OneNote page within a section",
    {
      title: z.string().describe("The title for the new page"),
      content: z.string().describe("The content of the page in HTML format"),
      sectionId: z
        .string()
        .describe("The ID of the section where the page will be created"),
    },
    async (args, extra): Promise<{
      content: {
        type: "resource";
        resource: { mimeType: string; text: string; uri: string }; // Added uri
      }[];
    }> => {
      const { title, content, sectionId } = args;
      const baseUri = `/me/onenote/sections/${sectionId}/pages`; // Define base uri
      try {
        const htmlContent = `<!DOCTYPE html>
            <html>
              <head>
                <title>${title}</title>
              </head>
              <body>
                ${content}
              </body>
            </html>`;

        const page = await azureClient
          .api(baseUri) // Use base uri
          .header("Content-Type", "application/xhtml+xml")
          .post(htmlContent);

        const createdPage: Page = {
          id: page.id,
          title: page.title,
          createdTime: page.createdDateTime,
          lastModifiedTime: page.lastModifiedDateTime,
          contentUrl: page.contentUrl,
        };
        const pageUri = `/me/onenote/pages/${page.id}`; // Define specific page uri
        return {
          content: [
            {
              type: "resource",
              resource: {
                mimeType: "application/json",
                text: JSON.stringify(createdPage),
                uri: pageUri, // Add uri to response
              },
            },
          ],
        };
      } catch (error) {
        throw new Error(
          `Failed to create page in section ${sectionId}: ${getErrorMessage(
            error,
          )}`, // Use getErrorMessage
        );
      }
    },
  );

  mcpServer.tool(
    "getPageContent",
    "Get the content of a specific OneNote page",
    {
      id: z.string().describe("The ID of the page to retrieve content for"),
    },
    async (args, extra): Promise<{
      content: {
        type: "resource";
        resource: { mimeType: string; text: string; uri: string }; // Added uri
      }[];
    }> => {
      const { id } = args;
      const uri = `/me/onenote/pages/${id}/content`; // Define uri
      try {
        const pageMeta = await azureClient
          .api(`/me/onenote/pages/${id}`)
          .select("id,title,createdDateTime,lastModifiedDateTime,contentUrl")
          .get();

        const contentStream = await azureClient
          .api(uri) // Use uri
          .get();

        const pageContent = await readBodyAsText(contentStream);

        const resultPage: Page = {
          id: pageMeta.id,
          title: pageMeta.title,
          createdTime: pageMeta.createdDateTime,
          lastModifiedTime: pageMeta.lastModifiedDateTime,
          content: pageContent,
          contentUrl: pageMeta.contentUrl,
        };

        return {
          content: [
            {
              type: "resource",
              resource: {
                mimeType: "text/html",
                text: resultPage.content ?? "",
                uri: uri, // Add uri to response
              },
            },
          ],
        };
      } catch (error) {
        throw new Error(
          `Failed to get content for page ${id}: ${getErrorMessage(error)}`, // Use getErrorMessage
        );
      }
    },
  );

  mcpServer.tool(
    "updatePageContent",
    "Update the content of an existing OneNote page",
    {
      id: z.string().describe("The ID of the page to update"),
      content: z
        .string()
        .describe(
          "The new page content in HTML format. This replaces the entire page body.",
        ),
    },
    async (args, extra): Promise<{ content: { type: "text"; text: string }[] }> => {
      const { id, content } = args;
      try {
        const patchPayload = [
          {
            target: "body",
            action: "replace",
            content: content,
          },
        ];

        await azureClient
          .api(`/me/onenote/pages/${id}/content`)
          .header("Content-Type", "application/json")
          .patch(patchPayload);

        return {
          content: [
            {
              type: "text",
              text: `Page ${id} content updated successfully.`,
            },
          ],
        };
      } catch (error) {
        throw new Error(
          `Failed to update page ${id}: ${getErrorMessage(error)}`,
        ); // Use getErrorMessage
      }
    },
  );

  mcpServer.tool(
    "deletePage",
    "Delete a specific OneNote page by its ID",
    {
      id: z.string().describe("The ID of the page to delete"),
    },
    async (args, extra): Promise<{ content: { type: "text"; text: string }[] }> => {
      const { id } = args;
      try {
        await azureClient.api(`/me/onenote/pages/${id}`).delete();
        return {
          content: [{ type: "text", text: `Page ${id} deleted successfully.` }],
        };
      } catch (error) {
        throw new Error(
          `Failed to delete page ${id}: ${getErrorMessage(error)}`,
        ); // Use getErrorMessage
      }
    },
  );
}
