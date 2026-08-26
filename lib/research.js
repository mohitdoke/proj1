// Server-side web research — News & Updates + Industry & Competitors —
// completely separate from the financial MIS data path. Uses Tavily,
// cached in company_research, and only ever reshapes results into the
// EXACT SAME data shapes the existing (unmodified) NewsUpdatesPage /
// IndustryCompetitorsPage components already render:
//   newsFeed:     Array<{ title, summary, category, publishedAt, sourceName,
//                          sourceUrl, secondarySourceName, secondarySourceUrl }>
//                 | []   (never invented — [] just means "nothing found")
//   industryData: { overviewDescription, categories[], snapshot[], trends[],
//                    competitors[], analysis[], methodology } | null
//   refreshMeta:  { newsRefreshedAt, industryRefreshedAt }
//
// Credit-conscious by design: at most ~2 Tavily searches per research
// cycle (one for news, one for industry/competitors), basic search depth,
// ~5 results each, no Tavily "research" mode, no crawling/extraction calls.
// A cycle only runs when the cache is missing/stale or force-refreshed —
// never just because a user switched tabs or re-opened the dashboard.
import { guessNewsCategory } from "../src/lib/misEngine.js";
import { getSupabaseAdmin } from "./supabaseAdmin.js";
import { getCompanyProfile } from "./dashboardRead.js";

const TAVILY_URL = "https://api.tavily.com/search";
const CACHE_TTL_MS = 24 * 60 * 60 * 1000; // refresh at most once/day per company
const MAX_RESULTS_PER_SEARCH = 5;

async function tavilySearch(query, { topic } = {}) {
  const apiKey = process.env.TAVILY_API_KEY;
  if (!apiKey) throw new Error("TAVILY_API_KEY is not set — research cannot run without it.");

  const res = await fetch(TAVILY_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      api_key: apiKey,
      query,
      search_depth: "basic", // never "advanced"/research mode — credit-conscious
      max_results: MAX_RESULTS_PER_SEARCH,
      include_answer: false,
      include_raw_content: false, // no crawling/extraction
      ...(topic ? { topic } : {}),
    }),
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Tavily search failed (${res.status}): ${body.slice(0, 300)}`);
  }
  const data = await res.json();
  return Array.isArray(data.results) ? data.results : [];
}

function toNewsItem(r) {
  if (!r || !r.url || !r.title) return null;
  const text = `${r.title} ${r.content || ""}`;
  return {
    title: r.title,
    summary: (r.content || "").slice(0, 280),
    category: guessNewsCategory(text),
    publishedAt: r.published_date ? new Date(r.published_date) : null,
    sourceName: hostnameOf(r.url),
    sourceUrl: r.url,
    secondarySourceName: null,
    secondarySourceUrl: null,
  };
}

function hostnameOf(url) {
  try {
    return new URL(url).hostname.replace(/^www\./, "");
  } catch {
    return "Source";
  }
}

function isoDateLabel(d) {
  return d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

/**
 * Runs the (at most 2-search) live research cycle for one company and
 * returns the shapes described above. Never throws for "nothing found" —
 * only for a genuine transport/config error, which the caller degrades
 * from gracefully (stale cache if any, else empty results).
 */
async function runResearchCycle({ companyName, description, tags }) {
  const newsQuery = `${companyName} news`;
  const newsResults = await tavilySearch(newsQuery);
  const newsFeed = newsResults.map(toNewsItem).filter(Boolean).slice(0, 5);

  const industryQuery = tags?.length
    ? `${companyName} competitors ${tags[0]} India`
    : `${companyName} competitors industry India`;
  const industryResults = await tavilySearch(industryQuery, { topic: "general" });

  // Keep this section deliberately concise, per the master prompt: industry
  // sub-sector line, 3-5 competitors, 2-4 trends — every externally-sourced
  // claim carries its own clickable source. We do not attempt to fill the
  // dashboard's optional capability-matrix columns (CRM/Loyalty/CDP/etc —
  // those were tailored to one specific company's sector); leaving them
  // unset renders as "N/A" in the existing table, which is accurate, not
  // fabricated.
  const competitorNames = new Set();
  const competitors = [];
  const trends = [];
  industryResults.slice(0, 5).forEach((r, i) => {
    if (!r || !r.url || !r.title) return;
    const source = { SourceName: hostnameOf(r.url), SourceUrl: r.url };
    if (i < 4 && competitors.length < 5) {
      // Best-effort "competitor" label from the result title — this is a
      // headline about the competitive landscape, not a verified roster;
      // RelevanceNote makes that provenance explicit in the UI.
      const name = r.title.length <= 60 ? r.title : `${companyName} competitor #${competitors.length + 1}`;
      if (!competitorNames.has(name)) {
        competitorNames.add(name);
        competitors.push({ name, PrimaryFocus: (r.content || "").slice(0, 120) || "N/A", ...source, RelevanceNote: "From live web research; not independently verified." });
      }
    } else {
      trends.push({ Title: r.title, Description: (r.content || "").slice(0, 220), WhyItMatters: null, ...source, PublishedAt: r.published_date ? isoDateLabel(new Date(r.published_date)) : null });
    }
  });

  const industryData = (competitors.length || trends.length)
    ? {
        overviewDescription: description ? `${description}\n\nCompetitive context below is drawn from live web search and refreshed periodically.` : null,
        categories: tags?.length ? tags.slice(0, 4).map(t => ({ Name: t })) : [],
        snapshot: [],
        trends: trends.slice(0, 4),
        competitors: competitors.slice(0, 5),
        analysis: [],
        methodology: "Industry & competitor items are sourced live via web search (Tavily) at most once a day per company; each item links to its original source.",
      }
    : null;

  return { newsFeed, industryData };
}

/**
 * Fetches this company's cached News/Industry research, refreshing it (at
 * most once per CACHE_TTL_MS, or when force is true) via a bounded Tavily
 * research cycle. Always returns the shapes NewsUpdatesPage /
 * IndustryCompetitorsPage expect — never null for newsFeed (an empty array
 * means "nothing found", not "go fetch client-side"), so the frontend never
 * needs the old browser-side GNews fallback.
 */
export async function getCompanyResearch(companyIdOrSlug, { force = false } = {}) {
  const supabase = getSupabaseAdmin();

  const profile = await getCompanyProfile(companyIdOrSlug);
  if (!profile.ok) return profile;
  const { company, companyInfo } = profile;

  const { data: cached, error: cacheErr } = await supabase
    .from("company_research")
    .select("news_data, industry_data, fetched_at, fetch_error")
    .eq("company_id", company.id)
    .maybeSingle();
  if (cacheErr) return { ok: false, error: `Database error loading cached research: ${cacheErr.message}` };

  const isFresh = cached?.fetched_at && (Date.now() - new Date(cached.fetched_at).getTime() < CACHE_TTL_MS);
  if (isFresh && !force) {
    return {
      ok: true,
      newsFeed: cached.news_data || [],
      industryData: cached.industry_data || null,
      refreshMeta: {
        newsRefreshedAt: isoDateLabel(new Date(cached.fetched_at)),
        industryRefreshedAt: isoDateLabel(new Date(cached.fetched_at)),
      },
    };
  }

  const companyName = companyInfo?.companyName || company.name;
  try {
    const { newsFeed, industryData } = await runResearchCycle({
      companyName,
      description: companyInfo?.description,
      tags: companyInfo?.tags,
    });
    const fetchedAt = new Date();
    // Store serialized dates as ISO strings in jsonb (Date objects don't
    // survive JSON.stringify the way we want otherwise).
    const newsForStorage = newsFeed.map(n => ({ ...n, publishedAt: n.publishedAt ? n.publishedAt.toISOString() : null }));
    await supabase
      .from("company_research")
      .upsert({ company_id: company.id, news_data: newsForStorage, industry_data: industryData, fetched_at: fetchedAt.toISOString(), fetch_error: null, updated_at: fetchedAt.toISOString() });

    return {
      ok: true,
      newsFeed: newsFeed, // in-memory Date objects — fine for this same request/response
      industryData,
      refreshMeta: { newsRefreshedAt: isoDateLabel(fetchedAt), industryRefreshedAt: isoDateLabel(fetchedAt) },
    };
  } catch (err) {
    // Live research failed (bad key, network, rate limit, ...) — degrade to
    // whatever was cached before rather than breaking the page; if there's
    // no prior cache at all, return empty (never fabricated) results.
    await supabase
      .from("company_research")
      .upsert({ company_id: company.id, fetch_error: String(err.message || err), updated_at: new Date().toISOString() });
    if (cached) {
      return {
        ok: true,
        newsFeed: (cached.news_data || []).map(n => ({ ...n, publishedAt: n.publishedAt ? new Date(n.publishedAt) : null })),
        industryData: cached.industry_data || null,
        refreshMeta: {
          newsRefreshedAt: cached.fetched_at ? isoDateLabel(new Date(cached.fetched_at)) : null,
          industryRefreshedAt: cached.fetched_at ? isoDateLabel(new Date(cached.fetched_at)) : null,
        },
        warning: `Live research refresh failed (${String(err.message || err)}); showing the last successful result.`,
      };
    }
    return { ok: true, newsFeed: [], industryData: null, refreshMeta: { newsRefreshedAt: null, industryRefreshedAt: null }, warning: String(err.message || err) };
  }
}
