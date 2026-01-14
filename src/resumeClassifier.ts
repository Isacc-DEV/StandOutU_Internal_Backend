import { randomUUID } from 'crypto';
import { buildAutofillPlanPrompt, buildJobAnalyzePrompt, buildRankResumesPrompt, buildResumeParsePrompt } from './promptPack';

type RankedResume = { id?: string; label: string; score: number };

const LABELS = [
  'Golang',
  'Java',
  'Rust',
  'Ruby',
  'DevOps',
  'AI',
  'Python',
  'C#',
  'Node.js',
  'PHP',
  'Kotlin',
  'Swift',
  'Frontend',
  'C++',
] as const;

const FALLBACK_PRIORITY = [
  'Golang',
  'Java',
  'Rust',
  'Ruby',
  'DevOps',
  'AI',
  'Python',
  'C#',
  'Node.js',
  'PHP',
  'Kotlin',
  'Swift',
  'Frontend',
  'C++',
] as const;

const TITLE_HIT_WEIGHT = 10;
const PRIMARY_HIT_WEIGHT = 8;
const GENERAL_HIT_WEIGHT = 0.9;
const REQUIRED_HIT_WEIGHT = 4;
const PREFERRED_HIT_WEIGHT = 0.7;

const CUSTOM_WEIGHTS: Record<string, number> = {
  Golang: 1.11,
  Java: 1.1,
  Rust: 1.0,
  Ruby: 1.0,
  DevOps: 1.08,
  AI: 0.9,
  Python: 0.93,
  'C#': 1.09,
  'Node.js': 0.85,
  PHP: 1.0,
  Kotlin: 0.9,
  Swift: 0.8,
  Frontend: 0.5,
  'C++': 1.0,
};

const STRONG_CONTEXT = [
  /main\s+stack/i,
  /primary\s+tech/i,
  /primary\s+stack/i,
  /primary\s+technology/i,
  /primary\s+language/i,
  /main\s+language/i,
  /core\s+tech/i,
  /core\s+stack/i,
  /core\s+language/i,
  /strong\s+(experience|background|skills?)/i,
  /focus(ed)?\s+on/i,
  /mainly\s+(work(ing)?\s+)?with/i,
];

const LANGUAGE_KEYWORDS: Record<string, RegExp[]> = {
  Golang: [/golang/i, /\bgo(lang)?\b/i],
  Java: [/java/i, /j2ee/i, /spring/i],
  DevOps: [
    /devops/i,
    /\bsre\b/i,
    /site\s+reliability/i,
    /kubernetes/i,
    /\bk8s\b/i,
    /docker/i,
    /terraform/i,
    /ansible/i,
    /cloudformation/i,
    /\bci\/?cd\b/i,
    /prometheus/i,
    /grafana/i,
  ],
  'C#': [/c#/i, /\.net/i, /dotnet/i, /asp\.?net/i],
  Ruby: [/ruby\b/i, /ruby\s+on\s+rails/i, /\brails\b/i],
  Rust: [/rust\b/i],
  PHP: [/php/i, /laravel/i],
  Kotlin: [/kotlin/i, /android/i],
  Swift: [/swift\b/i, /\bios\b/i, /xcode/i],
  'C++': [/c\+\+/i],
  Python: [/python/i, /django/i, /flask/i, /fastapi/i],
  'Node.js': [/node\.?js/i, /express\b/i, /\bnest\b/i],
  Frontend: [/frontend/i, /front\s*-?\s*end/i, /react\b/i, /vue\b/i, /angular\b/i, /typescript\b/i],
  AI: [/ai\b/i, /\bml\b/i, /machine\s+learning/i, /deep\s+learning/i, /\bnlp\b/i, /\bllm\b/i, /pytorch/i, /tensorflow/i],
};

const TITLE_KEYWORDS: Record<string, RegExp[]> = {
  Golang: [/golang\b/i, /\bgo\s+(developer|engineer|programmer)\b/i],
  Java: [/java\b/i, /\bjava\s+(developer|engineer)\b/i],
  DevOps: [/devops\b/i, /\bsre\b/i, /site reliability/i, /cloud\s+engineer/i, /platform\s+engineer/i],
  'C#': [/c#/i, /dotnet/i, /\.net/i],
  Ruby: [/ruby\b/i, /ruby\s+on\s+rails/i],
  Rust: [/rust\b/i],
  Python: [/python\b/i],
  'Node.js': [/node\.?js\b/i, /node\s+developer/i],
  AI: [/ai\b/i, /\bml\b/i, /machine\s+learning/i, /deep\s+learning/i, /data\s+scientist/i, /llm\b/i],
  PHP: [/php\b/i],
  Kotlin: [/kotlin\b/i],
  Swift: [/swift\b/i, /\bios\b/i, /\biphone\b/i],
  Frontend: [/frontend\b/i, /front\s*-?\s*end\b/i, /react\b/i, /vue\b/i, /angular\b/i],
  'C++': [/c\+\+/i],
};

const JAVASCRIPT_PATTERN = /\bjavascript\b/i;

const HF_TOKEN = process.env.HF_TOKEN || process.env.HUGGINGFACEHUB_API_TOKEN;
const HF_MODEL = 'meta-llama/Meta-Llama-3-8B-Instruct';

type ScoreMap = Record<string, number>;

function baseWeight(label: string) {
  return CUSTOM_WEIGHTS[label] ?? 1;
}

function requirementMultiplier(window: string) {
  let m = 1;
  if (/required|must have/i.test(window)) m *= REQUIRED_HIT_WEIGHT;
  if (/preferred|nice to have/i.test(window)) m *= PREFERRED_HIT_WEIGHT;
  return m;
}

function hasStrongContext(window: string) {
  return STRONG_CONTEXT.some((re) => re.test(window));
}

function scoreKeywords(text: string, pattern: RegExp, label: string, weightBase: number) {
  let score = 0;
  const re = pattern.global ? pattern : new RegExp(pattern.source, pattern.flags + 'g');
  for (const match of text.matchAll(re)) {
    const start = Math.max(0, match.index! - 40);
    const end = Math.min(text.length, match.index! + match[0].length + 40);
    const window = text.slice(start, end);
    const contextMultiplier = hasStrongContext(window) ? PRIMARY_HIT_WEIGHT : GENERAL_HIT_WEIGHT;
    const req = requirementMultiplier(window);
    score += weightBase * contextMultiplier * req;
  }
  return score;
}

function normalizeText(html: string) {
  return html.replace(/<script[\s\S]*?<\/script>/gi, '').replace(/<style[\s\S]*?<\/style>/gi, '').replace(/<[^>]+>/g, ' ');
}

async function fetchJobText(url: string): Promise<{ text: string; title?: string }> {
  try {
    const res = await fetch(url, { redirect: 'follow', signal: AbortSignal.timeout(15000) });
    const raw = await res.text();
    const plain = normalizeText(raw);
    const titleMatch = raw.match(/<title>([^<]{3,120})<\/title>/i);
    return { text: plain, title: titleMatch?.[1]?.trim() };
  } catch (err) {
    return { text: '', title: undefined };
  }
}

function classify(title: string | undefined, text: string) {
  const lower = text.toLowerCase();
  const roleScores: ScoreMap = {};
  const titleScores: ScoreMap = {};

  // Title-driven boosts
  if (title) {
    const tLower = title.toLowerCase();
    for (const [label, patterns] of Object.entries(TITLE_KEYWORDS)) {
      for (const pattern of patterns) {
        const hits = (tLower.match(pattern) || []).length;
        if (hits) {
          titleScores[label] = Math.max(titleScores[label] ?? 0, hits * baseWeight(label) * TITLE_HIT_WEIGHT);
        }
      }
    }
    if (JAVASCRIPT_PATTERN.test(tLower)) {
      titleScores['Frontend'] = Math.max(titleScores['Frontend'] ?? 0, baseWeight('Frontend') * TITLE_HIT_WEIGHT * 0.5);
      titleScores['Node.js'] = Math.max(titleScores['Node.js'] ?? 0, baseWeight('Node.js') * TITLE_HIT_WEIGHT * 0.5);
    }
  }

  // Body keyword scores
  for (const [label, patterns] of Object.entries(LANGUAGE_KEYWORDS)) {
    const weightBase = baseWeight(label);
    let score = 0;
    for (const pattern of patterns) {
      score += scoreKeywords(lower, pattern, label, weightBase);
    }
    roleScores[label] = score;
  }
  if (JAVASCRIPT_PATTERN.test(lower)) {
    const weightBase = baseWeight('Node.js');
    const base = scoreKeywords(lower, JAVASCRIPT_PATTERN, 'Node.js', weightBase);
    roleScores['Frontend'] = (roleScores['Frontend'] ?? 0) + base * 0.5;
    roleScores['Node.js'] = (roleScores['Node.js'] ?? 0) + base * 0.5;
  }

  // Merge
  const finalScores: ScoreMap = {};
  for (const label of LABELS) {
    finalScores[label] = Math.max(roleScores[label] ?? 0, titleScores[label] ?? 0);
  }

  // Pick winners
  const sorted = Object.entries(finalScores)
    .map(([label, score]) => ({ label, score }))
    .sort((a, b) => b.score - a.score || FALLBACK_PRIORITY.indexOf(a.label as any) - FALLBACK_PRIORITY.indexOf(b.label as any));

  const best = sorted[0]?.label ?? FALLBACK_PRIORITY[0];
  return { best, scores: finalScores, ranked: sorted };
}

async function callHuggingFace(prompt: string): Promise<string | undefined> {
  if (!HF_TOKEN) return undefined;
  try {
    const res = await fetch('https://router.huggingface.co/v1/chat/completions', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${HF_TOKEN}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        model: HF_MODEL,
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 256,
        temperature: 0.1,
      }),
    });
    const data = (await res.json()) as any;
    const content =
      data?.choices?.[0]?.message?.content ||
      (Array.isArray(data) && data[0]?.generated_text) ||
      data?.generated_text ||
      data?.text;
    if (typeof content === 'string') return content.trim();
  } catch {
    // ignore HF errors
  }
  return undefined;
}

export async function callPromptPack(prompt: string): Promise<any | undefined> {
  const raw = await callHuggingFace(prompt);
  if (!raw) return undefined;
  try {
    return JSON.parse(raw);
  } catch {
    return undefined;
  }
}

export const promptBuilders = {
  buildResumeParsePrompt,
  buildJobAnalyzePrompt,
  buildRankResumesPrompt,
  buildAutofillPlanPrompt,
};

async function classifyFromText(
  title: string | undefined,
  text: string,
  resumesInput?: { id: string; label: string; parsed?: Record<string, unknown>; resume_text?: string }[],
): Promise<{
  id: string;
  recommendedLabel?: string;
  recommendedResumeId?: string;
  ranked: RankedResume[];
  rawScores: Record<string, number>;
  title: string;
  jobText: string;
}> {
  const classified = classify(title, text);
  const ranked: RankedResume[] = classified.ranked.map((r) => ({
    id: r.label,
    label: r.label,
    score: Number.isFinite(r.score) ? Number(r.score) : 0,
  }));

  return {
    id: randomUUID(),
    recommendedLabel: classified.best,
    recommendedResumeId: undefined,
    ranked,
    rawScores: classified.scores,
    title: title ?? '',
    jobText: text,
  };
}

export async function analyzeJobFromUrl(
  url: string,
  resumesInput?: { id: string; label: string; parsed?: Record<string, unknown>; resume_text?: string }[],
) {
  const { text, title } = await fetchJobText(url);
  return classifyFromText(title, text, resumesInput);
}

export async function analyzeJobFromHtml(
  html: string,
  pageTitle?: string,
  resumesInput?: { id: string; label: string; parsed?: Record<string, unknown>; resume_text?: string }[],
) {
  const text = normalizeText(html || '');
  return classifyFromText(pageTitle, text, resumesInput);
}
