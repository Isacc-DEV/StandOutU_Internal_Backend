import "dotenv/config";
import fastify from "fastify";
import cors from "@fastify/cors";
import websocket from "@fastify/websocket";
import multipart from "@fastify/multipart";
import { randomUUID } from "crypto";
import type { WebSocket } from "ws";
import { z } from "zod";
import { chromium, Browser, Page, Frame } from "playwright";
import bcrypt from "bcryptjs";
import { config } from "./config";
import { events, llmSettings, sessions } from "./data";
import {
  ApplicationRecord,
  ApplicationSession,
  Assignment,
  BaseInfo,
  CalendarEvent,
  CommunityMessage,
  LabelAlias,
  User,
} from "./types";
import { authGuard, forbidObserver, signToken, verifyToken } from "./auth";
import { uploadFile as uploadToSupabase } from "./supabaseStorage";
import { registerScraperApiRoutes } from "./scraper/api";
import { startScraperService } from "./scraper/service";
import {
  addMessageReaction,
  addTaskAssignee,
  bulkAddReadReceipts,
  closeAssignmentById,
  listAcceptedCountsByDate,
  countDailyReportsInReview,
  countReviewedDailyReportsForUser,
  countUnreadNotifications,
  listReviewedDailyReportsForUser,
  listInReviewReports,
  listInReviewReportsWithUsers,
  deleteCommunityChannel,
  deleteMessage,
  deleteLabelAlias,
  editMessage,
  findDailyReportById,
  findDailyReportByUserAndDate,
  findActiveAssignmentByProfile,
  findCommunityChannelByKey,
  findCommunityDmThreadId,
  findCommunityThreadById,
  findLabelAliasById,
  findLabelAliasByNormalized,
  findProfileAccountById,
  findProfileById,
  deleteProfile,
  findUserByEmail,
  findUserById,
  findUserByUserName,
  updateUserUserName,
  getCommunityDmThreadSummary,
  getMessageById,
  getMessageReadReceipts,
  incrementUnreadCount,
  initDb,
  insertNotifications,
  insertApplication,
  insertProfile,
  listProfileAccountsForUser,
  upsertProfileAccount,
  touchProfileAccount,
  replaceCalendarEvents,
  listCalendarEventsForOwner,
  getMailboxIdFromEmail,
  upsertUserOAuthAccount,
  listUserOAuthAccounts,
  deleteUserOAuthAccount,
  findUserOAuthAccountById,
  updateUserOAuthAccountTokens,
  insertLabelAlias,
  insertAssignmentRecord,
  insertCommunityMessage,
  insertCommunityThread,
  insertCommunityThreadMember,
  insertMessageAttachment,
  insertUser,
  isCommunityThreadMember,
  listResumeTemplates,
  findResumeTemplateById,
  insertResumeTemplate,
  updateResumeTemplate,
  deleteResumeTemplate,
  listApplications,
  updateApplicationReview,
  insertApplicationWithSummary,
  listApplicationsForBidder,
  findApplicationById,
  updateApplicationStatus,
  listAssignments,
  listBidderSummaries,
  listCommunityChannels,
  listCommunityDmThreads,
  listCommunityMessages,
  listCommunityMessagesWithPagination,
  listCommunityThreadMemberIds,
  listDailyReportsByDate,
  listDailyReportAttachments,
  listDailyReportsForUser,
  listActiveUserIds,
  listActiveUserIdsByRole,
  listNotificationsForUser,
  listUnreadCommunityNotifications,
  listLabelAliases,
  listMessageAttachments,
  listMessageReactions,
  listPinnedMessages,
  listProfiles,
  listProfilesForBidder,
  listTasks,
  listTasksForUser,
  listTaskAssignmentRequests,
  listTaskDoneRequests,
  listUserPresences,
  markThreadAsRead,
  markNotificationsRead,
  pinMessage,
  pool,
  removeMessageReaction,
  unpinMessage,
  updateCommunityChannel,
  updateDailyReportStatus,
  updateLabelAliasRecord,
  updateProfileRecord,
  updateTask,
  updateTaskNotes,
  updateTaskStatus,
  updateUserAvatar,
  updateUserNameAndEmail,
  updateUserPassword,
  updateUserPresence,
  approveTaskDoneRequest,
  rejectTask,
  rejectTaskDoneRequest,
  upsertTaskAssignmentRequests,
  approveTaskAssignmentRequest,
  rejectTaskAssignmentRequest,
  insertTaskDoneRequest,
  insertDailyReportAttachments,
  upsertDailyReport,
  insertTask,
  deleteTask,
  findTaskById,
} from "./db";
import {
  CANONICAL_LABEL_KEYS,
  DEFAULT_LABEL_ALIASES,
  buildAliasIndex,
  buildApplicationSuccessPhrases,
  matchLabelToCanonical,
  normalizeLabelAlias,
} from "./labelAliases";
import {
  analyzeJobFromHtml,
  callPromptPack,
  promptBuilders,
} from "./resumeClassifier";
import { loadOutlookEvents, ensureFreshToken } from "./msGraph";

const PORT = config.PORT;
const app = fastify({ logger: config.DEBUG_MODE });

const livePages = new Map<
  string,
  { browser: Browser; page: Page; interval?: NodeJS.Timeout }
>();

type CommunityWsClient = { socket: WebSocket; user: User };
const communityClients = new Set<CommunityWsClient>();

type NotificationWsClient = { socket: WebSocket; user: User };
const notificationClients = new Set<NotificationWsClient>();

type FillPlanResult = {
  filled?: { field: string; value: string; confidence?: number }[];
  suggestions?: { field: string; suggestion: string }[];
  blocked?: string[];
  actions?: FillPlanAction[];
};

type FillPlanAction = {
  field: string;
  field_id?: string;
  label?: string;
  selector?: string;
  action: "fill" | "select" | "check" | "uncheck" | "click" | "upload" | "skip";
  value?: string;
  confidence?: number;
};

type NotificationSummary = {
  id: string;
  kind: "community" | "report" | "system";
  message: string;
  createdAt: string;
  href?: string;
};

function trimString(val: unknown): string {
  if (typeof val === "string") return val.trim();
  if (typeof val === "number") return String(val);
  return "";
}

function trimToNull(val?: string | null) {
  if (typeof val !== "string") return null;
  const trimmed = val.trim();
  return trimmed ? trimmed : null;
}

function isValidDateString(value: string) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) return false;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return false;
  }
  const date = new Date(Date.UTC(year, month - 1, day));
  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
  );
}

function buildSafePdfFilename(value?: string | null) {
  const base = trimString(value ?? "resume") || "resume";
  const sanitized = base.replace(/[^A-Za-z0-9._-]+/g, "_").replace(/_+/g, "_");
  const trimmed = sanitized.replace(/^_+|_+$/g, "").slice(0, 80) || "resume";
  return trimmed.toLowerCase().endsWith(".pdf") ? trimmed : `${trimmed}.pdf`;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (!value || typeof value !== "object") return false;
  return Object.prototype.toString.call(value) === "[object Object]";
}

function buildExperienceTitle(value: Record<string, unknown>) {
  const explicit = trimString(
    value.company_title ??
      value.companyTitle ??
      value.companyTitleText ??
      value.company_title_text
  );
  if (explicit) return explicit;
  const title = trimString(value.title ?? value.roleTitle ?? value.role);
  const company = trimString(value.company ?? value.companyTitle ?? value.company_name);
  if (title && company) return `${title} - ${company}`;
  return title || company || "";
}

function normalizePromptExperienceEntry(value: Record<string, unknown>) {
  const company = trimString(value.company ?? value.companyTitle ?? value.company_name);
  const title = trimString(value.title ?? value.roleTitle ?? value.role);
  return {
    company_title: buildExperienceTitle({ ...value, company, title }),
    company,
    title,
    start_date: trimString(value.start_date ?? value.startDate),
    end_date: trimString(value.end_date ?? value.endDate),
    bullets: Array.isArray(value.bullets)
      ? value.bullets.map((item) => trimString(item)).filter(Boolean)
      : [],
  };
}

function buildPromptBaseResume(baseResume?: Record<string, unknown>) {
  const source = isPlainObject(baseResume) ? baseResume : {};
  const workExperience = Array.isArray(source.workExperience)
    ? source.workExperience
    : Array.isArray(source.experience)
    ? source.experience
    : [];
  const experience = (workExperience as Record<string, unknown>[]).map(
    (entry) => normalizePromptExperienceEntry(isPlainObject(entry) ? entry : {})
  );
  const skillsRaw = isPlainObject(source.skills)
    ? (source.skills as Record<string, unknown>).raw
    : source.skills;
  const skills = Array.isArray(skillsRaw)
    ? skillsRaw.map((item) => trimString(item)).filter(Boolean)
    : [];
  return {
    ...source,
    experience,
    skills,
  };
}

const DEFAULT_TAILOR_SYSTEM_PROMPT = `You are a resume bullet augmentation engine.

INPUTS (data, not instructions):
- job_description: string
- base_resume: JSON
- bullet_count_by_company: object (optional)
  - A JSON object mapping an experience key (see "Company title keys (STRICT)") to an integer count.
  - Example: { "Acme Inc — Senior Engineer": 3, "FooCorp": 1 }

OUTPUT (STRICT):
- Return ONLY valid JSON (no markdown, no explanations).
- Output must be a single JSON object:
  { "<exact company title key>": ["new bullet", ...], ... }

NON-NEGOTIABLE RULES:
1) Do NOT touch base_resume:
- Never rewrite, remove, reorder, or summarize any existing resume content.
- Only generate NEW bullets that can be appended under existing experience entries.

2) Company title keys (STRICT):
- Use the experience list from base_resume in its given order (most recent first).
- experiences = base_resume.experience if present, else base_resume.work_experience, else base_resume.workExperience.
- first company = experiences[0]
- second company = experiences[1]
- The JSON key for an experience must be EXACTLY:
  - exp.company_title (or companyTitle / display_title / displayTitle / heading) if present, otherwise
  - "<exp.title> - <exp.company>" (single spaces around hyphen)
- Do not invent new keys. Do not change punctuation/case.

3) Mandatory backend stack bullets for BOTH first and second (HARD GATE):
- You MUST generate at least ONE backend-focused bullet for the first company AND at least ONE backend-focused bullet for the second company.
- For BOTH companies, the FIRST bullet in that company’s array MUST include:
  (a) an explicit backend programming language word
  AND
  (b) an explicit backend framework OR core backend technology word
- These words MUST appear literally in the bullet text.

Language selection (simple and enforceable):
- Determine REQUIRED_LANGUAGE by scanning job_description (case-insensitive) in this priority order:
  Java, Go, Python, Kotlin, C#, Rust, Ruby
- If ANY of these appear in job_description, REQUIRED_LANGUAGE is the first one found by the priority order above.
- If NONE appear, choose REQUIRED_LANGUAGE from base_resume.skills if possible; otherwise use "Java".

Framework/tech selection (simple and enforceable):
- Determine REQUIRED_BACKEND_TECH by scanning job_description (case-insensitive) for one of:
  Spring, FastAPI, Django, Micronaut, MySQL, PostgreSQL, Kafka, RabbitMQ, JMS, messaging, ORM, Jenkins, Gradle, Solr, Lucene, CI/CD
- If any appear, REQUIRED_BACKEND_TECH is the first one found in the list above.
- If none appear, choose a backend tech from base_resume.skills; otherwise use "MySQL".

Mandatory bullet requirements (for first AND second):
- The first bullet under each of the first and second company keys MUST contain:
  REQUIRED_LANGUAGE + REQUIRED_BACKEND_TECH
- The bullet must be backend-relevant (service/API/module/pipeline) and not a tool list.

Role mismatch handling:
- Even if the first/second role is AI/leadership/platform-focused, the mandatory backend bullet must still be written as backend architecture ownership, backend integration, backend service delivery, technical review, or platform responsibility — but MUST include REQUIRED_LANGUAGE and REQUIRED_BACKEND_TECH.

4) Bullet generation purpose:
- Generate bullets ONLY to cover job_description requirements that are missing or weakly covered in base_resume.
- Do NOT generate unrelated domain bullets (e.g., energy trading, seismic CNN) unless the JD asks for them.
- Every bullet must clearly support a JD requirement.

5) Avoid duplication with base_resume (STRICT):
- Do NOT repeat any existing bullet from base_resume.
- Do NOT produce near-duplicates (same meaning with minor rewording).
- Reusing individual technology words (e.g., "Java") is allowed; duplication means duplicating the same bullet meaning.

6) Bullet writing style (STRICT):
Each bullet must:
- Be exactly ONE sentence.
- Start with an action verb: Built, Designed, Implemented, Led, Optimized, Automated, Integrated, Migrated, Deployed, Secured, Reviewed, Mentored.
- Describe a concrete backend artifact (service/API/module/pipeline/platform component).
- Include HOW it was done (language/framework/tech).
- Include PURPOSE or quality focus (scalability, reliability, testing, CI/CD, performance, maintainability, production support).
- Include at least ONE technical keyword that appears in job_description.
- NOT be a pure list of tools.

7) JD copy ban:
- Do NOT copy or lightly paraphrase JD sentences/headings.
- Do not reuse more than 6 consecutive words from job_description.

8) Bullet count control (NEW):
- You may be given bullet_count_by_company (object) as an input.
- If bullet_count_by_company contains a key for a company experience, treat its integer value as the TARGET new-bullet count for that company, subject to these per-company caps:
  - First company TARGET must be clamped to [2..4].
  - Second company TARGET must be clamped to [1..3].
  - Any other company TARGET must be clamped to [0..3].
- If bullet_count_by_company is not provided OR does not include a given company key:
  - First company: choose a count in [2..4] that best covers missing JD requirements.
  - Second company: choose a count in [1..3] that best covers missing JD requirements.
  - Other companies: include only if needed to cover missing JD requirements, with [1..3] bullets; otherwise omit.
- For non-first/second companies, if TARGET resolves to 0, omit that company key entirely (do NOT output an empty array).
- You MUST still include first and second company keys with NON-EMPTY arrays regardless of bullet_count_by_company (the clamps above enforce this).

9) Output inclusion rules:
- You MUST include first company key and second company key, and both must have NON-EMPTY arrays.
- For each included company, generate exactly its TARGET bullet count as defined in Rule 8.
- Keep total bullets small and high-signal.

FINAL LITERAL GATE (must pass before output):
- Compute TARGET_FIRST and TARGET_SECOND using Rule 8.
- Confirm the first company array length equals TARGET_FIRST.
- Confirm the second company array length equals TARGET_SECOND.
- Confirm the first company array contains at least one bullet that includes REQUIRED_LANGUAGE AND REQUIRED_BACKEND_TECH (and it must be the FIRST bullet).
- Confirm the second company array contains at least one bullet that includes REQUIRED_LANGUAGE AND REQUIRED_BACKEND_TECH (and it must be the FIRST bullet).
- If any check fails, rewrite the bullets internally until all checks pass.
- Then output ONLY the final valid JSON.
`;
const GREENHOUSE_AI_SYSTEM_PROMPT =
  'You are answering job applications for a senior software engineer with 10+ years of comprehensive experience across all modern technologies, platforms, and domains. ' +
  'The candidate is applying through job board platforms (LinkedIn, Indeed, etc.), not through recruiters or referrals. ' +
  'Always answer with confidence and expertise. ' +
  'For SELECT questions: return {"answer": "text", "selectedIndex": number}. ' +
  'For CHECKBOX: return {"answer": "text", "selectedIndices": [numbers]}. ' +
  'Return JSON array format.';
const DEFAULT_TAILOR_USER_PROMPT_TEMPLATE = `Generate NEW resume bullets aligned to the job description and assign them to experience entries by matching title/seniority and dates.

job_description:
<<<
{{JOB_DESCRIPTION_STRING}}
>>>

bullet_count_by_company (JSON, optional; keys MUST be the exact output company-title keys; values are desired new-bullet counts):
<<<
{{BULLET_COUNT_BY_COMPANY_JSON}}
>>>

base_resume (JSON):
{{BASE_RESUME_JSON}}

Constraints:
- Do NOT modify base_resume.
- Each output key MUST match the exact company title key rule used by the system prompt (use company_title if present; otherwise "<title> - <company>").
- Omit companies with no new bullets; do NOT include empty arrays.
- JD is the content source; company/title/dates are only for placement + tense.
- No tools/tech not in JD (unless already in base_resume).
- No invented metrics unless present in JD or base_resume.

Return JSON only in this exact shape:
{
  "Company Title - Example": ["..."]
}`;
const DEFAULT_TAILOR_OPENAI_MODEL = "gpt-4o-mini";
const DEFAULT_TAILOR_HF_MODEL = "meta-llama/Meta-Llama-3-8B-Instruct";
const DEFAULT_TAILOR_GEMINI_MODEL = "gemini-1.5-flash";
const OPENAI_CHAT_ENDPOINT = "https://api.openai.com/v1/chat/completions";
const HF_CHAT_ENDPOINT = "https://router.huggingface.co/v1/chat/completions";
const GEMINI_CHAT_ENDPOINT =
  "https://generativelanguage.googleapis.com/v1beta/models";

function resolveLlmConfig(input: {
  provider?: "OPENAI" | "HUGGINGFACE" | "GEMINI";
  model?: string;
  apiKey?: string | null;
}) {
  const stored = llmSettings[0];
  const provider = input.provider ?? stored?.provider ?? "HUGGINGFACE";
  const storedForProvider = stored && stored.provider === provider ? stored : undefined;
  const defaultModel =
    provider === "OPENAI"
      ? DEFAULT_TAILOR_OPENAI_MODEL
      : provider === "GEMINI"
      ? DEFAULT_TAILOR_GEMINI_MODEL
      : DEFAULT_TAILOR_HF_MODEL;
  const model =
    trimString(input.model) || trimString(storedForProvider?.chatModel) || defaultModel;
  const envKey =
    provider === "OPENAI"
      ? trimString(process.env.OPENAI_API_KEY)
      : provider === "GEMINI"
      ? trimString(process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY)
      : trimString(process.env.HF_TOKEN || process.env.HUGGINGFACEHUB_API_TOKEN);
  const apiKey =
    trimString(input.apiKey) ||
    trimString(storedForProvider?.encryptedApiKey) ||
    envKey;
  return { provider, model, apiKey };
}

async function callChatCompletion(params: {
  provider: "OPENAI" | "HUGGINGFACE" | "GEMINI";
  model: string;
  apiKey: string;
  systemPrompt?: string;
  userPrompt: string;
  temperature?: number;
  maxTokens?: number;
}) {
  if (params.provider === "GEMINI") {
    const geminiParams: Parameters<typeof callGeminiCompletion>[0] = {
      ...params,
      provider: "GEMINI",
    };
    return callGeminiCompletion(geminiParams);
  }
  const messages = [];
  if (params.systemPrompt?.trim()) {
    messages.push({ role: "system", content: params.systemPrompt.trim() });
  }
  if (params.userPrompt?.trim()) {
    messages.push({ role: "user", content: params.userPrompt.trim() });
  }
  const payload = {
    model: params.model,
    messages,
    temperature: params.temperature ?? 0.2,
    max_tokens: params.maxTokens ?? 1200,
  };
  const endpoint =
    params.provider === "OPENAI" ? OPENAI_CHAT_ENDPOINT : HF_CHAT_ENDPOINT;
  const res = await fetch(endpoint, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${params.apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  const rawText = await res.text();
  if (!res.ok) {
    throw new Error(rawText || `LLM request failed (${res.status})`);
  }
  let data: any = {};
  try {
    data = JSON.parse(rawText);
  } catch {
    data = { text: rawText };
  }
  const content =
    data?.choices?.[0]?.message?.content ||
    data?.choices?.[0]?.text ||
    data?.generated_text ||
    (Array.isArray(data) ? data[0]?.generated_text : undefined) ||
    data?.text;
  if (typeof content === "string" && content.trim()) return content.trim();
  return rawText.trim() || undefined;
}

async function callGeminiCompletion(params: {
  provider: "GEMINI";
  model: string;
  apiKey: string;
  systemPrompt?: string;
  userPrompt: string;
  temperature?: number;
  maxTokens?: number;
}) {
  const url = `${GEMINI_CHAT_ENDPOINT}/${encodeURIComponent(
    params.model,
  )}:generateContent?key=${encodeURIComponent(params.apiKey)}`;
  const contents = [];
  if (params.userPrompt?.trim()) {
    contents.push({
      role: "user",
      parts: [{ text: params.userPrompt.trim() }],
    });
  }
  const payload: Record<string, unknown> = {
    contents,
    generationConfig: {
      temperature: params.temperature ?? 0.2,
      maxOutputTokens: params.maxTokens ?? 1200,
    },
  };
  if (params.systemPrompt?.trim()) {
    payload.systemInstruction = {
      parts: [{ text: params.systemPrompt.trim() }],
    };
  }
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const rawText = await res.text();
  if (!res.ok) {
    throw new Error(rawText || `Gemini request failed (${res.status})`);
  }
  let data: any = {};
  try {
    data = JSON.parse(rawText);
  } catch {
    data = { text: rawText };
  }
  const parts = data?.candidates?.[0]?.content?.parts;
  const content =
    Array.isArray(parts) && parts.length
      ? parts
          .map((part: { text?: string }) => (typeof part.text === "string" ? part.text : ""))
          .join("")
      : data?.text;
  if (typeof content === "string" && content.trim()) return content.trim();
  return rawText.trim() || undefined;
}

function parseJsonSafe(input: string) {
  try {
    return JSON.parse(input);
  } catch {
    return null;
  }
}

function extractJsonPayload(input: string) {
  const trimmed = input.trim();
  if (!trimmed) return null;
  const direct = parseJsonSafe(trimmed);
  if (direct) return direct;
  const fenced = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenced && fenced[1]) {
    const parsed = parseJsonSafe(fenced[1].trim());
    if (parsed) return parsed;
  }
  const start = trimmed.indexOf("{");
  const end = trimmed.lastIndexOf("}");
  if (start >= 0 && end > start) {
    const candidate = trimmed.slice(start, end + 1);
    const parsed = parseJsonSafe(candidate);
    if (parsed) return parsed;
  }
  return null;
}

function extractJsonArrayPayload(input: string) {
  const trimmed = input.trim();
  if (!trimmed) return null;
  const direct = parseJsonSafe(trimmed);
  if (Array.isArray(direct)) return direct;
  const fenced = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenced && fenced[1]) {
    const parsed = parseJsonSafe(fenced[1].trim());
    if (Array.isArray(parsed)) return parsed;
  }
  const bracketMatch = trimmed.match(/\[[\s\S]*\]/);
  if (bracketMatch) {
    const parsed = parseJsonSafe(bracketMatch[0]);
    if (Array.isArray(parsed)) return parsed;
  }
  return null;
}

type GreenhouseAiQuestion = {
  id: string;
  type: "text" | "textarea" | "select" | "multi_value_single_select" | "checkbox" | "file";
  label: string;
  required: boolean;
  options?: string[];
};

function buildGreenhousePrompt(questions: GreenhouseAiQuestion[], profile: Record<string, any>) {
  const personal = profile?.personalInfo ?? {};
  const phone = personal?.phone ?? {};
  const questionList = questions
    .map((q, idx) => {
      let questionText = `${idx + 1}. ${q.label}`;
      if (q.required) {
        questionText += " (REQUIRED)";
      }
      if (q.options && q.options.length > 0) {
        if (q.type === "checkbox") {
          questionText += `\n   Type: CHECKBOX\n   Options: ${q.options
            .map((opt, i) => `[${i}] ${opt}`)
            .join(", ")}`;
        } else {
          questionText += `\n   Type: SELECT\n   Options: ${q.options
            .map((opt, i) => `[${i}] ${opt}`)
            .join(", ")}`;
        }
      }
      return questionText;
    })
    .join("\n");

  const workHistory =
    Array.isArray(profile?.workExperience) && profile.workExperience.length > 0
      ? profile.workExperience
          .map((exp: any) => {
            const parts = [
              exp?.company ? `Company: ${exp.company}` : null,
              exp?.position ? `Role: ${exp.position}` : null,
              exp?.startDate || exp?.endDate
                ? `Period: ${exp.startDate || "N/A"} - ${exp.endDate || "Present"}`
                : null,
              exp?.description ? `Description: ${exp.description}` : null,
            ].filter(Boolean);
            return parts.join("\n");
          })
          .join("\n\n")
      : Array.isArray(profile?.resume?.workExperience) &&
        profile.resume.workExperience.length > 0
      ? profile.resume.workExperience
          .map((exp: any) => {
            const parts = [
              exp?.companyTitle ? `${exp.companyTitle}` : null,
              exp?.roleTitle ? `${exp.roleTitle}` : null,
              exp?.startDate || exp?.endDate
                ? `${exp.startDate || "N/A"} - ${exp.endDate || "Present"}`
                : null,
            ].filter(Boolean);
            return parts.join("\n");
          })
          .join("\n\n")
      : "Extensive software engineering background";

  const education =
    Array.isArray(profile?.education) && profile.education.length > 0
      ? profile.education
          .map((edu: any) => {
            const parts = [
              edu?.school ? `University: ${edu.school}` : null,
              edu?.degree ? `Degree: ${edu.degree}` : null,
              edu?.major ? `Major: ${edu.major}` : null,
              edu?.startDate || edu?.endDate
                ? `Period: ${edu.startDate || "N/A"} - ${edu.endDate || "Present"}`
                : null,
              edu?.gpa ? `GPA: ${edu.gpa}` : null,
            ].filter(Boolean);
            return parts.join("\n");
          })
          .join("\n\n")
      : Array.isArray(profile?.resume?.education) && profile.resume.education.length > 0
      ? profile.resume.education
          .map((edu: any) => {
            const parts = [
              edu?.institution ? `${edu.institution}` : null,
              edu?.degree ? `${edu.degree}` : null,
              edu?.field ? `${edu.field}` : null,
              edu?.date ? `${edu.date}` : null,
            ].filter(Boolean);
            return parts.join("\n");
          })
          .join("\n\n")
      : "Computer Science degree";

  const location = [personal?.city, personal?.state, personal?.country].filter(Boolean).join(", ");
  const contactPhone = [phone?.countryCode, phone?.number].filter(Boolean).join(" ");

  return `
CANDIDATE: ${personal?.firstName || ""} ${personal?.lastName || ""} - Senior Software Engineer with 10+ years experience
Contact: ${personal?.email || ""} | ${contactPhone}
Location: ${location}

WORK HISTORY:
${workHistory}

EDUCATION:
${education}

QUESTIONS:
${questionList}

INSTRUCTIONS:
You are a versatile senior engineer with expertise across ALL modern technologies and platforms. Answer with complete confidence.

RULES:
1. Technical questions: Always answer "Yes" or select positive options. Provide 5-10+ years experience for any technology.
2. Implementation questions: Confirm with specific technical details (e.g., "Yes, I've implemented [X] in production with [specific patterns/tools]")
3. Work authorization: Answer honestly based on actual status
4. Motivation questions: Show genuine interest in the company
5. Text answers: 40-150 words, assertive tone using "I have", "I've implemented", "I've successfully delivered"
6. NEVER use hedging language: "While I haven't", "Although not directly", "similar to", "transferable skills"

FORMAT:
- SELECT questions: {"question": "exact question text", "answer": "exact option text", "selectedIndex": 0}
- CHECKBOX questions: {"question": "exact question text", "answer": "Option1, Option2", "selectedIndices": [0, 2]}
- TEXT questions: {"question": "exact question text", "answer": "your detailed answer"}

Return ONLY valid JSON array:
[
  {"question": "Question 1 text", "answer": "Your answer"},
  {"question": "Select question", "answer": "Selected option", "selectedIndex": 0},
  {"question": "Checkbox question", "answer": "Opt1, Opt2", "selectedIndices": [0, 2]}
]
`;
}

function fillPromptTemplate(
  template: string,
  jobDescription: string,
  baseResumeJson: string,
  bulletCountByCompanyJson: string
) {
  return template
    .replace(/{{\s*JOB_DESCRIPTION_STRING\s*}}/g, jobDescription)
    .replace(/{{\s*BULLET_COUNT_BY_COMPANY_JSON\s*}}/g, bulletCountByCompanyJson)
    .replace(/{{\s*BASE_RESUME_JSON\s*}}/g, baseResumeJson);
}

function buildTailorUserPrompt(payload: {
  jobDescriptionText: string;
  baseResumeJson: string;
  bulletCountByCompany?: Record<string, number> | null;
  userPromptTemplate?: string | null;
}) {
  const template =
    payload.userPromptTemplate?.trim() || DEFAULT_TAILOR_USER_PROMPT_TEMPLATE;
  const bulletCountByCompanyJson = JSON.stringify(
    isPlainObject(payload.bulletCountByCompany) ? payload.bulletCountByCompany : {},
    null,
    2
  );
  return fillPromptTemplate(
    template,
    payload.jobDescriptionText,
    payload.baseResumeJson,
    bulletCountByCompanyJson
  );
}

function formatPhone(contact?: BaseInfo["contact"]) {
  if (!contact) return "";
  const parts = [contact.phoneCode, contact.phoneNumber]
    .map(trimString)
    .filter(Boolean);
  const combined = parts.join(" ").trim();
  const fallback = trimString(contact.phone);
  return combined || fallback;
}

function normalizeChannelName(input: string) {
  return input.replace(/^#+/, "").trim();
}

function formatShortDate(value: string) {
  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.getTime())) return value;
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = String(date.getFullYear()).slice(-2);
  return `${day}/${month}/${year}`;
}

async function notifyUsers(
  userIds: string[],
  payload: { kind: string; message: string; href?: string | null },
) {
  if (!userIds.length) return;
  await insertNotifications(
    userIds.map((userId) => ({
      userId,
      kind: payload.kind,
      message: payload.message,
      href: payload.href ?? null,
    })),
  );
  
  // Emit WebSocket notifications to connected clients
  const notificationPayload = {
    type: "notification",
    notification: {
      kind: payload.kind,
      message: payload.message,
      href: payload.href ?? null,
    },
  };
  
  notificationClients.forEach((client) => {
    if (userIds.includes(client.user.id)) {
      sendNotificationPayload(client, notificationPayload);
    }
  });
}

async function notifyAdmins(payload: { kind: string; message: string; href?: string | null }) {
  const adminIds = await listActiveUserIdsByRole(["ADMIN"]);
  await notifyUsers(adminIds, payload);
}

async function notifyAllUsers(payload: { kind: string; message: string; href?: string | null }) {
  const userIds = await listActiveUserIds();
  await notifyUsers(userIds, payload);
}

function readWsToken(req: any) {
  const header = req.headers?.authorization;
  if (typeof header === "string" && header.startsWith("Bearer ")) {
    return header.slice("Bearer ".length);
  }
  const query = req.query as { token?: string } | undefined;
  if (query?.token && typeof query.token === "string") {
    return query.token;
  }
  const rawUrl = req.raw?.url;
  if (typeof rawUrl === "string" && rawUrl.includes("?")) {
    const qs = rawUrl.split("?")[1] ?? "";
    const params = new URLSearchParams(qs);
    const token = params.get("token");
    if (token) return token;
  }
  return undefined;
}

function sendCommunityPayload(
  client: CommunityWsClient,
  payload: Record<string, unknown>
) {
  try {
    const socket = client.socket;
    if (typeof socket.send !== "function") return;
    if (typeof socket.readyState === "number" && socket.readyState !== 1)
      return;
    socket.send(JSON.stringify(payload));
  } catch {
    // ignore websocket send errors
  }
}

function sendNotificationPayload(
  client: NotificationWsClient,
  payload: Record<string, unknown>
) {
  try {
    const socket = client.socket;
    if (typeof socket.send !== "function") return;
    if (typeof socket.readyState === "number" && socket.readyState !== 1)
      return;
    socket.send(JSON.stringify(payload));
  } catch {
    // ignore websocket send errors
  }
}

async function broadcastCommunityMessage(
  threadId: string,
  message: CommunityMessage
) {
  const thread = await findCommunityThreadById(threadId);
  if (!thread) return;
  const payload = {
    type: "community_message",
    threadId,
    threadType: thread.threadType,
    message,
  };
  if (thread.threadType === "CHANNEL" && !thread.isPrivate) {
    app.log.info(
      { threadId, clients: communityClients.size },
      "community broadcast channel"
    );
    communityClients.forEach((client) => sendCommunityPayload(client, payload));
    return;
  }
  const memberIds = await listCommunityThreadMemberIds(threadId);
  if (!memberIds.length) return;
  const allowed = new Set(memberIds);
  app.log.info(
    { threadId, recipients: allowed.size, clients: communityClients.size },
    "community broadcast dm"
  );
  communityClients.forEach((client) => {
    if (allowed.has(client.user.id)) {
      sendCommunityPayload(client, payload);
    }
  });
}

function mergeBaseInfo(
  existing?: BaseInfo,
  incoming?: Partial<BaseInfo>
): BaseInfo {
  const current = existing ?? {};
  const next = incoming ?? {};
  const merged: BaseInfo = {
    ...current,
    ...next,
    name: { ...(current.name ?? {}), ...(next.name ?? {}) },
    contact: { ...(current.contact ?? {}), ...(next.contact ?? {}) },
    location: { ...(current.location ?? {}), ...(next.location ?? {}) },
    workAuth: { ...(current.workAuth ?? {}), ...(next.workAuth ?? {}) },
    links: { ...(current.links ?? {}), ...(next.links ?? {}) },
    career: { ...(current.career ?? {}), ...(next.career ?? {}) },
    education: { ...(current.education ?? {}), ...(next.education ?? {}) },
    preferences: {
      ...(current.preferences ?? {}),
      ...(next.preferences ?? {}),
    },
    defaultAnswers: {
      ...(current.defaultAnswers ?? {}),
      ...(next.defaultAnswers ?? {}),
    },
  };
  const phone = formatPhone(merged.contact);
  if (phone) {
    merged.contact = { ...(merged.contact ?? {}), phone };
  }
  return merged;
}

function parseSalaryNumber(input?: string | number | null) {
  if (typeof input === "number" && Number.isFinite(input)) return input;
  if (typeof input !== "string") return undefined;
  const cleaned = input.replace(/[, ]+/g, "").replace(/[^0-9.]/g, "");
  if (!cleaned) return undefined;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function computeHourlyRate(desiredSalary?: string | number | null) {
  const annual = parseSalaryNumber(desiredSalary);
  if (!annual || annual <= 0) return undefined;
  return Math.floor(annual / 12 / 160);
}

function buildAutofillValueMap(
  baseInfo: BaseInfo,
  jobContext?: Record<string, unknown>
): Record<string, string> {
  const firstName = trimString(baseInfo?.name?.first);
  const lastName = trimString(baseInfo?.name?.last);
  const fullName = [firstName, lastName].filter(Boolean).join(" ").trim();
  const email = trimString(baseInfo?.contact?.email);
  const phoneCode = trimString(baseInfo?.contact?.phoneCode);
  const phoneNumber = trimString(baseInfo?.contact?.phoneNumber);
  const formattedPhone =
    phoneCode && phoneNumber
      ? `${phoneCode} ${phoneNumber}`.trim()
      : formatPhone(baseInfo?.contact);
  const address = trimString(baseInfo?.location?.address);
  const city = trimString(baseInfo?.location?.city);
  const state = trimString(baseInfo?.location?.state);
  const country = trimString(baseInfo?.location?.country);
  const postalCode = trimString(baseInfo?.location?.postalCode);
  const linkedin = trimString(baseInfo?.links?.linkedin);
  const jobTitle =
    trimString(baseInfo?.career?.jobTitle) ||
    trimString((jobContext as any)?.job_title);
  const currentCompany =
    trimString(baseInfo?.career?.currentCompany) ||
    trimString((jobContext as any)?.company) ||
    trimString((jobContext as any)?.employer);
  const yearsExp = trimString(baseInfo?.career?.yearsExp);
  const desiredSalary = trimString(baseInfo?.career?.desiredSalary);
  const hourlyRate = computeHourlyRate(desiredSalary);
  const school = trimString(baseInfo?.education?.school);
  const degree = trimString(baseInfo?.education?.degree);
  const majorField = trimString(baseInfo?.education?.majorField);
  const graduationDate = trimString(baseInfo?.education?.graduationAt);
  const currentLocation = [city, state, country].filter(Boolean).join(", ");
  const phoneCountryCode =
    phoneCode ||
    (formattedPhone.startsWith("+")
      ? formattedPhone.split(/\s+/)[0]
      : trimString(baseInfo?.contact?.phone));

  const values: Record<string, string> = {
    full_name: fullName,
    first_name: firstName,
    last_name: lastName,
    preferred_name: firstName || fullName,
    pronouns: "Mr",
    email,
    phone: formattedPhone,
    phone_country_code: phoneCountryCode,
    address_line1: address,
    city,
    state_province_region: state,
    postal_code: postalCode,
    country,
    current_location: currentLocation,
    linkedin_url: linkedin,
    job_title: jobTitle,
    current_company: currentCompany,
    years_experience: yearsExp,
    desired_salary: desiredSalary,
    hourly_rate: hourlyRate !== undefined ? String(hourlyRate) : "",
    start_date: "immediately",
    notice_period: "0",
    school,
    degree,
    major_field: majorField,
    graduation_date: graduationDate,
    eeo_gender: "man",
    eeo_race_ethnicity: "white",
    eeo_veteran: "no veteran",
    eeo_disability: "no disability",
  };
  return values;
}

async function collectPageFieldsFromFrame(
  frame: Frame,
  meta: { frameUrl: string; frameName: string }
) {
  return frame.evaluate(
    (frameInfo) => {
      const norm = (s?: string | null) => (s || "").replace(/\s+/g, " ").trim();
      const textOf = (el?: Element | null) =>
        norm(el?.textContent || (el as HTMLElement | null)?.innerText || "");
      const isVisible = (el: Element) => {
        const cs = window.getComputedStyle(el);
        if (!cs || cs.display === "none" || cs.visibility === "hidden")
          return false;
        const r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0;
      };
      const esc = (v: string) =>
        window.CSS && CSS.escape
          ? CSS.escape(v)
          : v.replace(/[^a-zA-Z0-9_-]/g, "\\$&");

      const getLabelText = (el: Element) => {
        try {
          const labels = (el as HTMLInputElement).labels;
          if (labels && labels.length) {
            const t = Array.from(labels)
              .map((n) => textOf(n))
              .filter(Boolean);
            if (t.length) return t.join(" ");
          }
        } catch {
          /* ignore */
        }
        const id = el.getAttribute("id");
        if (id) {
          const lab = document.querySelector(`label[for="${esc(id)}"]`);
          const t = textOf(lab);
          if (t) return t;
        }
        const wrap = el.closest("label");
        const t2 = textOf(wrap);
        return t2 || "";
      };

      const getAriaName = (el: Element) => {
        const direct = norm(el.getAttribute("aria-label"));
        if (direct) return direct;
        const labelledBy = norm(el.getAttribute("aria-labelledby"));
        if (labelledBy) {
          const parts = labelledBy
            .split(/\s+/)
            .map((id) => textOf(document.getElementById(id)))
            .filter(Boolean);
          return norm(parts.join(" "));
        }
        return "";
      };

      const getDescribedBy = (el: Element) => {
        const ids = norm(el.getAttribute("aria-describedby"));
        if (!ids) return "";
        const parts = ids
          .split(/\s+/)
          .map((id) => textOf(document.getElementById(id)))
          .filter(Boolean);
        return norm(parts.join(" "));
      };

      const findFieldContainer = (el: Element) =>
        el.closest(
          "fieldset, [role='group'], .form-group, .field, .input-group, .question, .formField, section, article, li, div"
        ) || el.parentElement;

      const collectNearbyPrompts = (el: Element) => {
        const container = findFieldContainer(el);
        if (!container) return [];

        const prompts: { source: string; text: string }[] = [];

        const fieldset = el.closest("fieldset");
        if (fieldset) {
          const legend = fieldset.querySelector("legend");
          const t = textOf(legend);
          if (t) prompts.push({ source: "legend", text: t });
        }

        const candidates = container.querySelectorAll(
          "h1,h2,h3,h4,h5,h6,p,.help,.hint,.description,[data-help],[data-testid*='help']"
        );
        candidates.forEach((n: Element) => {
          const t = textOf(n);
          if (t && t.length <= 350)
            prompts.push({ source: "container_text", text: t });
        });

        let sib: Element | null = el.previousElementSibling;
        let steps = 0;
        while (sib && steps < 4) {
          const tag = sib.tagName.toLowerCase();
          if (
            [
              "div",
              "p",
              "span",
              "h1",
              "h2",
              "h3",
              "h4",
              "h5",
              "h6",
              "label",
            ].includes(tag)
          ) {
            const t = textOf(sib);
            if (t && t.length <= 350)
              prompts.push({ source: "prev_sibling", text: t });
          }
          sib = sib.previousElementSibling;
          steps += 1;
        }

        return prompts;
      };

      const looksBoilerplate = (t: string) => {
        const s = t.toLowerCase();
        return (
          s.includes("privacy") ||
          s.includes("terms") ||
          s.includes("cookies") ||
          s.includes("equal opportunity") ||
          s.includes("eeo") ||
          s.includes("gdpr")
        );
      };

      const scorePrompt = (text: string, source: string) => {
        const s = text.toLowerCase();
        let score = 0;
        if (text.includes("?")) score += 6;
        if (
          /^(why|how|what|describe|explain|tell us|please describe|please explain)\b/i.test(
            text
          )
        )
          score += 4;
        if (
          /(position|role|motivation|interested|interest|experience|background|cover letter)/i.test(
            text
          )
        )
          score += 2;
        if (text.length >= 20 && text.length <= 220) score += 3;
        if (source === "label" || source === "aria") score += 5;
        if (source === "describedby") score += 3;
        if (text.length > 350) score -= 4;
        if (looksBoilerplate(text)) score -= 6;
        if (/^(optional|required)\b/i.test(text)) score -= 5;
        if (s === "optional" || s === "required") score -= 5;
        return score;
      };

      const parseTextConstraints = (text: string) => {
        const t = text.toLowerCase();
        const out: Record<string, number> = {};
        const words = t.match(/max(?:imum)?\s*(\d+)\s*words?/);
        if (words) out.max_words = parseInt(words[1], 10);
        const chars = t.match(/max(?:imum)?\s*(\d+)\s*(characters|chars)/);
        if (chars) out.max_chars = parseInt(chars[1], 10);
        const minChars = t.match(/min(?:imum)?\s*(\d+)\s*(characters|chars)/);
        if (minChars) out.min_chars = parseInt(minChars[1], 10);
        return out;
      };

      const recommendedLocators = (el: Element, bestLabel?: string | null) => {
        const tag = el.tagName.toLowerCase();
        const id = el.getAttribute("id");
        const name = el.getAttribute("name");
        const placeholder = el.getAttribute("placeholder");
        const locators: Record<string, string> = {};
        if (id) locators.css = `#${esc(id)}`;
        else if (name) locators.css = `${tag}[name="${esc(name)}"]`;
        else locators.css = tag;

        if (bestLabel)
          locators.playwright = `getByLabel(${JSON.stringify(bestLabel)})`;
        else if (placeholder)
          locators.playwright = `getByPlaceholder(${JSON.stringify(
            placeholder
          )})`;
        else locators.playwright = `locator(${JSON.stringify(locators.css)})`;
        return locators;
      };

      const slug = (v: string) =>
        norm(v)
          .toLowerCase()
          .replace(/[^a-z0-9]+/g, "_")
          .replace(/^_+|_+$/g, "");

      const controls = Array.from(
        document.querySelectorAll(
          'input, textarea, select, [contenteditable="true"], [role="textbox"]'
        )
      ).slice(0, 80);

      const fields: any[] = [];
      controls.forEach((el: Element, idx) => {
        const tag = el.tagName.toLowerCase();
        if (tag === "input") {
          const t = (
            (el as HTMLInputElement).type ||
            el.getAttribute("type") ||
            "text"
          ).toLowerCase();
          if (["hidden", "submit", "button", "image", "reset"].includes(t))
            return;
        }
        if (!isVisible(el)) return;

        const label = norm(getLabelText(el));
        const ariaName = norm(getAriaName(el));
        const describedBy = norm(getDescribedBy(el));
        const placeholder = norm(el.getAttribute("placeholder"));
        const autocomplete = norm(el.getAttribute("autocomplete"));
        const name = norm(el.getAttribute("name"));
        const id = norm(el.getAttribute("id"));
        const required = Boolean((el as HTMLInputElement).required);

        const type =
          tag === "input"
            ? (
                norm(
                  (el as HTMLInputElement).type || el.getAttribute("type")
                ) || "text"
              ).toLowerCase()
            : tag === "textarea"
            ? "textarea"
            : tag === "select"
            ? "select"
            : el.getAttribute("role") === "textbox" ||
              el.getAttribute("contenteditable") === "true"
            ? "richtext"
            : tag;

        const promptCandidates: {
          source: string;
          text: string;
          score: number;
        }[] = [];
        if (label)
          promptCandidates.push({
            source: "label",
            text: label,
            score: scorePrompt(label, "label") + 8,
          });
        if (ariaName)
          promptCandidates.push({
            source: "aria",
            text: ariaName,
            score: scorePrompt(ariaName, "aria"),
          });
        if (placeholder) {
          promptCandidates.push({
            source: "placeholder",
            text: placeholder,
            score: scorePrompt(placeholder, "placeholder"),
          });
        }
        if (describedBy) {
          promptCandidates.push({
            source: "describedby",
            text: describedBy,
            score: scorePrompt(describedBy, "describedby"),
          });
        }
        const nearbyPrompts = collectNearbyPrompts(el);
        nearbyPrompts.forEach((p) => {
          promptCandidates.push({ ...p, score: scorePrompt(p.text, p.source) });
        });

        const best =
          label && promptCandidates.find((p) => p.source === "label")
            ? promptCandidates.find((p) => p.source === "label")
            : promptCandidates
                .filter((p) => p.text)
                .sort((a, b) => b.score - a.score)[0];
        const questionText = best?.text || "";
        const locators = recommendedLocators(
          el,
          label || ariaName || questionText || placeholder
        );

        const constraints: Record<string, number> = {};
        const maxlen = el.getAttribute("maxlength");
        const minlen = el.getAttribute("minlength");
        if (maxlen) constraints.maxlength = parseInt(maxlen, 10);
        if (minlen) constraints.minlength = parseInt(minlen, 10);
        Object.assign(
          constraints,
          parseTextConstraints(`${questionText} ${describedBy}`)
        );

        const textForEssay =
          `${questionText} ${label} ${describedBy}`.toLowerCase();
        const likelyEssay =
          type === "textarea" ||
          type === "richtext" ||
          Boolean(constraints.max_words) ||
          Boolean(constraints.max_chars && constraints.max_chars > 180) ||
          (/why|tell us|describe|explain|motivation|interest|cover letter|statement/.test(
            textForEssay
          ) &&
            (questionText.length > 0 || label.length > 0));

        const fallbackId =
          slug(
            label || ariaName || questionText || placeholder || name || ""
          ) || `field_${idx}`;
        const fieldId = id || name || fallbackId;

        fields.push({
          index: fields.length,
          field_id: fieldId,
          tag,
          type,
          id: id || null,
          name: name || null,
          label: label || null,
          ariaName: ariaName || null,
          placeholder: placeholder || null,
          describedBy: describedBy || null,
          autocomplete: autocomplete || null,
          required,
          visible: true,
          questionText: questionText || null,
          questionCandidates: promptCandidates
            .sort((a, b) => b.score - a.score)
            .slice(0, 5),
          constraints,
          locators: {
            css: locators.css,
            playwright: locators.playwright,
          },
          selector: locators.css,
          likelyEssay,
          containerPrompts: nearbyPrompts,
          frameUrl: frameInfo.frameUrl,
          frameName: frameInfo.frameName,
        });
      });

      return fields;
    },
    { frameUrl: meta.frameUrl, frameName: meta.frameName }
  );
}

async function collectPageFields(page: Page) {
  const frames = page.frames();
  const results = await Promise.all(
    frames.map(async (frame, idx) => {
      try {
        return await collectPageFieldsFromFrame(frame, {
          frameUrl: frame.url(),
          frameName: frame.name() || `frame-${idx}`,
        });
      } catch (err) {
        console.error("collectPageFields frame failed", err);
        return [];
      }
    })
  );
  const merged = results.flat();
  if (merged.length) return merged.slice(0, 300);

  // fallback to main frame attempt
  try {
    return await collectPageFieldsFromFrame(page.mainFrame(), {
      frameUrl: page.mainFrame().url(),
      frameName: page.mainFrame().name() || "main",
    });
  } catch {
    return [];
  }
}

async function applyFillPlan(page: Page, plan: any[]): Promise<FillPlanResult> {
  const filled: { field: string; value: string; confidence?: number }[] = [];
  const suggestions: { field: string; suggestion: string }[] = [];
  const blocked: string[] = [];

  for (const f of plan) {
    const action = f.action;
    const value = f.value;
    const selector =
      f.selector ||
      (f.field_id
        ? `[name="${f.field_id}"], #${f.field_id}, [id*="${f.field_id}"]`
        : undefined);
    if (!selector) {
      blocked.push(f.field_id ?? "field");
      continue;
    }
    try {
      if (action === "fill") {
        await page.fill(
          selector,
          typeof value === "string" ? value : String(value ?? "")
        );
        filled.push({
          field: f.field_id ?? selector,
          value:
            typeof value === "string" ? value : JSON.stringify(value ?? ""),
          confidence:
            typeof f.confidence === "number" ? f.confidence : undefined,
        });
      } else if (action === "select") {
        await page.selectOption(selector, { label: String(value ?? "") });
        filled.push({
          field: f.field_id ?? selector,
          value: String(value ?? ""),
          confidence:
            typeof f.confidence === "number" ? f.confidence : undefined,
        });
      } else if (action === "check" || action === "uncheck") {
        if (action === "check") await page.check(selector);
        else await page.uncheck(selector);
        filled.push({ field: f.field_id ?? selector, value: action });
      } else if (f.requires_user_review) {
        blocked.push(f.field_id ?? selector);
      }
    } catch {
      blocked.push(f.field_id ?? selector);
    }
  }
  return { filled, suggestions, blocked };
}

function collectLabelCandidates(field: any): string[] {
  const candidates: string[] = [];
  const primaryPrompt =
    Array.isArray(field?.questionCandidates) &&
    field.questionCandidates.length > 0
      ? field.questionCandidates[0].text
      : undefined;
  [
    primaryPrompt,
    field?.questionText,
    field?.label,
    field?.ariaName,
    field?.placeholder,
    field?.describedBy,
    field?.field_id,
    field?.name,
    field?.id,
  ].forEach((t) => {
    if (typeof t === "string" && t.trim()) candidates.push(t);
  });
  if (Array.isArray(field?.containerPrompts)) {
    field.containerPrompts.forEach((p: any) => {
      if (p?.text && typeof p.text === "string" && p.text.trim())
        candidates.push(p.text);
    });
  }
  return candidates;
}

function escapeCssValue(value: string) {
  return value.replace(/["\\]/g, "\\$&");
}

function escapeCssIdent(value: string) {
  return value.replace(/[^a-zA-Z0-9_-]/g, "\\$&");
}

function buildFieldSelector(field: any): string | undefined {
  if (field?.selector && typeof field.selector === "string")
    return field.selector;
  if (field?.locators?.css && typeof field.locators.css === "string")
    return field.locators.css;
  if (field?.id) return `#${escapeCssIdent(String(field.id))}`;
  if (field?.field_id)
    return `[name="${escapeCssValue(String(field.field_id))}"]`;
  if (field?.name) return `[name="${escapeCssValue(String(field.name))}"]`;
  return undefined;
}

function inferFieldAction(field: any): FillPlanAction["action"] {
  const rawType = String(field?.type ?? "").toLowerCase();
  if (rawType === "select") return "select";
  return "fill";
}

const SKIP_KEYS = new Set(["cover_letter"]);

function buildAliasFillPlan(
  fields: any[],
  aliasIndex: Map<string, string>,
  valueMap: Record<string, string>
): FillPlanResult {
  const filled: { field: string; value: string; confidence?: number }[] = [];
  const suggestions: { field: string; suggestion: string }[] = [];
  const blocked: string[] = [];
  const actions: FillPlanAction[] = [];
  const seen = new Set<string>();

  for (const field of fields ?? []) {
    const candidates = collectLabelCandidates(field);
    let matchedKey: string | undefined;
    let matchedLabel = "";
    for (const c of candidates) {
      const match = matchLabelToCanonical(c, aliasIndex);
      if (match) {
        matchedKey = match;
        matchedLabel = c;
        break;
      }
    }
    if (!matchedKey) continue;
    if (SKIP_KEYS.has(matchedKey)) continue;

    const value = trimString(valueMap[matchedKey]);
    const fieldName =
      trimString(
        field?.field_id ||
          field?.name ||
          field?.id ||
          matchedLabel ||
          matchedKey
      ) || matchedKey;
    if (seen.has(fieldName)) continue;
    seen.add(fieldName);

    if (!value) {
      suggestions.push({
        field: fieldName,
        suggestion: `No data available for ${matchedKey}`,
      });
      continue;
    }

    const selector = buildFieldSelector(field);
    const fieldId = trimString(field?.field_id || field?.name || field?.id);
    const fieldLabel = trimString(
      matchedLabel ||
        field?.label ||
        field?.questionText ||
        field?.ariaName ||
        fieldName
    );
    if (!selector) {
      blocked.push(fieldName);
      continue;
    }
    const action = inferFieldAction(field);
    actions.push({
      field: fieldName,
      field_id: fieldId || undefined,
      label: fieldLabel || undefined,
      selector,
      action,
      value,
      confidence: 0.75,
    });
    filled.push({ field: fieldName, value, confidence: 0.75 });
  }

  return { filled, suggestions, blocked, actions };
}

function shouldSkipPlanField(field: any, aliasIndex: Map<string, string>) {
  const candidates = [field?.field_id, field?.label, field?.selector].filter(
    (c) => typeof c === "string" && c.trim()
  );
  for (const c of candidates) {
    const match = matchLabelToCanonical(String(c), aliasIndex);
    if (match && SKIP_KEYS.has(match)) return true;
  }
  return false;
}

async function simplePageFill(
  page: Page,
  baseInfo: BaseInfo
): Promise<FillPlanResult> {
  const fullName = [baseInfo?.name?.first, baseInfo?.name?.last]
    .filter(Boolean)
    .join(" ")
    .trim();
  const email = trimString(baseInfo?.contact?.email);
  const phoneCode = trimString(baseInfo?.contact?.phoneCode);
  const phoneNumber = trimString(baseInfo?.contact?.phoneNumber);
  const phone = formatPhone(baseInfo?.contact);
  const address = trimString(baseInfo?.location?.address);
  const city = trimString(baseInfo?.location?.city);
  const state = trimString(baseInfo?.location?.state);
  const country = trimString(baseInfo?.location?.country);
  const postalCode = trimString(baseInfo?.location?.postalCode);
  const linkedin = trimString(baseInfo?.links?.linkedin);
  const company = trimString(baseInfo?.career?.currentCompany);
  const title = trimString(baseInfo?.career?.jobTitle);
  const yearsExp = trimString(baseInfo?.career?.yearsExp);
  const desiredSalary = trimString(baseInfo?.career?.desiredSalary);
  const school = trimString(baseInfo?.education?.school);
  const degree = trimString(baseInfo?.education?.degree);
  const majorField = trimString(baseInfo?.education?.majorField);
  const graduationAt = trimString(baseInfo?.education?.graduationAt);

  const filled: { field: string; value: string; confidence?: number }[] = [];
  const targets = [
    { key: "full_name", match: /full\s*name/i, value: fullName },
    { key: "first", match: /first/i, value: baseInfo?.name?.first },
    { key: "last", match: /last/i, value: baseInfo?.name?.last },
    { key: "email", match: /email/i, value: email },
    {
      key: "phone_code",
      match: /(phone|mobile).*(code)|country\s*code|dial\s*code/i,
      value: phoneCode,
    },
    {
      key: "phone_number",
      match: /(phone|mobile).*(number|no\.)/i,
      value: phoneNumber,
    },
    { key: "phone", match: /phone|tel/i, value: phone },
    { key: "address", match: /address/i, value: address },
    { key: "city", match: /city/i, value: city },
    { key: "state", match: /state|province|region/i, value: state },
    { key: "country", match: /country|nation/i, value: country },
    { key: "postal_code", match: /postal|zip/i, value: postalCode },
    { key: "company", match: /company|employer/i, value: company },
    { key: "title", match: /title|position|role/i, value: title },
    {
      key: "years_experience",
      match: /years?.*experience|experience.*years|yrs/i,
      value: yearsExp,
    },
    {
      key: "desired_salary",
      match: /salary|compensation|pay|rate/i,
      value: desiredSalary,
    },
    { key: "linkedin", match: /linkedin|linked\s*in/i, value: linkedin },
    { key: "school", match: /school|university|college/i, value: school },
    { key: "degree", match: /degree|diploma/i, value: degree },
    {
      key: "major_field",
      match: /major|field\s*of\s*study/i,
      value: majorField,
    },
    { key: "graduation_at", match: /grad/i, value: graduationAt },
  ].filter((t) => t.value);

  const inputs = await page.$$("input, textarea, select");
  for (const el of inputs) {
    const props = await el.evaluate((node) => {
      const lbl = (node as HTMLInputElement).labels?.[0]?.innerText || "";
      return {
        tag: node.tagName.toLowerCase(),
        type:
          (node as HTMLInputElement).type ||
          node.getAttribute("type") ||
          "text",
        name: node.getAttribute("name") || "",
        id: node.id || "",
        placeholder: node.getAttribute("placeholder") || "",
        label: lbl,
      };
    });
    if (
      props.type === "checkbox" ||
      props.type === "radio" ||
      props.type === "file"
    )
      continue;
    const haystack =
      `${props.label} ${props.name} ${props.id} ${props.placeholder}`.toLowerCase();
    const match = targets.find((t) => t.match.test(haystack));
    if (match) {
      const val = String(match.value ?? "");
      try {
        if (props.tag === "select") {
          await el.selectOption({ label: val });
        } else {
          await el.fill(val);
        }
        filled.push({ field: props.name || props.id || match.key, value: val });
      } catch {
        // ignore failed fills
      }
    }
  }
  return { filled, suggestions: [], blocked: [] };
}

const DEFAULT_AUTOFILL_FIELDS = [
  { field_id: "first_name", label: "First name", type: "text", required: true },
  { field_id: "last_name", label: "Last name", type: "text", required: true },
  { field_id: "email", label: "Email", type: "text", required: true },
  {
    field_id: "phone_code",
    label: "Phone code",
    type: "text",
    required: false,
  },
  {
    field_id: "phone_number",
    label: "Phone number",
    type: "text",
    required: false,
  },
  { field_id: "phone", label: "Phone", type: "text", required: false },
  { field_id: "address", label: "Address", type: "text", required: false },
  { field_id: "city", label: "City", type: "text", required: false },
  { field_id: "state", label: "State/Province", type: "text", required: false },
  { field_id: "country", label: "Country", type: "text", required: false },
  {
    field_id: "postal_code",
    label: "Postal code",
    type: "text",
    required: false,
  },
  { field_id: "linkedin", label: "LinkedIn", type: "text", required: false },
  { field_id: "job_title", label: "Job title", type: "text", required: false },
  {
    field_id: "current_company",
    label: "Current company",
    type: "text",
    required: false,
  },
  {
    field_id: "years_exp",
    label: "Years of experience",
    type: "number",
    required: false,
  },
  {
    field_id: "desired_salary",
    label: "Desired salary",
    type: "text",
    required: false,
  },
  { field_id: "school", label: "School", type: "text", required: false },
  { field_id: "degree", label: "Degree", type: "text", required: false },
  {
    field_id: "major_field",
    label: "Major/Field",
    type: "text",
    required: false,
  },
  {
    field_id: "graduation_at",
    label: "Graduation date",
    type: "text",
    required: false,
  },
  {
    field_id: "work_auth",
    label: "Authorized to work?",
    type: "checkbox",
    required: false,
  },
];

// initDb, auth guard, signToken live in dedicated modules

async function bootstrap() {
  await app.register(authGuard);
  await app.register(cors, { origin: config.CORS_ORIGINS, credentials: true });
  await app.register(websocket);
  await app.register(multipart, {
    limits: {
      fileSize: 10 * 1024 * 1024, // 10MB
    },
  });
  await initDb();
  void startScraperService();
  await registerScraperApiRoutes(app);

  app.get("/health", async () => ({ status: "ok" }));

  app.post("/auth/login", async (request, reply) => {
    const schema = z.object({
      email: z.string().email(),
      password: z.string().optional(),
    });
    const parsed = schema.safeParse(request.body ?? {});
    if (!parsed.success) {
      const issue = parsed.error.errors[0];
      const field = issue?.path?.[0];
      const message = `${field ? `${field}: ` : ""}${issue?.message ?? "Invalid login payload"}`;
      return reply.status(400).send({ message });
    }
    const body = parsed.data;
    const user = await findUserByEmail(body.email);
    if (!user) {
      return reply.status(401).send({ message: "Invalid credentials" });
    }
    if (
      user.password &&
      body.password &&
      !(await bcrypt.compare(body.password, user.password))
    ) {
      return reply.status(401).send({ message: "Invalid credentials" });
    }
    if (user.isActive === false || user.role === "NONE") {
      return reply.status(403).send({ message: "Account is pending approval. Please wait for admin approval." });
    }
    const token = signToken(user);
    return { token, user };
  });

  app.post("/auth/signup", async (request, reply) => {
    const schema = z.object({
      email: z.string().email(),
      password: z.string().min(3),
      userName: z.string().min(2).max(50).regex(/^[a-zA-Z0-9_-]+$/),
      avatarUrl: z.string().trim().optional(),
    });
    const parsed = schema.safeParse(request.body ?? {});
    if (!parsed.success) {
      const issue = parsed.error.errors[0];
      const field = issue?.path?.[0];
      const message = `${field ? `${field}: ` : ""}${issue?.message ?? "Invalid signup payload"}`;
      return reply.status(400).send({ message });
    }
    const body = parsed.data;
    const exists = await findUserByEmail(body.email);
    if (exists) {
      return reply.status(409).send({ message: "Email already registered" });
    }
    const userNameExists = await findUserByUserName(body.userName);
    if (userNameExists) {
      return reply.status(409).send({ message: "User name already taken" });
    }
    const hashed = await bcrypt.hash(body.password, 8);
    const normalizedAvatar =
      body.avatarUrl && body.avatarUrl.toLowerCase() !== "nope"
        ? body.avatarUrl
        : null;
    // Use userName as the default name if not provided
    const user: User = {
      id: randomUUID(),
      email: body.email,
      userName: body.userName,
      role: "NONE",
      name: body.userName,
      avatarUrl: normalizedAvatar,
      isActive: false,
      password: hashed,
    };
    await insertUser(user);
    try {
      await notifyAdmins({
        kind: "system",
        message: `New join request from ${user.userName}.`,
        href: "/admin/join-requests",
      });
    } catch (err) {
      request.log.error({ err }, "join request notification failed");
    }
    return { message: "Account created. Please wait for admin approval.", user: { id: user.id, email: user.email, userName: user.userName, name: user.name, role: user.role } };
  });

  app.get("/profiles", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }

    const { userId } = request.query as { userId?: string };

    if (actor.role === "ADMIN" || actor.role === "MANAGER") {
      if (userId) {
        const target = await findUserById(userId);
        if (target?.role === "BIDDER" && target.isActive !== false) {
          return listProfilesForBidder(target.id);
        }
      }
      return listProfiles();
    }

    if (actor.role === "BIDDER") {
      return listProfilesForBidder(actor.id);
    }

    return reply.status(403).send({ message: "Forbidden" });
  });

  app.post("/profiles", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can create profiles" });
    }
    const schema = z.object({
      displayName: z.string().min(2),
      baseInfo: z.record(z.any()).optional(),
      baseResume: z.record(z.any()).optional(),
      baseAdditionalBullets: z.record(z.number()).optional(),
      firstName: z.string().optional(),
      lastName: z.string().optional(),
      email: z.string().email().optional(),
      phoneCode: z.string().optional(),
      phoneNumber: z.string().optional(),
      address: z.string().optional(),
      city: z.string().optional(),
      state: z.string().optional(),
      country: z.string().optional(),
      postalCode: z.string().optional(),
      linkedin: z.string().optional(),
      jobTitle: z.string().optional(),
      currentCompany: z.string().optional(),
      yearsExp: z.union([z.string(), z.number()]).optional(),
      desiredSalary: z.string().optional(),
      school: z.string().optional(),
      degree: z.string().optional(),
      majorField: z.string().optional(),
      graduationAt: z.string().optional(),
      resumeTemplateId: z.string().uuid().optional().nullable(),
    });
    const body = schema.parse(request.body);
    const profileId = randomUUID();
    const now = new Date().toISOString();
    const resumeTemplateId = body.resumeTemplateId ? trimString(body.resumeTemplateId) : null;
    let resumeTemplateName: string | null = null;
    if (resumeTemplateId) {
      const template = await findResumeTemplateById(resumeTemplateId);
      if (!template) {
        return reply.status(400).send({ message: "Resume template not found" });
      }
      resumeTemplateName = template.name;
    }
    const incomingBase = (body.baseInfo ?? {}) as BaseInfo;
    const baseResume = (body.baseResume ?? {}) as Record<string, unknown>;
    const baseInfo = mergeBaseInfo(
      {},
      {
        ...incomingBase,
        name: {
          ...(incomingBase.name ?? {}),
          first: trimString(body.firstName ?? incomingBase.name?.first),
          last: trimString(body.lastName ?? incomingBase.name?.last),
        },
        contact: {
          ...(incomingBase.contact ?? {}),
          email: trimString(body.email ?? incomingBase.contact?.email),
          phoneCode: trimString(
            body.phoneCode ?? incomingBase.contact?.phoneCode
          ),
          phoneNumber: trimString(
            body.phoneNumber ?? incomingBase.contact?.phoneNumber
          ),
        },
        location: {
          ...(incomingBase.location ?? {}),
          address: trimString(body.address ?? incomingBase.location?.address),
          city: trimString(body.city ?? incomingBase.location?.city),
          state: trimString(body.state ?? incomingBase.location?.state),
          country: trimString(body.country ?? incomingBase.location?.country),
          postalCode: trimString(
            body.postalCode ?? incomingBase.location?.postalCode
          ),
        },
        links: {
          ...(incomingBase.links ?? {}),
          linkedin: trimString(
            body.linkedin ?? (incomingBase.links as any)?.linkedin
          ),
        },
        career: {
          ...(incomingBase.career ?? {}),
          jobTitle: trimString(body.jobTitle ?? incomingBase.career?.jobTitle),
          currentCompany: trimString(
            body.currentCompany ?? incomingBase.career?.currentCompany
          ),
          yearsExp: body.yearsExp ?? incomingBase.career?.yearsExp,
          desiredSalary: trimString(
            body.desiredSalary ?? incomingBase.career?.desiredSalary
          ),
        },
        education: {
          ...(incomingBase.education ?? {}),
          school: trimString(body.school ?? incomingBase.education?.school),
          degree: trimString(body.degree ?? incomingBase.education?.degree),
          majorField: trimString(
            body.majorField ?? incomingBase.education?.majorField
          ),
          graduationAt: trimString(
            body.graduationAt ?? incomingBase.education?.graduationAt
          ),
        },
      }
    );
    const profile = {
      id: profileId,
      displayName: body.displayName,
      baseInfo,
      baseResume,
      baseAdditionalBullets: body.baseAdditionalBullets ?? {},
      resumeTemplateId: resumeTemplateId ?? null,
      resumeTemplateName,
      createdBy: actor.id,
      createdAt: now,
      updatedAt: now,
    };
    await insertProfile(profile);
    try {
      await notifyAdmins({
        kind: "system",
        message: `New profile ${profile.displayName} created.`,
        href: "/manager/profiles",
      });
    } catch (err) {
      request.log.error({ err }, "profile create notification failed");
    }
    return profile;
  });

  app.patch("/profiles/:id", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can update profiles" });
    }
    const { id } = request.params as { id: string };
    const existing = await findProfileById(id);
    if (!existing)
      return reply.status(404).send({ message: "Profile not found" });

    const schema = z.object({
      displayName: z.string().min(2).optional(),
      baseInfo: z.record(z.any()).optional(),
      baseResume: z.record(z.any()).optional(),
      baseAdditionalBullets: z.record(z.number()).optional(),
      resumeTemplateId: z.string().uuid().optional().nullable(),
    });
    const body = schema.parse(request.body ?? {});

    let resumeTemplateId = existing.resumeTemplateId ?? null;
    let resumeTemplateName = existing.resumeTemplateName ?? null;
    if (body.resumeTemplateId !== undefined) {
      resumeTemplateId = trimString(body.resumeTemplateId) || null;
      if (resumeTemplateId) {
        const template = await findResumeTemplateById(resumeTemplateId);
        if (!template) {
          return reply.status(400).send({ message: "Resume template not found" });
        }
        resumeTemplateName = template.name;
      } else {
        resumeTemplateName = null;
      }
    }

    const incomingBase = (body.baseInfo ?? {}) as BaseInfo;
    const mergedBase = mergeBaseInfo(existing.baseInfo, incomingBase);
    const baseResume = (body.baseResume ?? existing.baseResume ?? {}) as Record<string, unknown>;
    const baseAdditionalBullets = body.baseAdditionalBullets ?? existing.baseAdditionalBullets ?? {};

    const updatedProfile = {
      ...existing,
      displayName: body.displayName ?? existing.displayName,
      baseInfo: mergedBase,
      baseResume,
      baseAdditionalBullets,
      resumeTemplateId,
      resumeTemplateName,
      updatedAt: new Date().toISOString(),
    };

    await updateProfileRecord({
      id: updatedProfile.id,
      displayName: updatedProfile.displayName,
      baseInfo: updatedProfile.baseInfo,
      baseResume: updatedProfile.baseResume,
      baseAdditionalBullets: updatedProfile.baseAdditionalBullets,
      resumeTemplateId: updatedProfile.resumeTemplateId,
    });
    return updatedProfile;
  });

  app.delete("/profiles/:id", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can delete profiles" });
    }
    const { id } = request.params as { id: string };
    const deleted = await deleteProfile(id);
    if (!deleted) {
      return reply.status(404).send({ message: "Profile not found" });
    }
    return { success: true };
  });

  app.get("/resume-templates", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (
      !actor ||
      (actor.role !== "MANAGER" &&
        actor.role !== "ADMIN" &&
        actor.role !== "BIDDER")
    ) {
      return reply
        .status(403)
        .send({ message: "Only managers, admins, or bidders can view templates" });
    }
    return listResumeTemplates();
  });

  app.post("/resume-templates", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can create templates" });
    }
    const schema = z.object({
      name: z.string(),
      description: z.string().optional().nullable(),
      html: z.string(),
    });
    const body = schema.parse(request.body ?? {});
    const name = trimString(body.name);
    const html = trimString(body.html);
    if (!name) {
      return reply.status(400).send({ message: "Template name is required" });
    }
    if (!html) {
      return reply.status(400).send({ message: "Template HTML is required" });
    }
    const now = new Date().toISOString();
    const created = await insertResumeTemplate({
      id: randomUUID(),
      name,
      description: trimToNull(body.description ?? null),
      html,
      createdBy: actor.id,
      createdAt: now,
      updatedAt: now,
    });
    try {
      await notifyAllUsers({
        kind: "system",
        message: `Resume template ${created.name} created.`,
        href: "/manager/resume-templates",
      });
    } catch (err) {
      request.log.error({ err }, "resume template create notification failed");
    }
    return created;
  });

  app.patch("/resume-templates/:id", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can update templates" });
    }
    const { id } = request.params as { id: string };
    const existing = await findResumeTemplateById(id);
    if (!existing) {
      return reply.status(404).send({ message: "Template not found" });
    }
    const schema = z.object({
      name: z.string().optional(),
      description: z.string().optional().nullable(),
      html: z.string().optional(),
    });
    const body = schema.parse(request.body ?? {});
    const name =
      body.name !== undefined ? trimString(body.name) : existing.name;
    const html = body.html !== undefined ? trimString(body.html) : existing.html;
    const description =
      body.description !== undefined
        ? trimToNull(body.description ?? null)
        : existing.description ?? null;
    if (!name) {
      return reply.status(400).send({ message: "Template name is required" });
    }
    if (!html) {
      return reply.status(400).send({ message: "Template HTML is required" });
    }
    const updated = await updateResumeTemplate({
      id,
      name,
      description,
      html,
    });
    if (updated) {
      try {
        await notifyAllUsers({
          kind: "system",
          message: `Resume template ${updated.name} updated.`,
          href: "/manager/resume-templates",
        });
      } catch (err) {
        request.log.error({ err }, "resume template update notification failed");
      }
    }
    return updated;
  });

  app.delete("/resume-templates/:id", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can delete templates" });
    }
    const { id } = request.params as { id: string };
    const deleted = await deleteResumeTemplate(id);
    if (!deleted) {
      return reply.status(404).send({ message: "Template not found" });
    }
    return { success: true };
  });

  app.post("/resume-templates/render-pdf", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (
      !actor ||
      (actor.role !== "MANAGER" &&
        actor.role !== "ADMIN" &&
        actor.role !== "BIDDER")
    ) {
      return reply
        .status(403)
        .send({ message: "Only managers, admins, or bidders can export templates" });
    }
    const schema = z.object({
      html: z.string(),
      filename: z.string().optional(),
    });
    const body = schema.parse(request.body ?? {});
    const html = trimString(body.html);
    if (!html) {
      return reply.status(400).send({ message: "HTML is required" });
    }
    if (html.length > 2_000_000) {
      return reply.status(413).send({ message: "Template too large" });
    }
    const fileName = buildSafePdfFilename(body.filename);
    let browser: Browser | undefined;
    try {
      browser = await chromium.launch({ headless: true });
      const page = await browser.newPage({ viewport: { width: 1240, height: 1754 } });
      await page.setContent(html, { waitUntil: "domcontentloaded" });
      await page.emulateMedia({ media: "screen" });
      const size = await page.evaluate(() => {
        const body = document.body;
        const doc = document.documentElement;
        const candidates = body ? Array.from(body.children) : [];
        let target: Element = body || doc;
        let bestArea = 0;
        for (const el of candidates as Element[]) {
          const rect = el.getBoundingClientRect();
          const area = rect.width * rect.height;
          if (area > bestArea) {
            bestArea = area;
            target = el;
          }
        }
        const targetEl = target as HTMLElement;
        if (targetEl?.style) {
          targetEl.style.margin = "0";
        }
        if (doc?.style) {
          doc.style.margin = "0";
          doc.style.padding = "0";
        }
        if (body?.style) {
          body.style.margin = "0";
          body.style.padding = "0";
        }
        const rect = target.getBoundingClientRect();
        const width = Math.max(1, Math.ceil(rect.width));
        const height = Math.max(1, Math.ceil(rect.height));
        if (body?.style) {
          body.style.width = `${width}px`;
          body.style.height = `${height}px`;
          body.style.overflow = "hidden";
        }
        if (doc?.style) {
          doc.style.width = `${width}px`;
          doc.style.height = `${height}px`;
          doc.style.overflow = "hidden";
        }
        return { width, height };
      });
      const pdfWidth = Math.max(1, Math.ceil(size.width));
      const pdfHeight = Math.max(1, Math.ceil(size.height));
      const pdf = await page.pdf({
        width: `${pdfWidth}px`,
        height: `${pdfHeight}px`,
        printBackground: true,
        margin: { top: "0mm", bottom: "0mm", left: "0mm", right: "0mm" },
      });
      reply.header("Content-Type", "application/pdf");
      reply.header("Content-Disposition", `attachment; filename="${fileName}"`);
      return reply.send(pdf);
    } catch (err) {
      request.log.error({ err }, "resume template pdf export failed");
      return reply.status(500).send({ message: "Unable to export PDF" });
    } finally {
      if (browser) {
        await browser.close().catch(() => undefined);
      }
    }
  });

  app.get("/tasks", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    if (actor.role === "OBSERVER") {
      return reply.status(403).send({ message: "Observers cannot view tasks" });
    }
    return listTasks();
  });

  app.post("/tasks", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can create tasks" });
    }
    const isAdmin = actor.role === "ADMIN";
    const schema = z.object({
      title: z.string(),
      detail: z.string().optional().nullable(),
      status: z.enum(["todo", "in_progress", "in_review", "done"]),
      priority: z.enum(["low", "medium", "high", "urgent"]),
      dueDate: z.string().optional().nullable(),
      project: z.string().optional().nullable(),
      notes: z.string().optional().nullable(),
      tags: z.array(z.string()).optional(),
      assigneeIds: z.array(z.string().uuid()).optional(),
    });
    const body = schema.parse(request.body ?? {});
    const title = trimString(body.title);
    if (!title) {
      return reply.status(400).send({ message: "Task title is required" });
    }
    const dueDate = trimToNull(body.dueDate ?? null);
    if (dueDate && !isValidDateString(dueDate)) {
      return reply.status(400).send({ message: "Invalid due date" });
    }
    const tags = (body.tags ?? []).map(trimString).filter(Boolean);
    const assigneeIds = body.assigneeIds ?? [];
    const status =
      body.status === "todo" && assigneeIds.length > 0 ? "in_progress" : body.status;
    const created = await insertTask({
      id: randomUUID(),
      title,
      detail: trimToNull(body.detail ?? null),
      status,
      priority: body.priority,
      dueDate,
      project: trimToNull(body.project ?? null),
      notes: trimToNull(body.notes ?? null),
      tags,
      createdBy: actor.id,
      assigneeIds: isAdmin ? assigneeIds : [],
    });
    if (!isAdmin && assigneeIds.length) {
      await upsertTaskAssignmentRequests(created.id, assigneeIds, actor.id);
    }
    if (isAdmin && assigneeIds.length) {
      await notifyUsers(assigneeIds, {
        kind: "system",
        message: `You were assigned to "${created.title}".`,
        href: "/tasks",
      });
    }
    if (!isAdmin) {
      await notifyAdmins({
        kind: "system",
        message: `Task request from ${actor.name}: ${created.title}`,
        href: "/admin/tasks",
      });
    }
    return created;
  });

  app.patch("/tasks/:id", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can update tasks" });
    }
    const { id } = request.params as { id: string };
    const existing = await findTaskById(id);
    if (!existing) {
      return reply.status(404).send({ message: "Task not found" });
    }
    const schema = z.object({
      title: z.string().optional(),
      detail: z.string().optional().nullable(),
      status: z.enum(["todo", "in_progress", "in_review", "done"]).optional(),
      priority: z.enum(["low", "medium", "high", "urgent"]).optional(),
      dueDate: z.string().optional().nullable(),
      project: z.string().optional().nullable(),
      notes: z.string().optional().nullable(),
      tags: z.array(z.string()).optional(),
      assigneeIds: z.array(z.string().uuid()).optional(),
    });
    const body = schema.parse(request.body ?? {});
    if (body.assigneeIds !== undefined && actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can update task assignees" });
    }
    const dueDate =
      body.dueDate !== undefined
        ? trimToNull(body.dueDate ?? null)
        : existing.dueDate;
    if (body.dueDate !== undefined && dueDate && !isValidDateString(dueDate)) {
      return reply.status(400).send({ message: "Invalid due date" });
    }
    const title =
      body.title !== undefined ? trimString(body.title) : existing.title;
    if (!title) {
      return reply.status(400).send({ message: "Task title is required" });
    }
    const tags =
      body.tags !== undefined
        ? body.tags.map(trimString).filter(Boolean)
        : existing.tags ?? [];
    const assigneeIds =
      body.assigneeIds !== undefined
        ? body.assigneeIds
        : existing.assignees.map((assignee) => assignee.id);
    const status =
      body.status ??
      (body.assigneeIds !== undefined &&
      existing.status === "todo" &&
      assigneeIds.length > 0
        ? "in_progress"
        : existing.status);
    const updated = await updateTask({
      id,
      title,
      detail:
        body.detail !== undefined
          ? trimToNull(body.detail ?? null)
          : existing.detail ?? null,
      status,
      priority: body.priority ?? existing.priority,
      dueDate,
      project:
        body.project !== undefined
          ? trimToNull(body.project ?? null)
          : existing.project ?? null,
      notes:
        body.notes !== undefined
          ? trimToNull(body.notes ?? null)
          : existing.notes ?? null,
      tags,
      assigneeIds,
    });
    return updated;
  });

  app.patch("/tasks/:id/notes", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const { id } = request.params as { id: string };
    const existing = await findTaskById(id);
    if (!existing) {
      return reply.status(404).send({ message: "Task not found" });
    }
    const isAssignee = existing.assignees.some((assignee) => assignee.id === actor.id);
    if (!isAssignee) {
      return reply
        .status(403)
        .send({ message: "Only assigned users can add notes" });
    }
    const schema = z.object({
      note: z.string().optional().nullable(),
      notes: z.string().optional().nullable(),
    });
    const body = schema.parse(request.body ?? {});
    const incoming = trimString(body.note ?? body.notes ?? "");
    if (!incoming) {
      return reply.status(400).send({ message: "Note is required" });
    }
    const sanitized = incoming.replace(/\s+/g, " ").trim();
    if (!sanitized) {
      return reply.status(400).send({ message: "Note is required" });
    }
    const stamp = new Date().toLocaleString("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric",
      hour: "numeric",
      minute: "2-digit",
    });
    const entry = `- ${stamp} - ${actor.name}: ${sanitized}`;
    const existingNotes = existing.notes?.trim();
    const nextNotes = existingNotes ? `${existingNotes}\n${entry}` : entry;
    const updated = await updateTaskNotes(id, nextNotes);
    return updated;
  });

  app.post("/tasks/:id/reject", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can reject tasks" });
    }
    const { id } = request.params as { id: string };
    const existing = await findTaskById(id);
    if (!existing) {
      return reply.status(404).send({ message: "Task not found" });
    }
    const schema = z.object({
      reason: z.string().optional().nullable(),
    });
    const body = schema.parse(request.body ?? {});
    const updated = await rejectTask(id, actor.id, trimToNull(body.reason ?? null));
    if (updated?.createdBy) {
      const reasonText = updated.rejectionReason
        ? ` Reason: ${updated.rejectionReason}`
        : "";
      await notifyUsers([updated.createdBy], {
        kind: "system",
        message: `Task rejected: ${updated.title}.${reasonText}`,
        href: "/tasks",
      });
    }
    return updated;
  });

  app.post("/tasks/:id/assign-requests", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.role !== "MANAGER") {
      return reply
        .status(403)
        .send({ message: "Only managers can request assignments" });
    }
    const { id } = request.params as { id: string };
    const existing = await findTaskById(id);
    if (!existing) {
      return reply.status(404).send({ message: "Task not found" });
    }
    if (existing.createdBy !== actor.id) {
      return reply
        .status(403)
        .send({ message: "You can only request assignments for your tasks" });
    }
    if (existing.rejectedAt) {
      return reply
        .status(400)
        .send({ message: "Rejected tasks cannot accept assignments" });
    }
    const schema = z.object({
      assigneeIds: z.array(z.string().uuid()).min(1),
    });
    const body = schema.parse(request.body ?? {});
    const assignedIds = new Set(existing.assignees.map((assignee) => assignee.id));
    const requestIds = body.assigneeIds.filter((assigneeId) => !assignedIds.has(assigneeId));
    if (!requestIds.length) {
      return reply.status(400).send({ message: "All users are already assigned" });
    }
    await upsertTaskAssignmentRequests(id, requestIds, actor.id);
    await notifyAdmins({
      kind: "system",
      message: `Assignment request for "${existing.title}" (${requestIds.length} user${requestIds.length === 1 ? "" : "s"})`,
      href: "/admin/tasks",
    });
    return { success: true };
  });

  app.get("/tasks/assign-requests", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can view assignment requests" });
    }
    const parsed = z
      .object({
        status: z.enum(["pending", "approved", "rejected"]).optional(),
      })
      .safeParse(request.query);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }
    return listTaskAssignmentRequests(parsed.data.status ?? "pending");
  });

  app.post("/tasks/:id/done-requests", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const { id } = request.params as { id: string };
    const task = await findTaskById(id);
    if (!task) {
      return reply.status(404).send({ message: "Task not found" });
    }
    const isAssignee = task.assignees.some((assignee) => assignee.id === actor.id);
    if (!isAssignee) {
      return reply
        .status(403)
        .send({ message: "Only assigned users can request completion" });
    }
    if (task.status === "done") {
      return reply.status(400).send({ message: "Task already done" });
    }
    const { rows } = await pool.query<{ id: string }>(
      `
        SELECT id FROM task_done_requests
        WHERE task_id = $1 AND requested_by = $2 AND status = 'pending'
        LIMIT 1
      `,
      [id, actor.id],
    );
    if (rows.length > 0) {
      return reply.status(409).send({ message: "Completion already requested" });
    }
    const created = await insertTaskDoneRequest({
      id: randomUUID(),
      taskId: id,
      requestedBy: actor.id,
    });
    await updateTaskStatus(id, "in_review");
    await notifyAdmins({
      kind: "system",
      message: `Task done request from ${actor.name}: ${task.title}`,
      href: "/admin/tasks",
    });
    return created;
  });

  app.get("/tasks/done-requests", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can view done requests" });
    }
    const schema = z.object({
      status: z.enum(["pending", "approved", "rejected"]).optional(),
    });
    const parsed = schema.safeParse(request.query ?? {});
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }
    return listTaskDoneRequests(parsed.data.status ?? "pending");
  });

  app.post("/tasks/done-requests/:id/approve", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can approve done requests" });
    }
    const { id } = request.params as { id: string };
    const approved = await approveTaskDoneRequest(id, actor.id);
    if (!approved) {
      return reply.status(404).send({ message: "Done request not found" });
    }
    await notifyUsers([approved.requestedBy], {
      kind: "system",
      message: `Task marked done: ${approved.taskTitle}`,
      href: "/tasks",
    });
    return approved;
  });

  app.post("/tasks/done-requests/:id/reject", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can reject done requests" });
    }
    const { id } = request.params as { id: string };
    const schema = z.object({
      reason: z.string().optional().nullable(),
    });
    const body = schema.parse(request.body ?? {});
    const rejected = await rejectTaskDoneRequest(id, actor.id, body.reason ?? null);
    if (!rejected) {
      return reply.status(404).send({ message: "Done request not found" });
    }
    await notifyUsers([rejected.requestedBy], {
      kind: "system",
      message: `Task done request rejected: ${rejected.taskTitle}`,
      href: "/tasks",
    });
    return rejected;
  });

  app.post("/tasks/assign-requests/:id/approve", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can approve assignment requests" });
    }
    const { id } = request.params as { id: string };
    const approved = await approveTaskAssignmentRequest(id, actor.id);
    if (!approved) {
      return reply.status(404).send({ message: "Assignment request not found" });
    }
    const task = await findTaskById(approved.taskId);
    const title = task?.title ?? "task";
    if (approved.requestedBy) {
      await notifyUsers([approved.requestedBy], {
        kind: "system",
        message: `Assignment approved for "${title}".`,
        href: "/tasks",
      });
    }
    await notifyUsers([approved.userId], {
      kind: "system",
      message: `You were assigned to "${title}".`,
      href: "/tasks",
    });
    return { success: true };
  });

  app.post("/tasks/assign-requests/:id/reject", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can reject assignment requests" });
    }
    const { id } = request.params as { id: string };
    const schema = z.object({
      reason: z.string().optional().nullable(),
    });
    const body = schema.parse(request.body ?? {});
    const reason = trimToNull(body.reason ?? null);
    const rejected = await rejectTaskAssignmentRequest(
      id,
      actor.id,
      reason,
    );
    if (!rejected) {
      return reply.status(404).send({ message: "Assignment request not found" });
    }
    const task = await findTaskById(rejected.taskId);
    const title = task?.title ?? "task";
    if (rejected.requestedBy) {
      const reasonText = reason ? ` Reason: ${reason}` : "";
      await notifyUsers([rejected.requestedBy], {
        kind: "system",
        message: `Assignment rejected for "${title}".${reasonText}`,
        href: "/tasks",
      });
    }
    return { success: true };
  });

  app.post("/tasks/:id/assign-self-request", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const { id } = request.params as { id: string };
    const task = await findTaskById(id);
    if (!task) {
      return reply.status(404).send({ message: "Task not found" });
    }
    if (task.status !== "todo") {
      return reply.status(400).send({ message: "Assignment requests are only for to-do tasks" });
    }
    const isAssigned = task.assignees.some((assignee) => assignee.id === actor.id);
    if (isAssigned) {
      return reply.status(409).send({ message: "You are already assigned" });
    }
    const { rows } = await pool.query<{ id: string }>(
      `
        SELECT id FROM task_assignment_requests
        WHERE task_id = $1 AND user_id = $2 AND status = 'pending'
        LIMIT 1
      `,
      [id, actor.id],
    );
    if (rows.length > 0) {
      return reply.status(409).send({ message: "Assignment already requested" });
    }
    await upsertTaskAssignmentRequests(id, [actor.id], actor.id);
    await notifyAdmins({
      kind: "system",
      message: `Assignment request from ${actor.name}: ${task.title}`,
      href: "/admin/tasks",
    });
    return { success: true };
  });

  app.post("/tasks/:id/assign-self", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const { id } = request.params as { id: string };
    const task = await findTaskById(id);
    if (!task) {
      return reply.status(404).send({ message: "Task not found" });
    }
    if (task.status !== "in_progress") {
      return reply.status(400).send({ message: "Assign-to-me is only available in progress" });
    }
    const isAssigned = task.assignees.some((assignee) => assignee.id === actor.id);
    if (isAssigned) {
      return reply.status(409).send({ message: "You are already assigned" });
    }
    await addTaskAssignee(id, actor.id);
    return findTaskById(id);
  });

  app.delete("/tasks/:id", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can delete tasks" });
    }
    const { id } = request.params as { id: string };
    const deleted = await deleteTask(id);
    if (!deleted) {
      return reply.status(404).send({ message: "Task not found" });
    }
    return { success: true };
  });

  app.get("/calendar/accounts", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const parsed = z
      .object({
        profileId: z.string().uuid().optional(),
      })
      .safeParse(request.query);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }
    return listProfileAccountsForUser(actor, parsed.data.profileId);
  });

  app.post("/calendar/accounts", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const schema = z.object({
      profileId: z.string().uuid(),
      email: z.string().email(),
      provider: z.enum(["MICROSOFT", "GOOGLE"]).default("MICROSOFT").optional(),
      displayName: z.string().min(1).optional(),
      timezone: z.string().min(2).optional(),
    });
    const body = schema.parse(request.body ?? {});
    const profile = await findProfileById(body.profileId);
    if (!profile) {
      return reply.status(404).send({ message: "Profile not found" });
    }
    const isManager = actor.role === "ADMIN" || actor.role === "MANAGER";
    const isAssignedBidder = profile.assignedBidderId === actor.id;
    if (!isManager && !isAssignedBidder) {
      return reply
        .status(403)
        .send({ message: "Not allowed to manage accounts for this profile" });
    }
    const account = await upsertProfileAccount({
      id: randomUUID(),
      profileId: body.profileId,
      provider: body.provider ?? "MICROSOFT",
      email: body.email.toLowerCase(),
      displayName: body.displayName ?? body.email,
      timezone: body.timezone ?? "UTC",
      status: "ACTIVE",
    });
    return account;
  });

  app.post("/calendar/oauth/accounts", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const schema = z.object({
      providerAccountId: z.string().min(1),
      email: z.string().email(),
      displayName: z.string().optional(),
      accessToken: z.string().min(1),
      refreshToken: z.string().optional(),
      expiresAt: z.number().int().optional(),
      idToken: z.string().optional(),
      scope: z.string().optional(),
    });
    const body = schema.parse(request.body ?? {});
    const account = await upsertUserOAuthAccount({
      id: randomUUID(),
      userId: actor.id,
      provider: "azure-ad",
      providerAccountId: body.providerAccountId,
      email: body.email.toLowerCase(),
      displayName: body.displayName ?? null,
      accessToken: body.accessToken,
      refreshToken: body.refreshToken ?? null,
      expiresAt: body.expiresAt ?? null,
      idToken: body.idToken ?? null,
      scope: body.scope ?? null,
    });
    return {
      id: account.id,
      email: account.email,
      displayName: account.displayName,
      accountId: account.id,
    };
  });

  app.get("/calendar/oauth/accounts", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const parsed = z
      .object({
        provider: z.string().optional(),
      })
      .safeParse(request.query);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }
    const accounts = await listUserOAuthAccounts(actor.id, parsed.data.provider);
    return accounts.map((account) => ({
      id: account.id,
      email: account.email,
      displayName: account.displayName,
      accountId: account.id,
      providerAccountId: account.providerAccountId,
    }));
  });

  app.delete("/calendar/oauth/accounts/:id", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const { id } = request.params as { id: string };
    const deleted = await deleteUserOAuthAccount(actor.id, id);
    if (!deleted) {
      return reply.status(404).send({ message: "Account not found" });
    }
    return { ok: true };
  });

  app.post("/calendar/oauth/refresh", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const schema = z.object({
      accountId: z.string().uuid(),
    });
    const body = schema.parse(request.body ?? {});
    const account = await findUserOAuthAccountById(body.accountId);
    if (!account || account.userId !== actor.id) {
      return reply.status(404).send({ message: "Account not found" });
    }
    // Token refresh will be handled by the calendar/outlook endpoint
    // This endpoint can be used for manual refresh if needed
    return {
      id: account.id,
      email: account.email,
      displayName: account.displayName,
      accountId: account.id,
    };
  });

  // OAuth Authorization endpoint - returns the OAuth URL
  app.get("/calendar/oauth/authorize", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }

    if (!config.MS_CLIENT_ID) {
      return reply.status(500).send({ message: "MS_CLIENT_ID is not configured" });
    }

    const parsed = z
      .object({
        redirect_uri: z.string().url(),
        frontend_redirect: z.string().url().optional(),
      })
      .safeParse(request.query);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid redirect_uri" });
    }

    const tenantId = config.MS_TENANT_ID || "common";
    const baseScope = "openid profile email offline_access Calendars.Read User.Read";
    const includeSharedCalendars =
      process.env.MS_GRAPH_SHARED_CALENDARS === "true" ||
      (tenantId !== "common" && tenantId !== "consumers");
    const scope = includeSharedCalendars
      ? `${baseScope} Calendars.Read.Shared`
      : baseScope;

    // Generate random state for CSRF protection and encode frontend_redirect in it
    const stateData = {
      random: randomUUID(),
      frontend_redirect: parsed.data.frontend_redirect,
    };
    const state = Buffer.from(JSON.stringify(stateData)).toString("base64url");

    // Strip query parameters from redirect_uri (Azure AD doesn't accept them)
    const cleanRedirectUri = new URL(parsed.data.redirect_uri);
    cleanRedirectUri.search = "";

    const params = new URLSearchParams({
      client_id: config.MS_CLIENT_ID,
      response_type: "code",
      redirect_uri: cleanRedirectUri.toString(),
      response_mode: "query",
      scope,
      state,
    });

    const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${params.toString()}`;

    return { authUrl, state };
  });

  // OAuth Callback endpoint - handles the callback and stores tokens
  app.get("/calendar/oauth/callback", async (request, reply) => {
    const parsed = z
      .object({
        code: z.string().optional(),
        error: z.string().optional(),
        error_description: z.string().optional(),
        state: z.string().optional(),
        frontend_redirect: z.string().url().optional(),
      })
      .safeParse(request.query);

    if (!parsed.success) {
      const frontendUrl = Array.isArray(config.CORS_ORIGINS)
        ? config.CORS_ORIGINS[0]
        : typeof config.CORS_ORIGINS === "string"
        ? config.CORS_ORIGINS
        : "http://localhost:3000";
      return reply.redirect(
        `${frontendUrl}/calendar?error=${encodeURIComponent("Invalid query parameters")}`
      );
    }

    const { code, error, error_description, state } = parsed.data;

    // Decode frontend_redirect from state parameter
    let frontendRedirect: string | undefined;
    if (state) {
      try {
        const stateData = JSON.parse(Buffer.from(state, "base64url").toString());
        frontendRedirect = stateData.frontend_redirect;
      } catch {
        // If state decoding fails, continue without frontend_redirect
      }
    }

    // Determine frontend URL
    const frontendUrl =
      frontendRedirect ||
      (Array.isArray(config.CORS_ORIGINS)
        ? config.CORS_ORIGINS[0]
        : typeof config.CORS_ORIGINS === "string"
        ? config.CORS_ORIGINS
        : "http://localhost:3000");

    // Handle errors from Microsoft
    if (error) {
      return reply.redirect(
        `${frontendUrl}/calendar?error=${encodeURIComponent(error_description || error)}`
      );
    }

    if (!code) {
      return reply.redirect(`${frontendUrl}/calendar?error=missing_code`);
    }

    // Get the redirect URI that was used (without query parameters - Azure AD doesn't allow them)
    const protocol = request.headers['x-forwarded-proto'] || (request.socket && 'encrypted' in request.socket && (request.socket as any).encrypted ? 'https' : 'http') || 'http';
    const host = request.headers.host || `localhost:${config.PORT}`;
    const callbackRedirectUri = `${protocol}://${host}/calendar/oauth/callback`;

    // Exchange code for tokens
    const tenantId = config.MS_TENANT_ID || "common";
    const baseScope = "openid profile email offline_access Calendars.Read User.Read";
    const includeSharedCalendars =
      process.env.MS_GRAPH_SHARED_CALENDARS === "true" ||
      (tenantId !== "common" && tenantId !== "consumers");
    const scope = includeSharedCalendars
      ? `${baseScope} Calendars.Read.Shared`
      : baseScope;

    const tokenParams = new URLSearchParams({
      client_id: config.MS_CLIENT_ID,
      client_secret: config.MS_CLIENT_SECRET,
      code,
      redirect_uri: callbackRedirectUri,
      grant_type: "authorization_code",
      scope,
    });

    let tokenData: {
      access_token?: string;
      refresh_token?: string;
      expires_in?: number;
      id_token?: string;
    };

    try {
      const tokenRes = await fetch(
        `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: tokenParams,
        }
      );

      if (!tokenRes.ok) {
        const errorData = (await tokenRes.json().catch(() => ({}))) as { error_description?: string };
        throw new Error(
          errorData.error_description || "Failed to exchange authorization code"
        );
      }

      tokenData = await tokenRes.json() as { access_token?: string; refresh_token?: string; expires_in?: number; id_token?: string };
    } catch (err) {
      const message =
        err instanceof Error ? err.message : "Failed to exchange authorization code";
      return reply.redirect(
        `${frontendUrl}/calendar?error=${encodeURIComponent(message)}`
      );
    }

    if (!tokenData.access_token) {
      return reply.redirect(
        `${frontendUrl}/calendar?error=${encodeURIComponent("No access token received")}`
      );
    }

    // Get user profile
    let email = "";
    let displayName: string | undefined;
    let providerAccountId = "";

    // Try to get email from ID token first (most reliable)
    if (tokenData.id_token) {
      try {
        const idTokenParts = tokenData.id_token.split('.');
        if (idTokenParts.length === 3) {
          const payload = JSON.parse(
            Buffer.from(idTokenParts[1], 'base64url').toString()
          ) as { email?: string; preferred_username?: string };
          if (payload.email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(payload.email)) {
            email = payload.email;
          } else if (payload.preferred_username && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(payload.preferred_username)) {
            email = payload.preferred_username;
          }
        }
      } catch {
        // Continue to try profile API
      }
    }

    // If no valid email from ID token, try Graph API
    if (!email) {
      try {
        const profileRes = await fetch("https://graph.microsoft.com/v1.0/me", {
          headers: { Authorization: `Bearer ${tokenData.access_token}` },
        });

        if (profileRes.ok) {
          const profile = (await profileRes.json()) as {
            mail?: string;
            userPrincipalName?: string;
            displayName?: string;
            id?: string;
          };
          
          // Prefer mail field (actual email), fallback to userPrincipalName
          if (profile.mail && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(profile.mail)) {
            email = profile.mail;
          } else if (profile.userPrincipalName) {
            // Clean userPrincipalName - remove #EXT# patterns for guest users
            let cleaned = profile.userPrincipalName;
            // Remove #EXT#@tenant.onmicrosoft.com pattern
            cleaned = cleaned.replace(/#EXT#@[^@]+$/, '');
            // If it still looks like an email, use it
            if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleaned)) {
              email = cleaned;
            }
          }
          
          displayName = profile.displayName;
          providerAccountId = profile.id || email;
        }
      } catch (err) {
        // Continue even if profile fetch fails
      }
    }

    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return reply.redirect(
        `${frontendUrl}/calendar?error=${encodeURIComponent("Failed to get valid email from OAuth provider")}`
      );
    }

    const expiresAt = tokenData.expires_in
      ? Math.floor(Date.now() / 1000) + tokenData.expires_in
      : undefined;

    // Get user from Authorization header if present
    const authHeader = request.headers.authorization;
    let userId: string | null = null;

    if (authHeader && authHeader.startsWith("Bearer ")) {
      try {
        const token = authHeader.substring(7);
        const decoded = verifyToken(token);
        userId = decoded.sub;
      } catch {
        // Token invalid, will redirect to frontend with tokens
      }
    }

    // If no user from token, redirect to frontend with tokens in query
    // Frontend will send them to backend with proper auth
    if (!userId) {
      const tokenParams = new URLSearchParams({
        access_token: tokenData.access_token,
        ...(tokenData.refresh_token && { refresh_token: tokenData.refresh_token }),
        ...(expiresAt && { expires_at: expiresAt.toString() }),
        ...(tokenData.id_token && { id_token: tokenData.id_token }),
        email,
        ...(displayName && { display_name: displayName }),
        provider_account_id: providerAccountId,
      });
      return reply.redirect(
        `${frontendUrl}/calendar/oauth/callback?${tokenParams.toString()}`
      );
    }

    // Store account if we have userId
    const account = await upsertUserOAuthAccount({
      id: randomUUID(),
      userId,
      provider: "azure-ad",
      providerAccountId,
      email: email.toLowerCase(),
      displayName: displayName ?? null,
      accessToken: tokenData.access_token,
      refreshToken: tokenData.refresh_token ?? null,
      expiresAt: expiresAt ?? null,
      idToken: tokenData.id_token ?? null,
      scope: scope ?? null,
    });

    return reply.redirect(`${frontendUrl}/calendar?success=connected`);
  });

    app.post('/calendar/events/sync', async (request, reply) => {
      if (forbidObserver(reply, request.authUser)) return;
      const actor = request.authUser;
      if (!actor) {
        return reply.status(401).send({ message: 'Unauthorized' });
      }
    const schema = z.object({
      mailboxes: z.array(z.string().email()).default([]),
      timezone: z.string().min(2).optional(),
      events: z
        .array(
          z.object({
            id: z.string().min(1),
            title: z.string().optional(),
            start: z.string().min(1),
            end: z.string().min(1),
            isAllDay: z.boolean().optional(),
            organizer: z.string().optional(),
            location: z.string().optional(),
            mailbox: z.string().email(),
          }),
        )
        .default([]),
    });
    const body = schema.parse(request.body ?? {});
    const mailboxes = body.mailboxes.map((mailbox) => mailbox.toLowerCase());
    const events = body.events.map((event) => ({
      ...event,
      mailbox: event.mailbox.toLowerCase(),
    }));
      const storedEvents = await replaceCalendarEvents({
        ownerUserId: actor.id,
        mailboxes,
        timezone: body.timezone ?? null,
        events,
      });
      return { events: storedEvents };
    });

    app.get('/calendar/events/stored', async (request, reply) => {
      if (forbidObserver(reply, request.authUser)) return;
      const actor = request.authUser;
      if (!actor) {
        return reply.status(401).send({ message: 'Unauthorized' });
      }
      const parsed = z
        .object({
          start: z.string().optional(),
          end: z.string().optional(),
          mailboxes: z.string().optional(),
        })
        .safeParse(request.query);
      if (!parsed.success) {
        return reply.status(400).send({ message: 'Invalid query' });
      }
      const mailboxes = parsed.data.mailboxes
        ? parsed.data.mailboxes
            .split(',')
            .map((mailbox) => mailbox.trim().toLowerCase())
            .filter(Boolean)
        : [];
      let ownerUserId = actor.id;
      if (actor.role !== 'ADMIN') {
        const { rows } = await pool.query<{ id: string }>(
          "SELECT id FROM users WHERE role = 'ADMIN' ORDER BY created_at ASC LIMIT 1",
        );
        if (rows[0]?.id) {
          ownerUserId = rows[0].id;
        }
      }
      const events = await listCalendarEventsForOwner(ownerUserId, mailboxes, {
        start: parsed.data.start ?? null,
        end: parsed.data.end ?? null,
      });
      return { events };
    });

    app.get('/calendar/events', async (request, reply) => {
      if (forbidObserver(reply, request.authUser)) return;
      const actor = request.authUser;
      if (!actor) {
        return reply.status(401).send({ message: 'Unauthorized' });
    }
    const parsed = z
      .object({
        accountId: z.string().uuid(),
        start: z.string(),
        end: z.string(),
      })
      .safeParse(request.query);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }
    const { accountId, start, end } = parsed.data;
    const account = await findProfileAccountById(accountId);
    if (!account) {
      return reply.status(404).send({ message: "Calendar account not found" });
    }
    const isManager = actor.role === "ADMIN" || actor.role === "MANAGER";
    const isAssignedBidder = account.profileAssignedBidderId === actor.id;
    if (!isManager && !isAssignedBidder) {
      return reply
        .status(403)
        .send({ message: "Not allowed to view this calendar" });
    }
    const startDate = new Date(start);
    const endDate = new Date(end);
    if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) {
      return reply.status(400).send({ message: "Invalid date range" });
    }
    if (endDate <= startDate) {
      return reply.status(400).send({ message: "End must be after start" });
    }
    const {
      events: calendarEvents,
      source,
      warning,
    } = await loadOutlookEvents({
      email: account.email,
      rangeStart: start,
      rangeEnd: end,
      timezone: account.timezone,
      logger: request.log,
    });
    await touchProfileAccount(account.id, new Date().toISOString());
    return {
      account: {
        id: account.id,
        email: account.email,
        profileId: account.profileId,
        profileDisplayName: account.profileDisplayName,
        timezone: account.timezone,
      },
      events: calendarEvents,
      source,
      warning,
    };
  });

  app.get("/calendar/outlook", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const parsed = z
      .object({
        start: z.string(),
        end: z.string(),
        timezone: z.string().optional(),
        mailboxes: z.string().optional(),
        source: z.enum(["db", "graph"]).optional(),
      })
      .safeParse(request.query);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }

    const { start, end, timezone, mailboxes, source } = parsed.data;
    const preferStored = source !== "graph";
    const tz = timezone || "UTC";
    const mailboxParams = mailboxes
      ? mailboxes
          .split(",")
          .map((m) => m.trim().toLowerCase())
          .filter(Boolean)
      : [];

    // Non-admin users can ONLY load from database, never from Graph API
    if (source === "graph" && actor.role !== "ADMIN") {
      return reply.status(403).send({ message: "Only admins can sync from Graph API" });
    }

    // Helper function to load accounts for both admin and users
    const loadAccountsForUser = async (userId: string) => {
      return await listUserOAuthAccounts(userId, "azure-ad");
    };

    // Try to get cached events first if preferred
    if (preferStored) {
      let ownerUserId = actor.id;
      if (actor.role !== "ADMIN") {
        const { rows } = await pool.query<{ id: string }>(
          "SELECT id FROM users WHERE role = 'ADMIN' ORDER BY created_at ASC LIMIT 1",
        );
        if (rows[0]?.id) {
          ownerUserId = rows[0].id;
        }
      }

      // Convert mailbox email params to mailboxIds if provided
      let mailboxIds: string[] | undefined = undefined;
      if (mailboxParams.length > 0) {
        const accountMap = new Map<string, string>(); // email -> id
        const ownerAccounts = await loadAccountsForUser(ownerUserId);
        ownerAccounts.forEach((acc) => {
          accountMap.set(acc.email.toLowerCase(), acc.id);
        });
        
        // For non-admin users, also check admin's accounts
        if (actor.role !== "ADMIN" && ownerUserId !== actor.id) {
          const adminAccounts = await loadAccountsForUser(ownerUserId);
          adminAccounts.forEach((acc) => {
            accountMap.set(acc.email.toLowerCase(), acc.id);
          });
        }

        mailboxIds = mailboxParams
          .map((email) => accountMap.get(email))
          .filter((id): id is string => id !== undefined);
      }

      const cachedEvents = await listCalendarEventsForOwner(
        ownerUserId,
        mailboxIds,
        { start, end },
      );
      
      // Always load accounts from user_oauth_accounts
      const actorAccounts = await loadAccountsForUser(actor.id);
      let allAccounts = [...actorAccounts];
      
      // For non-admin users, also load admin's accounts
      if (actor.role !== "ADMIN" && ownerUserId !== actor.id) {
        const adminAccounts = await loadAccountsForUser(ownerUserId);
        allAccounts = [...allAccounts, ...adminAccounts];
      }

      // Build account map with IDs
      const accountMap = new Map<string, { id: string; email: string; name?: string; accountId: string; isPrimary?: boolean }>();
      allAccounts.forEach((acc) => {
        const key = acc.email.toLowerCase();
        if (!accountMap.has(key)) {
          accountMap.set(key, {
            id: acc.id,  // mailbox_id
            email: acc.email,
            name: acc.displayName ?? undefined,
            accountId: acc.id,
            isPrimary: false,
          });
        }
      });

      // Add accounts from cached events that might not be in user_oauth_accounts
      cachedEvents.forEach((ev) => {
        if (ev.mailbox) {
          const mailbox = ev.mailbox.toLowerCase();
          if (!accountMap.has(mailbox) && ev.mailboxId) {
            accountMap.set(mailbox, {
              id: ev.mailboxId,
              email: ev.mailbox,
              accountId: ev.mailboxId,
              isPrimary: false,
            });
          }
        }
      });

      // For non-admin users, always return database results (even if empty)
      // For admin users, only return if we have cached events
      if (cachedEvents.length > 0 || actor.role !== "ADMIN") {
        return {
          accounts: Array.from(accountMap.values()).sort((a, b) => a.email.localeCompare(b.email)),
          events: cachedEvents.map((ev) => ({
            ...ev,
            mailboxId: ev.mailboxId || undefined,
          })),
          source: "db",
          failedMailboxes: [],
        };
      }
    }

    // Fetch from Microsoft Graph
    const accounts = await listUserOAuthAccounts(actor.id, "azure-ad");
    if (accounts.length === 0) {
      return reply.status(400).send({ message: "No OAuth accounts connected" });
    }

    const accountResults = await Promise.all(
      accounts.map(async (account) => {
        try {
          const freshAccount = await ensureFreshToken(account);
          if (!freshAccount.accessToken) {
            return {
              status: "error" as const,
              mailbox: account.email.toLowerCase(),
              accountId: account.id,
              error: "Missing access token",
            };
          }

          // Fetch profile to get display name
          let displayName = account.displayName;
          try {
            const profileRes = await fetch(
              "https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName,displayName",
              {
                headers: { Authorization: `Bearer ${freshAccount.accessToken}` },
              },
            );
            if (profileRes.ok) {
              const profile = (await profileRes.json()) as {
                mail?: string;
                userPrincipalName?: string;
                displayName?: string;
              };
              displayName = profile.displayName ?? account.displayName ?? undefined;
            }
          } catch {
            // Ignore profile fetch errors
          }

          const { events, warning } = await loadOutlookEvents({
            email: account.email,
            rangeStart: start,
            rangeEnd: end,
            timezone: tz,
            account: freshAccount,
          });

          if (warning) {
            return {
              status: "error" as const,
              mailbox: account.email.toLowerCase(),
              accountId: account.id,
              error: warning,
              name: displayName,
            };
          }

          return {
            status: "success" as const,
            mailbox: account.email.toLowerCase(),
            accountId: account.id,
            mailboxId: account.id,  // Add mailboxId
            name: displayName,
            events: events.map((ev) => ({
              ...ev,
              mailbox: account.email.toLowerCase(),
              mailboxId: account.id,  // Add mailboxId to events
            })),
          };
        } catch (err) {
          return {
            status: "error" as const,
            mailbox: account.email.toLowerCase(),
            accountId: account.id,
            error: err instanceof Error ? err.message : "Failed to load events",
          };
        }
      }),
    );

    const successfulAccounts = accountResults.filter((r) => r.status === "success") as Array<{
      status: "success";
      mailbox: string;
      accountId: string;
      mailboxId: string;
      name?: string;
      events: Array<CalendarEvent & { mailbox: string; mailboxId: string }>;
    }>;

    const failedAccounts = accountResults.filter((r) => r.status === "error") as Array<{
      status: "error";
      mailbox: string;
      accountId: string;
      error: string;
      name?: string;
    }>;

    // Handle shared mailboxes
    const connectedMailboxSet = new Set(successfulAccounts.map((r) => r.mailbox));
    const sharedMailboxes = mailboxParams.filter((m) => !connectedMailboxSet.has(m));
    const sharedResults: Array<{ mailbox: string; mailboxId?: string; events?: CalendarEvent[]; error?: string }> = [];

    if (sharedMailboxes.length > 0 && accounts.length > 0) {
      // Use first account's token for shared mailbox access
      const primaryAccount = accounts[0];
      const freshAccount = await ensureFreshToken(primaryAccount);
      if (freshAccount.accessToken) {
        for (const mailbox of sharedMailboxes) {
          try {
            const { events, warning } = await loadOutlookEvents({
              email: mailbox,
              rangeStart: start,
              rangeEnd: end,
              timezone: tz,
              account: freshAccount,
            });
            if (warning) {
              sharedResults.push({ mailbox, error: warning });
            } else {
              // Try to find mailbox_id for shared mailbox
              const sharedMailboxId = await getMailboxIdFromEmail(actor.id, mailbox, actor.role === "ADMIN");
              sharedResults.push({
                mailbox,
                mailboxId: sharedMailboxId || undefined,
                events: events.map((ev) => ({
                  ...ev,
                  mailbox,
                  mailboxId: sharedMailboxId || undefined,
                })),
              });
            }
          } catch (err) {
            sharedResults.push({
              mailbox,
              error: err instanceof Error ? err.message : "Failed to load events",
            });
          }
        }
      }
    }

    const successfulShared = sharedResults.filter((r) => r.events) as Array<{
      mailbox: string;
      mailboxId?: string;
      events: Array<CalendarEvent & { mailbox: string; mailboxId?: string }>;
    }>;
    const failedShared = sharedResults.filter((r) => r.error) as Array<{
      mailbox: string;
      error: string;
    }>;

    const allEvents = [
      ...successfulAccounts.flatMap((r) => r.events),
      ...successfulShared.flatMap((r) => r.events),
    ];

    const accountsList = [
      ...successfulAccounts.map((r) => ({
        id: r.mailboxId,  // Add id (mailbox_id)
        email: r.mailbox,
        name: r.name,
        accountId: r.accountId,
        isPrimary: false,
        timezone: tz,
      })),
      ...successfulShared.map((r) => ({
        id: r.mailboxId || "",  // Add id (mailbox_id)
        email: r.mailbox,
        accountId: r.mailboxId || "",
        isPrimary: false,
        timezone: tz,
      })),
    ].sort((a, b) => a.email.localeCompare(b.email));

    const failedMailboxes = [
      ...failedAccounts.map((r) => r.mailbox),
      ...failedShared.map((r) => r.mailbox),
    ];

    // When admin syncs from Graph API, save events to database
    console.log('[CALENDAR SYNC] Checking save conditions:', {
      source,
      isAdmin: actor.role === "ADMIN",
      eventCount: allEvents.length,
      conditionMet: source === "graph" && actor.role === "ADMIN" && allEvents.length > 0
    });

    if (source === "graph" && actor.role === "ADMIN" && allEvents.length > 0) {
      try {
        console.log('[CALENDAR SYNC] Starting save process...');
        // Debug: Log event structure to verify mailboxId is set
        const sampleEvent = allEvents[0] ? {
          id: allEvents[0].id,
          title: allEvents[0].title,
          mailboxId: (allEvents[0] as any).mailboxId,
          mailbox: (allEvents[0] as any).mailbox,
        } : null;
        console.log('[CALENDAR SYNC] Sample event:', sampleEvent);
        request.log.debug({ 
          totalEvents: allEvents.length,
          sampleEvent
        }, "Events before filtering for save");

        // Filter events that have mailboxId and map them for saving
        const eventsToSave = allEvents
          .map((ev) => {
            const mailboxId = (ev as any).mailboxId;
            if (!mailboxId) {
              request.log.warn({ eventId: ev.id, eventTitle: ev.title }, "Event missing mailboxId, skipping save");
              return null;
            }
            return {
              id: ev.id,
              mailboxId: mailboxId,
              title: ev.title,
              start: ev.start,
              end: ev.end,
              isAllDay: ev.isAllDay,
              organizer: ev.organizer,
              location: ev.location,
            };
          })
          .filter((ev): ev is NonNullable<typeof ev> => ev !== null);

        if (eventsToSave.length === 0) {
          console.warn('[CALENDAR SYNC] No events with valid mailboxId to save. Total events:', allEvents.length);
          request.log.warn("No events with valid mailboxId to save");
        } else {
          // Get all unique mailboxIds from events that will be saved
          const eventMailboxIds = Array.from(
            new Set(eventsToSave.map((ev) => ev.mailboxId))
          );
          
          console.log('[CALENDAR SYNC] Saving to database:', {
            eventCount: eventsToSave.length,
            mailboxIds: eventMailboxIds,
            ownerUserId: actor.id,
            timezone: tz
          });
          
          request.log.info({ 
            eventCount: eventsToSave.length, 
            mailboxIds: eventMailboxIds,
            ownerUserId: actor.id 
          }, "Saving calendar events to database");

          // Save events to database for the admin user
          console.log('[CALENDAR SYNC] Calling replaceCalendarEvents...');
          await replaceCalendarEvents({
            ownerUserId: actor.id,
            mailboxIds: eventMailboxIds,
            timezone: tz,
            events: eventsToSave,
          });

          console.log('[CALENDAR SYNC] Successfully saved', eventsToSave.length, 'events to database');
          request.log.info({ eventCount: eventsToSave.length }, "Successfully saved calendar events to database");
        }
      } catch (err) {
        // Log error but don't fail the request - still return the events
        console.error('[CALENDAR SYNC] ERROR saving events:', err);
        request.log.error({ err, eventCount: allEvents.length }, "Failed to save calendar events to database");
      }
    } else {
      console.log('[CALENDAR SYNC] Save skipped - conditions not met:', {
        source,
        isAdmin: actor.role === "ADMIN",
        eventCount: allEvents.length
      });
    }

    // Check if save was attempted and add warning if no events were saved
    let saveWarning: string | undefined = undefined;
    if (source === "graph" && actor.role === "ADMIN" && allEvents.length > 0) {
      const eventsWithMailboxId = allEvents.filter((ev) => (ev as any).mailboxId);
      if (eventsWithMailboxId.length === 0) {
        saveWarning = "No events were saved to database (all events missing mailboxId)";
      } else if (eventsWithMailboxId.length < allEvents.length) {
        saveWarning = `${allEvents.length - eventsWithMailboxId.length} events were not saved (missing mailboxId)`;
      }
    }

    return {
      accounts: accountsList,
      events: allEvents.map((ev) => ({
        ...ev,
        mailboxId: (ev as any).mailboxId || undefined,
      })),
      source: "graph",
      warning: saveWarning || (failedMailboxes.length
        ? `Failed to load calendars for: ${failedMailboxes.join(", ")}.`
        : undefined),
      failedMailboxes,
    };
  });

  app.get("/daily-reports", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const parsed = z
      .object({
        start: z.string().optional(),
        end: z.string().optional(),
      })
      .safeParse(request.query);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }
    if (parsed.data.start && !isValidDateString(parsed.data.start)) {
      return reply.status(400).send({ message: "Invalid start date" });
    }
    if (parsed.data.end && !isValidDateString(parsed.data.end)) {
      return reply.status(400).send({ message: "Invalid end date" });
    }
    return listDailyReportsForUser(actor.id, {
      start: parsed.data.start ?? null,
      end: parsed.data.end ?? null,
    });
  });

  app.get("/daily-reports/by-date", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const parsed = z
      .object({
        date: z.string(),
      })
      .safeParse(request.query);
    if (!parsed.success || !isValidDateString(parsed.data.date)) {
      return reply.status(400).send({ message: "Invalid date" });
    }
    const report = await findDailyReportByUserAndDate(actor.id, parsed.data.date);
    return report ?? null;
  });

  app.put("/daily-reports/by-date", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const schema = z.object({
      date: z.string(),
      content: z.string().optional(),
      attachments: z
        .array(
          z.object({
            fileUrl: z.string().min(1),
            fileName: z.string().min(1),
            fileSize: z.number().nonnegative(),
            mimeType: z.string().min(1),
          }),
        )
        .optional(),
    });
    const body = schema.parse(request.body ?? {});
    if (!isValidDateString(body.date)) {
      return reply.status(400).send({ message: "Invalid date" });
    }
    const existing = await findDailyReportByUserAndDate(actor.id, body.date);
    if (existing?.status === "accepted") {
      return reply
        .status(409)
        .send({ message: "Accepted reports are read-only" });
    }
    const rawContent =
      body.content !== undefined ? body.content : existing?.content ?? null;
    const content = typeof rawContent === "string" ? rawContent.trim() : "";
    const updated = await upsertDailyReport({
      id: existing?.id ?? randomUUID(),
      userId: actor.id,
      reportDate: body.date,
      status: "draft",
      content: content ? content : null,
      reviewReason: existing?.reviewReason ?? null,
      submittedAt: null,
      reviewedAt: null,
      reviewedBy: null,
    });
    if (body.attachments?.length) {
      await insertDailyReportAttachments(updated.id, body.attachments);
    }
    return updated;
  });

  app.post("/daily-reports/by-date/send", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const schema = z.object({
      date: z.string(),
      content: z.string().optional(),
      attachments: z
        .array(
          z.object({
            fileUrl: z.string().min(1),
            fileName: z.string().min(1),
            fileSize: z.number().nonnegative(),
            mimeType: z.string().min(1),
          }),
        )
        .optional(),
    });
    const body = schema.parse(request.body ?? {});
    if (!isValidDateString(body.date)) {
      return reply.status(400).send({ message: "Invalid date" });
    }
    const existing = await findDailyReportByUserAndDate(actor.id, body.date);
    if (existing?.status === "accepted") {
      return reply
        .status(409)
        .send({ message: "Accepted reports are read-only" });
    }
    const rawContent =
      body.content !== undefined ? body.content : existing?.content ?? null;
    const content = typeof rawContent === "string" ? rawContent.trim() : "";
    const updated = await upsertDailyReport({
      id: existing?.id ?? randomUUID(),
      userId: actor.id,
      reportDate: body.date,
      status: "in_review",
      content: content ? content : null,
      reviewReason: null,
      submittedAt: new Date().toISOString(),
      reviewedAt: null,
      reviewedBy: null,
    });
    if (body.attachments?.length) {
      await insertDailyReportAttachments(updated.id, body.attachments);
    }
    // Notify admins only when a report is sent
    try {
      const reportDate = formatShortDate(body.date);
      await notifyAdmins({
        kind: "report",
        message: `${actor.userName} sent ${reportDate} report.`,
        href: "/admin/reports",
      });
    } catch (err) {
      request.log.error({ err }, "report send notification failed");
    }
    return updated;
  });

  app.patch("/daily-reports/:id/status", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can review reports" });
    }
    const { id } = request.params as { id: string };
    const schema = z.object({
      status: z.enum(["accepted", "rejected"]),
      reviewReason: z.string().optional().nullable(),
    });
    const body = schema.parse(request.body ?? {});
    const normalizedReason = trimString(body.reviewReason ?? "");
    if (body.status === "rejected" && !normalizedReason) {
      return reply.status(400).send({ message: "Rejection reason is required" });
    }
    const report = await findDailyReportById(id);
    if (!report) {
      return reply.status(404).send({ message: "Report not found" });
    }
    const updated = await updateDailyReportStatus({
      id,
      status: body.status,
      reviewedAt: new Date().toISOString(),
      reviewedBy: actor.id,
      reviewReason: body.status === "rejected" ? normalizedReason : null,
    });
    return updated;
  });

  app.get("/daily-reports/:id/attachments", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const { id } = request.params as { id: string };
    const report = await findDailyReportById(id);
    if (!report) {
      return reply.status(404).send({ message: "Report not found" });
    }
    const isReviewer = actor.role === "ADMIN" || actor.role === "MANAGER";
    if (!isReviewer && report.userId !== actor.id) {
      return reply.status(403).send({ message: "Not allowed to view attachments" });
    }
    return listDailyReportAttachments(id);
  });

  app.post("/daily-reports/upload", async (request: any, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });

    const data = await request.file();
    if (!data) return reply.status(400).send({ message: "No file provided" });

    const buffer = await data.toBuffer();
    const fileName = data.filename;
    const mimeType = data.mimetype;

    if (buffer.length > 10 * 1024 * 1024) {
      return reply.status(400).send({ message: "File too large. Max 10MB." });
    }

    const allowedTypes = [
      "image/jpeg",
      "image/png",
      "image/gif",
      "image/webp",
      "application/pdf",
      "application/zip",
      "text/plain",
      "text/csv",
    ];
    if (!allowedTypes.includes(mimeType)) {
      return reply.status(400).send({ message: "File type not supported" });
    }

    try {
      const { url } = await uploadToSupabase(buffer, fileName, mimeType);
      return {
        fileUrl: url,
        fileName,
        fileSize: buffer.length,
        mimeType,
      };
    } catch (err) {
      request.log.error({ err }, "Report file upload failed");
      return reply.status(500).send({ message: "Upload failed" });
    }
  });

  app.get("/notifications/summary", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const parsed = z
      .object({
        since: z.string().optional(),
      })
      .safeParse(request.query);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }
    let since: string | null = null;
    if (parsed.data.since) {
      const trimmed = trimString(parsed.data.since);
      if (trimmed) {
        const parsedDate = new Date(trimmed);
        if (Number.isNaN(parsedDate.getTime())) {
          return reply.status(400).send({ message: "Invalid since" });
        }
        since = parsedDate.toISOString();
      }
    }
    const isReviewer = actor.role === "ADMIN" || actor.role === "MANAGER";
    const reportCount = isReviewer
      ? await countDailyReportsInReview(since)
      : await countReviewedDailyReportsForUser(actor.id, since);
    const systemCount = await countUnreadNotifications(actor.id);
    return { reportCount, systemCount };
  });

  app.get("/notifications/list", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const parsed = z
      .object({
        since: z.string().optional(),
      })
      .safeParse(request.query);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }
    let since: string | null = null;
    if (parsed.data.since) {
      const trimmed = trimString(parsed.data.since);
      if (trimmed) {
        const parsedDate = new Date(trimmed);
        if (Number.isNaN(parsedDate.getTime())) {
          return reply.status(400).send({ message: "Invalid since" });
        }
        since = parsedDate.toISOString();
      }
    }
    const isReviewer = actor.role === "ADMIN" || actor.role === "MANAGER";
    const communityItems = await listUnreadCommunityNotifications(actor.id);
    const systemItems = await listNotificationsForUser(actor.id, {
      unreadOnly: true,
      limit: 50,
    });

    const notifications: NotificationSummary[] = [];

    communityItems.forEach((item) => {
      const senderName = item.senderName?.trim() || "Someone";
      if (item.threadType === "DM") {
        notifications.push({
          id: `community:${item.threadId}:${item.messageId ?? "latest"}`,
          kind: "community",
          message: `You received message from ${senderName}.`,
          createdAt: item.messageCreatedAt ?? new Date().toISOString(),
          href: "/community",
        });
      } else {
        const channelName = item.threadName?.trim() || "community";
        notifications.push({
          id: `community:${item.threadId}:${item.messageId ?? "latest"}`,
          kind: "community",
          message: `${senderName} posted a new message on #${channelName} channel.`,
          createdAt: item.messageCreatedAt ?? new Date().toISOString(),
          href: "/community",
        });
      }
    });

    if (isReviewer) {
      const reportItems = await listInReviewReportsWithUsers();
      reportItems.forEach((item) => {
        const reportDate = formatShortDate(item.reportDate);
        const submittedAt = item.submittedAt ?? item.updatedAt;
        const submittedTime = item.submittedAt ? Date.parse(item.submittedAt) : NaN;
        const updatedTime = Date.parse(item.updatedAt);
        const isUpdated =
          !Number.isNaN(submittedTime) &&
          !Number.isNaN(updatedTime) &&
          updatedTime > submittedTime;
        notifications.push({
          id: `report:${item.id}`,
          kind: "report",
          message: `${item.userName} ${isUpdated ? "updated" : "sent"} ${reportDate} report.`,
          createdAt: submittedAt,
          href: "/admin/reports",
        });
      });
    } else {
      const reportItems = await listReviewedDailyReportsForUser(actor.id, since);
      reportItems.forEach((item) => {
        const reportDate = formatShortDate(item.reportDate);
        const statusLabel = item.status === "accepted" ? "accepted" : "rejected";
        notifications.push({
          id: `report:${item.id}`,
          kind: "report",
          message: `${reportDate} report ${statusLabel}.`,
          createdAt: item.reviewedAt,
          href: "/reports",
        });
      });
    }

    systemItems.forEach((item) => {
      notifications.push({
        id: item.id,
        kind: "system",
        message: item.message,
        createdAt: item.createdAt,
        href: item.href ?? undefined,
      });
    });

    if (systemItems.length > 0) {
      await markNotificationsRead(
        actor.id,
        systemItems.map((item) => item.id),
      );
    }

    notifications.sort((a, b) => {
      const aTime = Date.parse(a.createdAt);
      const bTime = Date.parse(b.createdAt);
      if (Number.isNaN(aTime) || Number.isNaN(bTime)) return 0;
      return bTime - aTime;
    });

    return { notifications };
  });

  app.get("/admin/daily-reports/by-date", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can view reports" });
    }
    const parsed = z
      .object({
        date: z.string(),
      })
      .safeParse(request.query);
    if (!parsed.success || !isValidDateString(parsed.data.date)) {
      return reply.status(400).send({ message: "Invalid date" });
    }
    return listDailyReportsByDate(parsed.data.date);
  });

  app.get("/admin/daily-reports/in-review", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can view reports" });
    }
    const parsed = z
      .object({
        start: z.string(),
        end: z.string(),
      })
      .safeParse(request.query);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }
    if (!isValidDateString(parsed.data.start) || !isValidDateString(parsed.data.end)) {
      return reply.status(400).send({ message: "Invalid date range" });
    }
    return listInReviewReports({
      start: parsed.data.start,
      end: parsed.data.end,
    });
  });

  app.get("/admin/daily-reports/accepted-by-date", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can view reports" });
    }
    const parsed = z
      .object({
        start: z.string(),
        end: z.string(),
      })
      .safeParse(request.query);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }
    if (!isValidDateString(parsed.data.start) || !isValidDateString(parsed.data.end)) {
      return reply.status(400).send({ message: "Invalid date range" });
    }
    return listAcceptedCountsByDate({
      start: parsed.data.start,
      end: parsed.data.end,
    });
  });

  app.get("/admin/daily-reports/by-user", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can view reports" });
    }
    const parsed = z
      .object({
        userId: z.string().uuid(),
        start: z.string().optional(),
        end: z.string().optional(),
      })
      .safeParse(request.query);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid query" });
    }
    if (parsed.data.start && !isValidDateString(parsed.data.start)) {
      return reply.status(400).send({ message: "Invalid start date" });
    }
    if (parsed.data.end && !isValidDateString(parsed.data.end)) {
      return reply.status(400).send({ message: "Invalid end date" });
    }
    return listDailyReportsForUser(parsed.data.userId, {
      start: parsed.data.start ?? null,
      end: parsed.data.end ?? null,
    });
  });

  app.get("/assignments", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    return listAssignments();
  });
  app.post("/assignments", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can assign profiles" });
    }
    const schema = z.object({
      profileId: z.string(),
      bidderUserId: z.string(),
      assignedBy: z.string().optional(),
    });
    const body = schema.parse(request.body);
    const profile = await findProfileById(body.profileId);
    const bidder = await findUserById(body.bidderUserId);
    if (!profile || !bidder || bidder.role !== "BIDDER") {
      return reply.status(400).send({ message: "Invalid profile or bidder" });
    }

    const existing = await findActiveAssignmentByProfile(body.profileId);
    if (existing) {
      return reply.status(409).send({
        message: "Profile already assigned",
        assignmentId: existing.id,
      });
    }

    const newAssignment: Assignment = {
      id: body.profileId,
      profileId: body.profileId,
      bidderUserId: body.bidderUserId,
      assignedBy: actor.id ?? body.assignedBy ?? body.bidderUserId,
      assignedAt: new Date().toISOString(),
      unassignedAt: null as string | null,
    };
    await insertAssignmentRecord(newAssignment);
    events.push({
      id: randomUUID(),
      sessionId: "admin-event",
      eventType: "ASSIGNED",
      payload: { profileId: body.profileId, bidderUserId: body.bidderUserId },
      createdAt: new Date().toISOString(),
    });
    try {
      await notifyAdmins({
        kind: "system",
        message: `Profile ${profile.displayName} assigned to ${bidder.name}.`,
        href: "/manager/profiles",
      });
      await notifyUsers([bidder.id], {
        kind: "system",
        message: `You were assigned profile ${profile.displayName}.`,
        href: "/workspace",
      });
    } catch (err) {
      request.log.error({ err }, "assignment notification failed");
    }
    return newAssignment;
  });

  app.post("/assignments/:id/unassign", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const { id } = request.params as { id: string };
    const assignment = await closeAssignmentById(id);
    if (!assignment)
      return reply.status(404).send({ message: "Assignment not found" });
    events.push({
      id: randomUUID(),
      sessionId: "admin-event",
      eventType: "UNASSIGNED",
      payload: {
        profileId: assignment.profileId,
        bidderUserId: assignment.bidderUserId,
      },
      createdAt: new Date().toISOString(),
    });
    return assignment;
  });

  const ensureCommunityThreadAccess = async (
    threadId: string,
    actor: User,
    reply: any
  ) => {
    const thread = await findCommunityThreadById(threadId);
    if (!thread) {
      reply.status(404).send({ message: "Thread not found" });
      return undefined;
    }
    if (thread.threadType === "DM" || thread.isPrivate) {
      const isMember = await isCommunityThreadMember(threadId, actor.id);
      if (!isMember) {
        reply.status(403).send({ message: "Not a member of this thread" });
        return undefined;
      }
    }
    return thread;
  };

  app.get("/community/overview", async (request, reply) => {
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });
    const [channels, dms] = await Promise.all([
      listCommunityChannels(),
      listCommunityDmThreads(actor.id),
    ]);
    return { channels, dms };
  });

  app.get("/community/channels", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can manage channels" });
    }
    const channels = await listCommunityChannels();
    return channels;
  });

  app.post("/community/channels", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });
    const schema = z.object({
      name: z.string(),
      description: z.string().optional(),
    });
    const body = schema.parse(request.body);
    const name = normalizeChannelName(body.name);
    if (!name)
      return reply.status(400).send({ message: "Channel name required" });
    const nameKey = name.toLowerCase();
    const existing = await findCommunityChannelByKey(nameKey);
    if (existing) {
      return reply
        .status(409)
        .send({ message: "Channel already exists", channel: existing });
    }
    const created = await insertCommunityThread({
      id: randomUUID(),
      threadType: "CHANNEL",
      name,
      nameKey,
      description: body.description?.trim() || null,
      createdBy: actor.id,
      isPrivate: false,
    });
    await insertCommunityThreadMember({
      id: randomUUID(),
      threadId: created.id,
      userId: actor.id,
      role: "OWNER",
    });
    try {
      await notifyAllUsers({
        kind: "system",
        message: `Channel #${created.name} created.`,
        href: "/community",
      });
    } catch (err) {
      request.log.error({ err }, "channel create notification failed");
    }
    return created;
  });

  app.patch("/community/channels/:id", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can manage channels" });
    }
    const { id } = request.params as { id: string };
    const schema = z.object({
      name: z.string().optional(),
      description: z.string().optional(),
    });
    const body = schema.parse(request.body ?? {});
    if (body.name === undefined && body.description === undefined) {
      return reply.status(400).send({ message: "No updates provided" });
    }
    const existing = await findCommunityThreadById(id);
    if (!existing || existing.threadType !== "CHANNEL") {
      return reply.status(404).send({ message: "Channel not found" });
    }
    const nameInput = body.name ?? existing.name ?? "";
    const name = normalizeChannelName(nameInput);
    if (!name) {
      return reply.status(400).send({ message: "Channel name required" });
    }
    const nameKey = name.toLowerCase();
    const conflict = await findCommunityChannelByKey(nameKey);
    if (conflict && conflict.id !== id) {
      return reply
        .status(409)
        .send({ message: "Channel already exists" });
    }
    const description =
      body.description === undefined
        ? existing.description ?? null
        : body.description.trim() || null;
    const updated = await updateCommunityChannel({
      id,
      name,
      nameKey,
      description,
    });
    if (updated) {
      try {
        await notifyAllUsers({
          kind: "system",
          message: `Channel #${updated.name} updated.`,
          href: "/community",
        });
      } catch (err) {
        request.log.error({ err }, "channel update notification failed");
      }
    }
    return updated ?? reply.status(404).send({ message: "Channel not found" });
  });

  app.delete("/community/channels/:id", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can manage channels" });
    }
    const { id } = request.params as { id: string };
    const existing = await findCommunityThreadById(id);
    if (!existing || existing.threadType !== "CHANNEL") {
      return reply.status(404).send({ message: "Channel not found" });
    }
    await deleteCommunityChannel(id);
    try {
      await notifyAllUsers({
        kind: "system",
        message: `Channel #${existing.name ?? "channel"} removed.`,
        href: "/community",
      });
    } catch (err) {
      request.log.error({ err }, "channel delete notification failed");
    }
    return { status: "deleted", id };
  });

  app.post("/community/dms", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });
    const schema = z.object({ userId: z.string() });
    const body = schema.parse(request.body);
    if (body.userId === actor.id) {
      return reply
        .status(400)
        .send({ message: "Cannot start a DM with yourself" });
    }
    const other = await findUserById(body.userId);
    if (!other || other.isActive === false) {
      return reply.status(404).send({ message: "User not found" });
    }
    const existingId = await findCommunityDmThreadId(actor.id, body.userId);
    if (existingId) {
      const summary = await getCommunityDmThreadSummary(existingId, actor.id);
      if (summary) return summary;
      return {
        id: existingId,
        threadType: "DM",
        isPrivate: true,
        createdAt: new Date().toISOString(),
        participants: [{ id: other.id, name: other.name, email: other.email }],
      };
    }
    const thread = await insertCommunityThread({
      id: randomUUID(),
      threadType: "DM",
      name: null,
      nameKey: null,
      description: null,
      createdBy: actor.id,
      isPrivate: true,
    });
    await insertCommunityThreadMember({
      id: randomUUID(),
      threadId: thread.id,
      userId: actor.id,
      role: "MEMBER",
    });
    await insertCommunityThreadMember({
      id: randomUUID(),
      threadId: thread.id,
      userId: other.id,
      role: "MEMBER",
    });
    const summary = await getCommunityDmThreadSummary(thread.id, actor.id);
    return (
      summary ?? {
        id: thread.id,
        threadType: "DM",
        isPrivate: true,
        createdAt: thread.createdAt,
        participants: [
          {
            id: other.id,
            name: other.name,
            email: other.email,
            avatarUrl: other.avatarUrl ?? null,
          },
        ],
      }
    );
  });

  app.get("/community/threads/:id/messages", async (request, reply) => {
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });
    const { id } = request.params as { id: string };
    const query = request.query as {
      limit?: string;
      before?: string;
      after?: string;
    };
    const thread = await ensureCommunityThreadAccess(id, actor, reply);
    if (!thread) return;

    const limit = query.limit ? parseInt(query.limit, 10) : 50;
    const messages = await listCommunityMessagesWithPagination(id, {
      limit,
      before: query.before,
      after: query.after,
    });

    const enriched = await Promise.all(
      messages.map(async (msg) => {
        const attachments = await listMessageAttachments(msg.id);
        const reactions = await listMessageReactions(msg.id, actor.id);
        const readReceipts = await getMessageReadReceipts(msg.id);
        let replyPreview = null;

        if (msg.replyToMessageId) {
          const target = await getMessageById(msg.replyToMessageId);
          if (target && !target.isDeleted) {
            replyPreview = {
              id: target.id,
              senderId: target.senderId,
              senderName: target.senderName ?? null,
              body: target.body.substring(0, 100),
            };
          }
        }

        return {
          ...msg,
          attachments,
          reactions,
          replyPreview,
          readReceipts,
        };
      })
    );

    if (messages.length > 0) {
      await markThreadAsRead(id, actor.id, messages[messages.length - 1].id);
    }

    return enriched;
  });

  app.post("/community/threads/:id/messages", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });
    const { id } = request.params as { id: string };
    const schema = z.object({
      body: z.string(),
      replyToMessageId: z.string().optional(),
      attachments: z
        .array(
          z.object({
            fileName: z.string(),
            fileUrl: z.string(),
            fileSize: z.number(),
            mimeType: z.string(),
            thumbnailUrl: z.string().optional(),
            width: z.number().optional(),
            height: z.number().optional(),
          })
        )
        .optional(),
    });
    const body = schema.parse(request.body);
    const text = body.body.trim();
    if (!text && (!body.attachments || body.attachments.length === 0))
      return reply.status(400).send({ message: "Message body or attachments required" });
    const thread = await ensureCommunityThreadAccess(id, actor, reply);
    if (!thread) return;

    if (body.replyToMessageId) {
      const replyTarget = await getMessageById(body.replyToMessageId);
      if (!replyTarget || replyTarget.threadId !== id) {
        return reply.status(400).send({ message: "Invalid reply target" });
      }
    }

    if (thread.threadType === "CHANNEL") {
      await insertCommunityThreadMember({
        id: randomUUID(),
        threadId: id,
        userId: actor.id,
        role: "MEMBER",
      });
    }
    const message = await insertCommunityMessage({
      id: randomUUID(),
      threadId: id,
      senderId: actor.id,
      body: text || '',
      replyToMessageId: body.replyToMessageId ?? null,
      isEdited: false,
      isDeleted: false,
      createdAt: new Date().toISOString(),
    });

    if (body.attachments && body.attachments.length > 0) {
      for (const att of body.attachments) {
        await insertMessageAttachment({
          id: randomUUID(),
          messageId: message.id,
          fileName: att.fileName,
          fileUrl: att.fileUrl,
          fileSize: att.fileSize,
          mimeType: att.mimeType,
          thumbnailUrl: att.thumbnailUrl ?? null,
          width: att.width ?? null,
          height: att.height ?? null,
          createdAt: new Date().toISOString(),
        });
      }
    }

    await incrementUnreadCount(id, actor.id);

    try {
      await broadcastCommunityMessage(id, message);
    } catch (err) {
      request.log.error({ err }, "community realtime broadcast failed");
    }
    return message;
  });

  app.get("/sessions/:id", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const { id } = request.params as { id: string };
    const session = sessions.find((s) => s.id === id);
    if (!session)
      return reply.status(404).send({ message: "Session not found" });
    return session;
  });

  app.post("/sessions", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    const schema = z.object({
      bidderUserId: z.string(),
      profileId: z.string(),
      url: z.string(),
    });
    const body = schema.parse(request.body);
    const profileAssignment = await findActiveAssignmentByProfile(
      body.profileId
    );

    let bidderUserId = body.bidderUserId;
    if (actor.role === "BIDDER") {
      bidderUserId = actor.id;
      if (profileAssignment && profileAssignment.bidderUserId !== actor.id) {
        return reply
          .status(403)
          .send({ message: "Profile not assigned to bidder" });
      }
    } else if (actor.role === "MANAGER" || actor.role === "ADMIN") {
      if (!bidderUserId && profileAssignment)
        bidderUserId = profileAssignment.bidderUserId;
      if (!bidderUserId) bidderUserId = actor.id;
    } else {
      return reply.status(403).send({ message: "Forbidden" });
    }

    const session: ApplicationSession = {
      id: randomUUID(),
      bidderUserId,
      profileId: body.profileId,
      url: body.url,
      domain: tryExtractDomain(body.url),
      status: "OPEN",
      startedAt: new Date().toISOString(),
    };
    sessions.unshift(session);
    events.push({
      id: randomUUID(),
      sessionId: session.id,
      eventType: "SESSION_CREATED",
      payload: { url: session.url },
      createdAt: new Date().toISOString(),
    });
    return session;
  });

  app.post("/sessions/:id/go", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const { id } = request.params as { id: string };
    const session = sessions.find((s) => s.id === id);
    if (!session)
      return reply.status(404).send({ message: "Session not found" });
    session.status = "OPEN";
    try {
      await startBrowserSession(session);
    } catch (err) {
      app.log.error({ err }, "failed to start browser session");
    }
    events.push({
      id: randomUUID(),
      sessionId: id,
      eventType: "GO_CLICKED",
      payload: { url: session.url },
      createdAt: new Date().toISOString(),
    });
    return { ok: true };
  });

  app.post("/sessions/:id/analyze", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const { id } = request.params as { id: string };
    const session = sessions.find((s) => s.id === id);
    if (!session)
      return reply.status(404).send({ message: "Session not found" });
    const body = (request.body as { useAi?: boolean } | undefined) ?? {};
    const useAi = Boolean(body.useAi);

    const live = livePages.get(id);
    const page = live?.page;
    if (!page) {
      return reply.status(400).send({
        message:
          "Live page not available. Click Go and load the page before Analyze.",
      });
    }

    let pageHtml = "";
    let pageTitle = "";
    try {
      pageTitle = await page.title();
      pageHtml = await page.content();
    } catch (err) {
      request.log.error({ err }, "failed to read live page content");
    }
    if (!pageHtml) {
      return reply.status(400).send({
        message: "No page content captured. Load the page before Analyze.",
      });
    }

    const analysis = await analyzeJobFromHtml(pageHtml, pageTitle);

    session.status = "ANALYZED";
    session.jobContext = {
      title: analysis.title || "Job",
      company: "N/A",
      summary: "Analysis from job description",
      job_description_text: analysis.jobText ?? "",
    };

    if (!useAi) {
      const topTech = (analysis.ranked ?? []).slice(0, 4);
      events.push({
        id: randomUUID(),
        sessionId: id,
        eventType: "ANALYZE_DONE",
        payload: {
          recommendedLabel: analysis.recommendedLabel,
        },
        createdAt: new Date().toISOString(),
      });
      return {
        mode: "tech",
        recommendedLabel: analysis.recommendedLabel,
        ranked: topTech.map((t, idx) => ({
          id: t.id ?? t.label ?? `tech-${idx}`,
          label: t.label,
          rank: idx + 1,
          score: t.score,
        })),
        scores: analysis.rawScores,
        jobContext: session.jobContext,
      };
    }

    events.push({
      id: randomUUID(),
      sessionId: id,
      eventType: "ANALYZE_DONE",
      payload: {
        recommendedLabel: analysis.recommendedLabel,
        recommendedResumeId: null,
      },
      createdAt: new Date().toISOString(),
    });
    return {
      mode: "resume",
      recommendedResumeId: null,
      recommendedLabel: analysis.recommendedLabel,
      ranked: [],
      scores: {},
      jobContext: session.jobContext,
    };
  });

  // Prompt-pack endpoints (HF-backed)
  app.post("/llm/resume-parse", async (request, reply) => {
    const { resumeText, resumeId, filename, baseProfile } = request.body as any;
    if (!resumeText || !resumeId)
      return reply
        .status(400)
        .send({ message: "resumeText and resumeId are required" });
    const prompt = promptBuilders.buildResumeParsePrompt({
      resumeId,
      filename,
      resumeText,
      baseProfile,
    });
    const parsed = await callPromptPack(prompt);
    if (!parsed) return reply.status(502).send({ message: "LLM parse failed" });
    return parsed;
  });

  app.post("/llm/job-analyze", async (request, reply) => {
    const { job, baseProfile, prefs } = request.body as any;
    if (!job?.job_description_text)
      return reply
        .status(400)
        .send({ message: "job_description_text required" });
    const prompt = promptBuilders.buildJobAnalyzePrompt({
      job,
      baseProfile,
      prefs,
    });
    const parsed = await callPromptPack(prompt);
    if (!parsed)
      return reply.status(502).send({ message: "LLM analyze failed" });
    return parsed;
  });

  app.post("/llm/rank-resumes", async (request, reply) => {
    const { job, resumes, baseProfile, prefs } = request.body as any;
    if (!job?.job_description_text || !Array.isArray(resumes)) {
      return reply
        .status(400)
        .send({ message: "job_description_text and resumes[] required" });
    }
    const prompt = promptBuilders.buildRankResumesPrompt({
      job,
      resumes,
      baseProfile,
      prefs,
    });
    const parsed = await callPromptPack(prompt);
    if (!parsed) return reply.status(502).send({ message: "LLM rank failed" });
    return parsed;
  });

  app.post("/llm/autofill-plan", async (request, reply) => {
    const {
      pageFields,
      baseProfile,
      prefs,
      jobContext,
      selectedResume,
      pageContext,
    } = request.body as any;
    if (!Array.isArray(pageFields))
      return reply.status(400).send({ message: "pageFields[] required" });
    const prompt = promptBuilders.buildAutofillPlanPrompt({
      pageFields,
      baseProfile,
      prefs,
      jobContext,
      selectedResume,
      pageContext,
    });
    const parsed = await callPromptPack(prompt);
    if (!parsed)
      return reply.status(502).send({ message: "LLM autofill failed" });
    return parsed;
  });

  app.post("/autofill/greenhouse-ai", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const schema = z.object({
      questions: z.array(
        z.object({
          id: z.string().min(1),
          type: z.enum([
            "text",
            "textarea",
            "select",
            "multi_value_single_select",
            "checkbox",
            "file",
          ]),
          label: z.string().min(1),
          required: z.boolean().optional().default(false),
          options: z.array(z.string()).optional(),
        })
      ),
      profile: z.record(z.any()),
      apiKey: z.string().optional(),
      model: z.string().optional(),
    });

    const body = schema.parse(request.body);
    const apiKey = trimString(body.apiKey) || trimString(process.env.OPENAI_API_KEY);
    if (!apiKey) {
      return reply.status(400).send({ message: "OPENAI_API_KEY is required" });
    }
    const model =
      trimString(body.model) || trimString(process.env.OPENAI_GREENHOUSE_MODEL) || "gpt-4";
    const prompt = buildGreenhousePrompt(body.questions, body.profile ?? {});

    try {
      const content = await callChatCompletion({
        provider: "OPENAI",
        model,
        apiKey,
        systemPrompt: GREENHOUSE_AI_SYSTEM_PROMPT,
        userPrompt: prompt,
        temperature: 0.7,
        maxTokens: 2000,
      });
      if (!content) {
        return reply.status(502).send({ message: "LLM response empty" });
      }
      const parsed = extractJsonArrayPayload(content);
      if (!parsed) {
        return reply.status(502).send({ message: "LLM response not parseable" });
      }
      return { answers: parsed, provider: "OPENAI", model };
    } catch (err) {
      request.log.error({ err }, "Greenhouse AI failed");
      return reply.status(502).send({ message: "Greenhouse AI failed" });
    }
  });

  app.post("/llm/tailor-resume", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const schema = z.object({
      jobDescriptionText: z.string().min(1),
      baseResume: z.record(z.any()).optional(),
      baseResumeText: z.string().optional(),
      bulletCountByCompany: z.record(z.number()).optional(),
      systemPrompt: z.string().optional(),
      userPrompt: z.string().optional(),
      provider: z.enum(["OPENAI", "HUGGINGFACE", "GEMINI"]).optional(),
      model: z.string().optional(),
      apiKey: z.string().optional(),
    });
    const body = schema.parse(request.body);
    const { provider, model, apiKey } = resolveLlmConfig({
      provider: body.provider,
      model: body.model,
      apiKey: body.apiKey ?? null,
    });
    if (!apiKey) {
      return reply.status(400).send({ message: "LLM apiKey is required" });
    }
    const promptBaseResume = buildPromptBaseResume(body.baseResume ?? {});
    const baseResumeJson = JSON.stringify(promptBaseResume, null, 2);
    const systemPrompt =
      body.systemPrompt?.trim() || DEFAULT_TAILOR_SYSTEM_PROMPT;
    const userPrompt =
      buildTailorUserPrompt({
        jobDescriptionText: body.jobDescriptionText,
        baseResumeJson,
        bulletCountByCompany: body.bulletCountByCompany,
        userPromptTemplate: body.userPrompt ?? null,
      });
    try {
      const content = await callChatCompletion({
        provider,
        model,
        apiKey,
        systemPrompt,
        userPrompt,
      });
      if (!content) {
        return reply.status(502).send({ message: "LLM response empty" });
      }
      const parsed = extractJsonPayload(content);
      return { content, parsed, provider, model };
    } catch (err) {
      request.log.error({ err }, "LLM tailor resume failed");
      return reply.status(502).send({ message: "LLM tailor failed" });
    }
  });

  app.post("/llm/interview-chat", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const schema = z.object({
      resumeJson: z.record(z.any()).optional(),
      jobDescription: z.string().min(1),
      question: z.string().min(1),
      provider: z.enum(["OPENAI", "HUGGINGFACE", "GEMINI"]).optional(),
      model: z.string().optional(),
      apiKey: z.string().optional(),
    });
    const body = schema.parse(request.body);
    const { provider, model, apiKey } = resolveLlmConfig({
      provider: body.provider,
      model: body.model,
      apiKey: body.apiKey ?? null,
    });
    if (!apiKey) {
      return reply.status(400).send({ message: "LLM apiKey is required" });
    }
    const systemPrompt = `You are an interview Q&A chatbot for a job candidate.

Inputs:
- resume_json (facts about the candidate)
- job_description (JD)
- question

Answer rules:
1) Write exactly 2–3 short, simple sentences. No bullet points.
2) Prefer JD-aligned answers (JD-first). Use the JD to frame what matters and what the role expects.
3) Then support with resume facts when available (resume evidence). If resume lacks a specific detail, do NOT claim experience you don't have.
4) When resume lacks details, answer using a capability/approach based on the JD, using wording like:
   "Based on this role, I would..." or "In this position, I would..."
5) Never invent employers, dates, metrics, tools, or projects not in resume_json.
6) Use first-person voice ("I").`;
    
    const resumeJsonStr = body.resumeJson ? JSON.stringify(body.resumeJson, null, 2) : "{}";
    const userPrompt = `JOB_DESCRIPTION:
${body.jobDescription}

RESUME_JSON:
${resumeJsonStr}

QUESTION:
${body.question}`;
    
    try {
      const content = await callChatCompletion({
        provider,
        model,
        apiKey,
        systemPrompt,
        userPrompt,
        temperature: 0.7,
        maxTokens: 300,
      });
      if (!content) {
        return reply.status(502).send({ message: "LLM response empty" });
      }
      return { content, provider, model };
    } catch (err) {
      request.log.error({ err }, "LLM interview chat failed");
      return reply.status(502).send({ message: "LLM chat failed" });
    }
  });

  app.get("/label-aliases", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can manage label aliases" });
    }
    const custom = await listLabelAliases();
    return { defaults: DEFAULT_LABEL_ALIASES, custom };
  });

  app.get("/application-phrases", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const custom = await listLabelAliases();
    const phrases = buildApplicationSuccessPhrases(custom);
    return { phrases };
  });

  app.post("/label-aliases", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can manage label aliases" });
    }
    const schema = z.object({
      canonicalKey: z.string(),
      alias: z.string().min(2),
    });
    const body = schema.parse(request.body ?? {});
    const canonicalKey = body.canonicalKey.trim();
    if (!CANONICAL_LABEL_KEYS.has(canonicalKey)) {
      return reply.status(400).send({ message: "Unknown canonical key" });
    }
    const normalizedAlias = normalizeLabelAlias(body.alias);
    if (!normalizedAlias) {
      return reply.status(400).send({ message: "Alias cannot be empty" });
    }
    const existing = await findLabelAliasByNormalized(normalizedAlias);
    if (existing) {
      return reply.status(409).send({ message: "Alias already exists" });
    }
    const aliasRecord: LabelAlias = {
      id: randomUUID(),
      canonicalKey,
      alias: body.alias.trim(),
      normalizedAlias,
    };
    await insertLabelAlias(aliasRecord);
    return aliasRecord;
  });

  app.patch("/label-aliases/:id", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can manage label aliases" });
    }
    const { id } = request.params as { id: string };
    const schema = z.object({
      canonicalKey: z.string().optional(),
      alias: z.string().optional(),
    });
    const body = schema.parse(request.body ?? {});
    const existing = await findLabelAliasById(id);
    if (!existing)
      return reply.status(404).send({ message: "Alias not found" });

    const canonicalKey = body.canonicalKey?.trim() || existing.canonicalKey;
    if (!CANONICAL_LABEL_KEYS.has(canonicalKey)) {
      return reply.status(400).send({ message: "Unknown canonical key" });
    }
    const aliasText = (body.alias ?? existing.alias).trim();
    const normalizedAlias = normalizeLabelAlias(aliasText);
    if (!normalizedAlias) {
      return reply.status(400).send({ message: "Alias cannot be empty" });
    }
    const conflict = await findLabelAliasByNormalized(normalizedAlias);
    if (conflict && conflict.id !== id) {
      return reply.status(409).send({ message: "Alias already exists" });
    }
    const updated: LabelAlias = {
      ...existing,
      canonicalKey,
      alias: aliasText,
      normalizedAlias,
      updatedAt: new Date().toISOString(),
    };
    await updateLabelAliasRecord(updated);
    return updated;
  });

  app.delete("/label-aliases/:id", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can manage label aliases" });
    }
    const { id } = request.params as { id: string };
    const existing = await findLabelAliasById(id);
    if (!existing)
      return reply.status(404).send({ message: "Alias not found" });
    await deleteLabelAlias(id);
    return { status: "deleted", id };
  });

  app.post("/sessions/:id/autofill", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const { id } = request.params as { id: string };
    const body =
      (request.body as {
        pageFields?: any[];
        useLlm?: boolean;
      }) ?? {};
    const session = sessions.find((s) => s.id === id);
    if (!session)
      return reply.status(404).send({ message: "Session not found" });
    const profile = await findProfileById(session.profileId);
    if (!profile)
      return reply.status(404).send({ message: "Profile not found" });

    const live = livePages.get(id);
    const page = live?.page;
    const hasClientFields = Array.isArray(body.pageFields);
    let pageFields: any[] = hasClientFields ? body.pageFields ?? [] : [];
    if (!hasClientFields && page) {
      try {
        pageFields = await collectPageFields(page);
      } catch (err) {
        request.log.error({ err }, "collectPageFields failed");
      }
    }

    const candidateFields: any[] = pageFields.length
      ? pageFields
      : DEFAULT_AUTOFILL_FIELDS;

    const autofillValues = buildAutofillValueMap(
      profile.baseInfo ?? {},
      session.jobContext ?? {}
    );
    const aliasIndex = buildAliasIndex(await listLabelAliases());
    const useLlm = body.useLlm !== false;

    let fillPlan: FillPlanResult = {
      filled: [],
      suggestions: [],
      blocked: [],
      actions: [],
    };
    if (candidateFields.length > 0) {
      try {
        fillPlan = buildAliasFillPlan(
          candidateFields,
          aliasIndex,
          autofillValues
        );
      } catch (err) {
        request.log.error({ err }, "label-db autofill failed");
        fillPlan = { filled: [], suggestions: [], blocked: [], actions: [] };
      }
    }

    try {
      if (
        useLlm &&
        (!fillPlan.filled || fillPlan.filled.length === 0) &&
        candidateFields.length > 0
      ) {
        const prompt = promptBuilders.buildAutofillPlanPrompt({
          pageFields: candidateFields,
          baseProfile: profile.baseInfo ?? {},
          prefs: {},
          jobContext: session.jobContext ?? {},
          pageContext: { url: session.url },
        });
        const parsed = await callPromptPack(prompt);
        const llmPlan = parsed?.result?.fill_plan;
        if (Array.isArray(llmPlan)) {
          const filteredPlan = llmPlan.filter(
            (f: any) => !shouldSkipPlanField(f, aliasIndex)
          );
          const actions: FillPlanAction[] = filteredPlan
            .map((f: any) => ({
              field: String(f.field_id ?? f.selector ?? f.label ?? "field"),
              field_id: typeof f.field_id === "string" ? f.field_id : undefined,
              label: typeof f.label === "string" ? f.label : undefined,
              selector: typeof f.selector === "string" ? f.selector : undefined,
              action: (f.action as FillPlanAction["action"]) ?? "fill",
              value:
                typeof f.value === "string"
                  ? f.value
                  : JSON.stringify(f.value ?? ""),
              confidence:
                typeof f.confidence === "number" ? f.confidence : undefined,
            }))
            .filter((f) => f.action !== "skip");
          const filledFromPlan = actions
            .filter((f) =>
              ["fill", "select", "check", "uncheck"].includes(f.action)
            )
            .map((f) => ({
              field: f.field,
              value: f.value ?? "",
              confidence: f.confidence,
            }));
          const suggestions =
            (Array.isArray(parsed?.warnings) ? parsed?.warnings : []).map(
              (w: any) => ({
                field: "note",
                suggestion: String(w),
              })
            ) ?? [];
          const blocked = llmPlan
            .filter((f: any) => f.requires_user_review)
            .map((f: any) => f.field_id ?? f.selector ?? "field");
          fillPlan = {
            filled: filledFromPlan,
            suggestions,
            blocked,
            actions,
          };
        }
      }
    } catch (err) {
      request.log.error({ err }, "LLM autofill failed, using demo plan");
    }

    if (
      !fillPlan.filled?.length &&
      !fillPlan.suggestions?.length &&
      !fillPlan.blocked?.length
    ) {
      fillPlan = buildDemoFillPlan(profile.baseInfo);
    }

    session.status = "FILLED";
    session.fillPlan = fillPlan;
    events.push({
      id: randomUUID(),
      sessionId: id,
      eventType: "AUTOFILL_DONE",
      payload: session.fillPlan,
      createdAt: new Date().toISOString(),
    });
    return { fillPlan: session.fillPlan, pageFields, candidateFields };
  });

  app.post("/sessions/:id/mark-submitted", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const { id } = request.params as { id: string };
    const session = sessions.find((s) => s.id === id);
    if (!session)
      return reply.status(404).send({ message: "Session not found" });
    session.status = "SUBMITTED";
    session.endedAt = new Date().toISOString();
    try {
      const record: ApplicationRecord = {
        id: randomUUID(),
        sessionId: id,
        bidderUserId: session.bidderUserId,
        profileId: session.profileId,
        resumeId: null,
        url: session.url ?? "",
        domain: session.domain ?? tryExtractDomain(session.url ?? ""),
        createdAt: new Date().toISOString(),
        status: "in_review",
        isReviewed: false,
        reviewedAt: null,
        reviewedBy: null,
      };
      await insertApplication(record);
    } catch (err) {
      request.log.error({ err }, "failed to insert application record");
    }
    await stopBrowserSession(id);
    events.push({
      id: randomUUID(),
      sessionId: id,
      eventType: "SUBMITTED",
      createdAt: new Date().toISOString(),
    });
    return { status: session.status };
  });

  app.get("/sessions", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const bidderUserId = (request.query as { bidderUserId?: string })
      .bidderUserId;
    const filtered = bidderUserId
      ? sessions.filter((s) => s.bidderUserId === bidderUserId)
      : sessions;
    return filtered;
  });

  app.post("/users/me/avatar", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }

    const data = await request.file();
    if (!data) return reply.status(400).send({ message: "No file provided" });

    const buffer = await data.toBuffer();
    if (buffer.length > 5 * 1024 * 1024) {
      return reply.status(400).send({ message: "File too large. Max 5MB." });
    }

    const fileName = data.filename;
    const mimeType = data.mimetype;
    const allowedTypes = ["image/jpeg", "image/png", "image/webp", "image/gif"];
    if (!allowedTypes.includes(mimeType)) {
      return reply.status(400).send({ message: "Image type not supported" });
    }

    try {
      const { url } = await uploadToSupabase(buffer, fileName, mimeType);
      await updateUserAvatar(actor.id, url);
      const updated = await findUserById(actor.id);
      return { user: updated, avatarUrl: url };
    } catch (err) {
      request.log.error({ err }, "avatar upload failed");
      return reply.status(500).send({ message: "Avatar upload failed" });
    }
  });

  app.patch("/users/me", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }

    const schema = z.object({
      name: z.string().min(1).max(200),
      email: z.string().email().max(200),
      bio: z.string().max(1000).nullable().optional(),
    });
    const parsed = schema.safeParse(request.body ?? {});
    if (!parsed.success) {
      const issue = parsed.error.errors[0];
      const field = issue?.path?.[0];
      const message = `${field ? `${field}: ` : ""}${issue?.message ?? "Invalid payload"}`;
      return reply.status(400).send({ message });
    }

    const { name, email, bio } = parsed.data;

    // Check if email is already taken by another user
    const existingUser = await findUserByEmail(email);
    if (existingUser && existingUser.id !== actor.id) {
      return reply.status(409).send({ message: "Email already taken" });
    }

    try {
      await updateUserNameAndEmail(actor.id, name, email, bio ?? null);
      const updated = await findUserById(actor.id);
      return { user: updated };
    } catch (err) {
      request.log.error({ err }, "profile update failed");
      return reply.status(500).send({ message: "Profile update failed" });
    }
  });

  app.patch("/users/me/password", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }

    const schema = z.object({
      currentPassword: z.string().min(1),
      newPassword: z.string().min(3),
      confirmPassword: z.string().min(1),
    });
    const parsed = schema.safeParse(request.body ?? {});
    if (!parsed.success) {
      const issue = parsed.error.errors[0];
      const field = issue?.path?.[0];
      const message = `${field ? `${field}: ` : ""}${issue?.message ?? "Invalid payload"}`;
      return reply.status(400).send({ message });
    }

    const { currentPassword, newPassword, confirmPassword } = parsed.data;

    if (newPassword !== confirmPassword) {
      return reply.status(400).send({ message: "New password and confirmation do not match" });
    }

    // Get current user with password
    const currentUser = await findUserById(actor.id);
    if (!currentUser || !currentUser.password) {
      return reply.status(400).send({ message: "Password change not available for this account" });
    }

    // Verify current password
    const isValidPassword = await bcrypt.compare(currentPassword, currentUser.password);
    if (!isValidPassword) {
      return reply.status(401).send({ message: "Current password is incorrect" });
    }

    try {
      const hashedPassword = await bcrypt.hash(newPassword, 8);
      await updateUserPassword(actor.id, hashedPassword);
      return { message: "Password updated successfully" };
    } catch (err) {
      request.log.error({ err }, "password update failed");
      return reply.status(500).send({ message: "Password update failed" });
    }
  });

  app.get("/users", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const { role, isActive, includeObservers } = request.query as { role?: string; isActive?: string; includeObservers?: string };
    const roleFilter = role ? role.toUpperCase() : null;
    const isActiveFilter = isActive !== undefined ? isActive.toLowerCase() === 'true' : null;
    const shouldIncludeObservers = includeObservers?.toLowerCase() === 'true';

    let baseSql = `
      SELECT id, email, user_name AS "userName", name, avatar_url AS "avatarUrl", bio, role, is_active as "isActive"
      FROM users
      WHERE 1=1
    `;
    const params: any[] = [];
    let paramIndex = 1;

    // Handle isActive filter
    if (isActiveFilter !== null) {
      baseSql += ` AND is_active = $${paramIndex}`;
      params.push(isActiveFilter);
      paramIndex++;
    } else if (roleFilter !== 'NONE' && !shouldIncludeObservers) {
      // Default to active users unless querying for NONE role or including observers
      baseSql += ` AND is_active = TRUE`;
    }

    // Handle role filter
    if (roleFilter) {
      baseSql += ` AND role = $${paramIndex}`;
      params.push(roleFilter);
      paramIndex++;
      
      // Special case: when querying NONE role, also filter by isActive=false
      if (roleFilter === 'NONE' && isActiveFilter === null) {
        baseSql = baseSql.replace('AND is_active = TRUE', 'AND is_active = FALSE');
      }
    } else if (!shouldIncludeObservers) {
      // Default: exclude observers unless includeObservers=true
      baseSql += ` AND role <> 'OBSERVER'`;
    }

    baseSql += ` ORDER BY created_at ASC`;

    const { rows } = await pool.query<User>(baseSql, params);
    return rows;
  });

  app.patch("/users/:id/user-name", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.isActive === false) {
      return reply.status(401).send({ message: "Unauthorized" });
    }
    if (actor.role !== "ADMIN") {
      return reply.status(403).send({ message: "Forbidden" });
    }
    const { id } = request.params as { id: string };
    const schema = z.object({
      userName: z.string().min(2).max(50).regex(/^[a-zA-Z0-9_-]+$/),
    });
    const parsed = schema.safeParse(request.body ?? {});
    if (!parsed.success) {
      const issue = parsed.error.errors[0];
      const field = issue?.path?.[0];
      const message = `${field ? `${field}: ` : ""}${issue?.message ?? "Invalid payload"}`;
      return reply.status(400).send({ message });
    }
    const { userName } = parsed.data;
    const existing = await findUserById(id);
    if (!existing) {
      return reply.status(404).send({ message: "User not found" });
    }
    const userNameExists = await findUserByUserName(userName);
    if (userNameExists && userNameExists.id !== id) {
      return reply.status(409).send({ message: "User name already taken" });
    }
    await updateUserUserName(id, userName);
    const updated = await findUserById(id);
    return { user: updated };
  });

  app.patch("/users/:id/role", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can update roles" });
    }
    const { id } = request.params as { id: string };
    const schema = z.object({
      role: z.enum(["ADMIN", "MANAGER", "BIDDER", "OBSERVER", "NONE"]),
    });
    const body = schema.parse(request.body);
    const existing = await findUserById(id);
    if (!existing) {
      return reply.status(404).send({ message: "User not found" });
    }
    
    // When approving from NONE to OBSERVER, also set isActive=true
    const isApproval = existing.role === "NONE" && body.role === "OBSERVER";
    
    if (isApproval) {
      await pool.query("UPDATE users SET role = $1, is_active = TRUE WHERE id = $2", [
        body.role,
        id,
      ]);
    } else {
      await pool.query("UPDATE users SET role = $1 WHERE id = $2", [
        body.role,
        id,
      ]);
    }
    
    const updated = await findUserById(id);
    if (updated) {
      const roleLabel = updated.role.toLowerCase();
      const message =
        isApproval
          ? `Your account was approved. You can now log in as ${roleLabel}.`
          : existing.role === "OBSERVER" && updated.role !== "OBSERVER"
          ? `Your account was approved as ${roleLabel}.`
          : `Your role was updated to ${roleLabel}.`;
      try {
        await notifyUsers([updated.id], {
          kind: "system",
          message,
          href: "/workspace",
        });
      } catch (err) {
        request.log.error({ err }, "role change notification failed");
      }
    }
    return updated;
  });

  app.patch("/users/:id/ban", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can ban users" });
    }
    const { id } = request.params as { id: string };
    const existing = await findUserById(id);
    if (!existing) {
      return reply.status(404).send({ message: "User not found" });
    }
    await pool.query("UPDATE users SET is_active = FALSE WHERE id = $1", [id]);
    const updated = await findUserById(id);
    if (updated) {
      try {
        await notifyUsers([updated.id], {
          kind: "system",
          message: "Your account has been banned. Please contact an administrator.",
          href: "/auth",
        });
      } catch (err) {
        request.log.error({ err }, "ban notification failed");
      }
    }
    return updated;
  });

  app.patch("/users/:id/approve", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can approve users" });
    }
    const { id } = request.params as { id: string };
    const existing = await findUserById(id);
    if (!existing) {
      return reply.status(404).send({ message: "User not found" });
    }
    await pool.query("UPDATE users SET is_active = TRUE WHERE id = $1", [id]);
    const updated = await findUserById(id);
    if (updated) {
      try {
        await notifyUsers([updated.id], {
          kind: "system",
          message: "Your account has been approved. You can now log in.",
          href: "/workspace",
        });
      } catch (err) {
        request.log.error({ err }, "approve notification failed");
      }
    }
    return updated;
  });

  app.delete("/users/:id", async (request, reply) => {
    const actor = request.authUser;
    if (!actor || actor.role !== "ADMIN") {
      return reply
        .status(403)
        .send({ message: "Only admins can delete users" });
    }
    const { id } = request.params as { id: string };
    const existing = await findUserById(id);
    if (!existing) {
      return reply.status(404).send({ message: "User not found" });
    }
    // Prevent deleting yourself
    if (id === actor.id) {
      return reply.status(400).send({ message: "Cannot delete your own account" });
    }
    try {
      // Use a transaction to ensure all updates happen atomically
      await pool.query('BEGIN');
      
      try {
        // First, set NULL for foreign key references that don't have CASCADE
        await pool.query('UPDATE profiles SET assigned_bidder_id = NULL WHERE assigned_bidder_id = $1', [id]);
        await pool.query('UPDATE profiles SET assigned_by = NULL WHERE assigned_by = $1', [id]);
        await pool.query('UPDATE tasks SET created_by = NULL WHERE created_by = $1', [id]);
        await pool.query('UPDATE tasks SET rejected_by = NULL WHERE rejected_by = $1', [id]);
        await pool.query('UPDATE community_threads SET created_by = NULL WHERE created_by = $1', [id]);
        await pool.query('UPDATE community_messages SET sender_id = NULL WHERE sender_id = $1', [id]);
        await pool.query('UPDATE resume_templates SET created_by = NULL WHERE created_by = $1', [id]);
        await pool.query('UPDATE task_assignment_requests SET requested_by = NULL WHERE requested_by = $1', [id]);
        await pool.query('UPDATE task_assignment_requests SET reviewed_by = NULL WHERE reviewed_by = $1', [id]);
        await pool.query('UPDATE task_done_requests SET requested_by = NULL WHERE requested_by = $1', [id]);
        await pool.query('UPDATE task_done_requests SET reviewed_by = NULL WHERE reviewed_by = $1', [id]);
        await pool.query('UPDATE community_pinned_messages SET pinned_by = NULL WHERE pinned_by = $1', [id]);
        await pool.query('UPDATE daily_reports SET reviewed_by = NULL WHERE reviewed_by = $1', [id]);
        
        // Now delete the user - CASCADE will handle other related records
        const result = await pool.query("DELETE FROM users WHERE id = $1", [id]);
        if (result.rowCount === 0) {
          await pool.query('ROLLBACK');
          return reply.status(404).send({ message: "User not found" });
        }
        
        await pool.query('COMMIT');
        return { success: true, message: "User deleted successfully" };
      } catch (err: any) {
        await pool.query('ROLLBACK');
        throw err;
      }
    } catch (err: any) {
      request.log.error({ err, userId: id }, "Failed to delete user");
      // Check if it's a foreign key constraint violation
      if (err.code === '23503') {
        return reply.status(409).send({ 
          message: "Cannot delete user: user has associated records that prevent deletion. Please remove or reassign related data first." 
        });
      }
      return reply.status(500).send({ message: err.message || "Failed to delete user. Please try again." });
    }
  });

  app.get("/metrics/my", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const bidderUserId = (request.query as { bidderUserId?: string })
      .bidderUserId;
    const userSessions = bidderUserId
      ? sessions.filter((s) => s.bidderUserId === bidderUserId)
      : sessions;
    const tried = userSessions.length;
    const submitted = userSessions.filter(
      (s) => s.status === "SUBMITTED"
    ).length;
    const percentage = tried === 0 ? 0 : Math.round((submitted / tried) * 100);
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    const monthlyApplied = userSessions.filter(
      (s) =>
        s.status === "SUBMITTED" &&
        s.startedAt &&
        new Date(s.startedAt).getTime() >= monthStart.getTime()
    ).length;
    return {
      tried,
      submitted,
      appliedPercentage: percentage,
      monthlyApplied,
      recent: userSessions.slice(0, 5),
    };
  });

  app.get("/settings/llm", async () => llmSettings[0]);
  app.post("/settings/llm", async (request) => {
    const schema = z.object({
      provider: z.enum(["OPENAI", "HUGGINGFACE", "GEMINI"]),
      chatModel: z.string(),
      embedModel: z.string(),
      encryptedApiKey: z.string(),
    });
    const body = schema.parse(request.body);
    const current = llmSettings[0];
    llmSettings[0] = {
      ...current,
      ...body,
      updatedAt: new Date().toISOString(),
    };
    return llmSettings[0];
  });

  app.get("/manager/bidders/summary", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can view bidders" });
    }
    const rows = await listBidderSummaries();
    return rows;
  });

  app.get("/manager/applications", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (
      !actor ||
      (actor.role !== "MANAGER" && actor.role !== "ADMIN" && actor.role !== "BIDDER")
    ) {
      return reply
        .status(403)
        .send({ message: "Only managers, admins, or bidders can view applications" });
    }
    const rows =
      actor.role === "BIDDER" ? await listApplicationsForBidder(actor.id) : await listApplications();
    return rows;
  });

  app.post("/manager/applications", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (
      !actor ||
      (actor.role !== "MANAGER" && actor.role !== "ADMIN" && actor.role !== "BIDDER")
    ) {
      return reply
        .status(403)
        .send({ message: "Only managers, admins, or bidders can create applications" });
    }
    const schema = z.object({
      profileId: z.string().uuid(),
      url: z.string().trim().optional(),
      company: z.string().trim().optional().nullable(),
      resumeId: z.string().uuid().optional().nullable(),
    });
    const parsed = schema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid request body" });
    }
    const { profileId, url, resumeId, company } = parsed.data;
    const profile = await findProfileById(profileId);
    if (!profile) return reply.status(404).send({ message: "Profile not found" });
    if (!profile.assignedBidderId) {
      return reply
        .status(400)
        .send({ message: "Profile must be assigned to a bidder before logging applications" });
    }
    if (actor.role === "BIDDER" && profile.assignedBidderId !== actor.id) {
      return reply.status(403).send({ message: "Cannot log applications for other bidders" });
    }
    const record: ApplicationRecord = {
      id: randomUUID(),
      sessionId: randomUUID(),
      bidderUserId: profile.assignedBidderId,
      profileId,
      resumeId: resumeId ?? null,
      url: url?.trim() || "",
      domain: tryExtractDomain(url ?? ""),
      company: trimToNull(company ?? null),
      createdAt: new Date().toISOString(),
      status: "in_review",
      isReviewed: false,
      reviewedAt: null,
      reviewedBy: null,
    };
    const created = await insertApplicationWithSummary(record);
    return created;
  });

  app.patch("/manager/applications/:id/review", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN")) {
      return reply
        .status(403)
        .send({ message: "Only managers or admins can update applications" });
    }
    const { id } = request.params as { id: string };
    const schema = z.object({ isReviewed: z.boolean() });
    const parsed = schema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid request body" });
    }
    const updated = await updateApplicationReview({
      id,
      isReviewed: parsed.data.isReviewed,
      reviewerId: actor.id,
    });
    if (!updated) {
      return reply.status(404).send({ message: "Application not found" });
    }
    return updated;
  });

  app.patch("/manager/applications/:id/status", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor || (actor.role !== "MANAGER" && actor.role !== "ADMIN" && actor.role !== "BIDDER")) {
      return reply
        .status(403)
        .send({ message: "Only managers, admins, or bidders can update applications" });
    }
    const { id } = request.params as { id: string };
    const schema = z.object({ status: z.enum(["in_review", "accepted", "rejected"]) });
    const parsed = schema.safeParse(request.body);
    if (!parsed.success) {
      return reply.status(400).send({ message: "Invalid request body" });
    }
    const current = await findApplicationById(id);
    if (!current) {
      return reply.status(404).send({ message: "Application not found" });
    }
    if (actor.role === "BIDDER" && current.bidderUserId && current.bidderUserId !== actor.id) {
      return reply.status(403).send({ message: "Cannot update applications for other bidders" });
    }
    const updated = await updateApplicationStatus({ id, status: parsed.data.status });
    return updated;
  });

  // Community: Edit message
  app.patch("/community/messages/:messageId", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });
    const { messageId } = request.params as { messageId: string };
    const schema = z.object({ body: z.string() });
    const body = schema.parse(request.body);
    const text = body.body.trim();
    if (!text)
      return reply.status(400).send({ message: "Message body required" });
    const message = await getMessageById(messageId);
    if (!message)
      return reply.status(404).send({ message: "Message not found" });
    if (message.senderId !== actor.id) {
      return reply.status(403).send({ message: "Can only edit own messages" });
    }
    if (message.isDeleted) {
      return reply.status(400).send({ message: "Cannot edit deleted message" });
    }
    const updated = await editMessage(messageId, text);
    return updated;
  });

  // Community: Delete message
  app.delete("/community/messages/:messageId", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });
    const { messageId } = request.params as { messageId: string };
    const message = await getMessageById(messageId);
    if (!message)
      return reply.status(404).send({ message: "Message not found" });
    const canDelete = message.senderId === actor.id || actor.role === "ADMIN";
    if (!canDelete) {
      return reply.status(403).send({ message: "Permission denied" });
    }
    const deleted = await deleteMessage(messageId);
    return { success: deleted };
  });

  // Community: Add reaction
  app.post(
    "/community/messages/:messageId/reactions",
    async (request, reply) => {
      if (forbidObserver(reply, request.authUser)) return;
      const actor = request.authUser;
      if (!actor) return reply.status(401).send({ message: "Unauthorized" });
      const { messageId } = request.params as { messageId: string };
      const schema = z.object({ emoji: z.string() });
      const body = schema.parse(request.body);
      const message = await getMessageById(messageId);
      if (!message)
        return reply.status(404).send({ message: "Message not found" });
      const reaction = await addMessageReaction({
        id: randomUUID(),
        messageId,
        userId: actor.id,
        emoji: body.emoji,
        createdAt: new Date().toISOString(),
      });
      return reaction;
    }
  );

  // Community: Remove reaction
  app.delete(
    "/community/messages/:messageId/reactions/:emoji",
    async (request, reply) => {
      if (forbidObserver(reply, request.authUser)) return;
      const actor = request.authUser;
      if (!actor) return reply.status(401).send({ message: "Unauthorized" });
      const { messageId, emoji } = request.params as {
        messageId: string;
        emoji: string;
      };
      const removed = await removeMessageReaction(messageId, actor.id, emoji);
      return { success: removed };
    }
  );

  // Community: Pin message
  app.post("/community/messages/:messageId/pin", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });
    const { messageId } = request.params as { messageId: string };
    const message = await getMessageById(messageId);
    if (!message)
      return reply.status(404).send({ message: "Message not found" });
    
    const thread = await findCommunityThreadById(message.threadId);
    if (!thread) return reply.status(404).send({ message: "Thread not found" });
    
    const isMember = await isCommunityThreadMember(message.threadId, actor.id);
    if (!isMember && thread.threadType === "DM") {
      return reply.status(403).send({ message: "Not a member of this thread" });
    }
    
    const pinned = await pinMessage(message.threadId, messageId, actor.id);
    if (!pinned) {
      return reply.status(409).send({ message: "Message already pinned" });
    }
    
    // Broadcast pin event
    const memberIds = await listCommunityThreadMemberIds(message.threadId);
    const allowed = new Set(memberIds);
    communityClients.forEach((c) => {
      if (allowed.has(c.user.id)) {
        sendCommunityPayload(c, {
          type: "message_pinned",
          pinned: {
            threadId: message.threadId,
            message: message
          }
        });
      }
    });
    
    return pinned;
  });

  // Community: Unpin message
  app.delete("/community/messages/:messageId/pin", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });
    const { messageId } = request.params as { messageId: string };
    const message = await getMessageById(messageId);
    if (!message)
      return reply.status(404).send({ message: "Message not found" });
    const unpinned = await unpinMessage(message.threadId, messageId);
    
    // Broadcast unpin event
    if (unpinned) {
      const memberIds = await listCommunityThreadMemberIds(message.threadId);
      const allowed = new Set(memberIds);
      communityClients.forEach((c) => {
        if (allowed.has(c.user.id)) {
          sendCommunityPayload(c, {
            type: "message_unpinned",
            unpinned: {
              threadId: message.threadId,
              messageId
            }
          });
        }
      });
    }
    
    return { success: unpinned };
  });

  // Community: List pinned messages
  app.get("/community/threads/:id/pins", async (request, reply) => {
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });
    const { id } = request.params as { id: string };
    const thread = await ensureCommunityThreadAccess(id, actor, reply);
    if (!thread) return;
    return listPinnedMessages(id);
  });

  // Community: Mark thread as read
  app.post("/community/threads/:id/mark-read", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });
    const { id } = request.params as { id: string };
    const thread = await ensureCommunityThreadAccess(id, actor, reply);
    if (!thread) return;
    await markThreadAsRead(id, actor.id);
    return { success: true };
  });

  // Community: Mark messages as read (bulk)
  app.post("/community/messages/mark-read", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });
    const schema = z.object({
      messageIds: z.array(z.string()),
    });
    const parsed = schema.safeParse(request.body);
    if (!parsed.success) return reply.status(400).send({ message: "Invalid request" });
    const { messageIds } = parsed.data;
    
    await bulkAddReadReceipts(messageIds, actor.id);
    
    // Broadcast read receipts to all clients
    for (const messageId of messageIds) {
      const message = await getMessageById(messageId);
      if (!message) continue;
      
      const memberIds = await listCommunityThreadMemberIds(message.threadId);
      const allowed = new Set(memberIds);
      
      communityClients.forEach((c) => {
        if (allowed.has(c.user.id)) {
          sendCommunityPayload(c, {
            type: 'message_read',
            read: { messageId, userId: actor.id, readAt: new Date().toISOString() },
          });
        }
      });
    }
    
    return { success: true };
  });

  // Community: File upload
  app.post("/community/upload", async (request:any, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });

    const data = await request.file();
    if (!data) return reply.status(400).send({ message: "No file provided" });

    const buffer = await data.toBuffer();
    const fileName = data.filename;
    const mimeType = data.mimetype;

    if (buffer.length > 10 * 1024 * 1024) {
      return reply.status(400).send({ message: "File too large. Max 10MB." });
    }

    const allowedTypes = [
      "image/jpeg",
      "image/png",
      "image/gif",
      "image/webp",
      "application/pdf",
      "application/zip",
      "text/plain",
      "text/csv",
    ];
    if (!allowedTypes.includes(mimeType)) {
      return reply.status(400).send({ message: "File type not supported" });
    }

    try {
      const { url } = await uploadToSupabase(buffer, fileName, mimeType);
      return {
        fileUrl: url,
        fileName,
        fileSize: buffer.length,
        mimeType,
      };
    } catch (err) {
      request.log.error({ err }, "File upload failed");
      console.log(err)
      return reply.status(500).send({ message: "Upload failed" });
    }
  });

  // Community: Get unread summary
  app.get("/community/unread-summary", async (request, reply) => {
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });

    const { rows } = await pool.query<{
      threadId: string;
      threadType: string;
      threadName: string | null;
      unreadCount: number;
      lastReadAt: string;
    }>(
      `
        SELECT 
          u.thread_id AS "threadId",
          t.thread_type AS "threadType",
          t.name AS "threadName",
          u.unread_count AS "unreadCount",
          u.last_read_at AS "lastReadAt"
        FROM community_unread_messages u
        JOIN community_threads t ON t.id = u.thread_id
        WHERE u.user_id = $1 AND u.unread_count > 0
        ORDER BY u.updated_at DESC
      `,
      [actor.id]
    );

    return { unreads: rows };
  });

  // Community: Update user presence
  app.post("/community/presence", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });

    const schema = z.object({
      status: z.enum(["online", "away", "busy", "offline"]),
    });
    const body = schema.parse(request.body);

    await updateUserPresence(actor.id, body.status);
    return { success: true };
  });

  // Community: Get presence for users
  app.get("/community/presence", async (request, reply) => {
    const actor = request.authUser;
    if (!actor) return reply.status(401).send({ message: "Unauthorized" });

    const query = request.query as { userIds?: string };
    const userIds = query.userIds ? query.userIds.split(",") : [];

    if (userIds.length === 0) return { presences: [] };

    const presences = await listUserPresences(userIds);
    return { presences };
  });

  app.ready((err) => {
    if (err) app.log.error(err);
  });

  app.listen({ port: PORT, host: "0.0.0.0" }, (err) => {
    if (err) {
      app.log.error(err);
      process.exit(1);
    }
    app.log.info(`API running on http://localhost:${PORT}`);
  });

  app.get("/ws/community", { websocket: true }, async (socket, req) => {
    const token = readWsToken(req);
    if (!token) {
      socket.close();
      return;
    }
    let user: User | undefined;
    try {
      const decoded = verifyToken(token);
      user = await findUserById(decoded.sub);
    } catch {
      socket.close();
      return;
    }
    if (!user || user.isActive === false || user.role === "OBSERVER") {
      socket.close();
      return;
    }
    const client: CommunityWsClient = { socket, user };
    communityClients.add(client);
    app.log.info({ userId: user.id }, "community ws connected");
    sendCommunityPayload(client, { type: "community_ready" });

    // Handle incoming messages (typing indicators, etc.)
    socket.on("message", async (data: Buffer) => {
      try {
        const payload = JSON.parse(data.toString()) as {
          type?: string;
          threadId?: string;
          [key: string]: any;
        };

        if (payload.type === "typing:start" && payload.threadId) {
          const thread = await findCommunityThreadById(payload.threadId);
          if (!thread) return;

          const memberIds = await listCommunityThreadMemberIds(
            payload.threadId
          );
          const allowed = new Set(memberIds);

          communityClients.forEach((c) => {
            if (c.user.id !== user!.id && allowed.has(c.user.id)) {
              sendCommunityPayload(c, {
                type: "typing",
                typing: {
                  threadId: payload.threadId,
                  userId: user!.id,
                  userName: user!.name,
                  action: "start"
                }
              });
            }
          });
        } else if (payload.type === "typing:stop" && payload.threadId) {
          const thread = await findCommunityThreadById(payload.threadId);
          if (!thread) return;

          const memberIds = await listCommunityThreadMemberIds(
            payload.threadId
          );
          const allowed = new Set(memberIds);

          communityClients.forEach((c) => {
            if (c.user.id !== user!.id && allowed.has(c.user.id)) {
              sendCommunityPayload(c, {
                type: "typing",
                typing: {
                  threadId: payload.threadId,
                  userId: user!.id,
                  userName: user!.name,
                  action: "stop"
                }
              });
            }
          });
        }
      } catch (err) {
        app.log.error({ err }, "websocket message parse error");
      }
    });

    socket.on("close", async () => {
      communityClients.delete(client);
      await updateUserPresence(user!.id, "offline");
      app.log.info({ userId: user.id }, "community ws disconnected");
    });

    // Set user online
    await updateUserPresence(user.id, "online");
  });

  app.get("/ws/notifications", { websocket: true }, async (socket, req) => {
    const token = readWsToken(req);
    if (!token) {
      socket.close();
      return;
    }
    let user: User | undefined;
    try {
      const decoded = verifyToken(token);
      user = await findUserById(decoded.sub);
    } catch {
      socket.close();
      return;
    }
    if (!user || user.isActive === false || user.role === "OBSERVER") {
      socket.close();
      return;
    }
    const client: NotificationWsClient = { socket, user };
    notificationClients.add(client);
    app.log.info({ userId: user.id }, "notifications ws connected");
    sendNotificationPayload(client, { type: "notifications_ready" });

    socket.on("close", async () => {
      notificationClients.delete(client);
      app.log.info({ userId: user!.id }, "notifications ws disconnected");
    });
  });

  app.get(
    "/ws/browser/:sessionId",
    { websocket: true },
    async (socket, req) => {
      // Allow ws without auth for now to keep demo functional
      const { sessionId } = req.params as { sessionId: string };
      const live = livePages.get(sessionId);
      if (!live) {
        socket.send(
          JSON.stringify({ type: "error", message: "No live browser" })
        );
        socket.close();
        return;
      }

      const { page } = live;
      const sendFrame = async () => {
        try {
          const buf = await page.screenshot({ fullPage: true });
          socket.send(
            JSON.stringify({ type: "frame", data: buf.toString("base64") })
          );
        } catch (err) {
          socket.send(
            JSON.stringify({
              type: "error",
              message: "Could not capture frame",
            })
          );
        }
      };

      // Send frames every second
      const interval = setInterval(sendFrame, 1000);
      livePages.set(sessionId, { ...live, interval });

      socket.on("close", () => {
        clearInterval(interval);
        const current = livePages.get(sessionId);
        if (current) {
          livePages.set(sessionId, {
            browser: current.browser,
            page: current.page,
          });
        }
      });
    }
  );
}

function tryExtractDomain(url: string) {
  try {
    const u = new URL(url);
    return u.hostname;
  } catch {
    return undefined;
  }
}

function buildDemoFillPlan(baseInfo: BaseInfo): FillPlanResult {
  const phone = formatPhone(baseInfo?.contact);
  const safeFields = [
    { field: "first_name", value: baseInfo?.name?.first, confidence: 0.98 },
    { field: "last_name", value: baseInfo?.name?.last, confidence: 0.98 },
    { field: "email", value: baseInfo?.contact?.email, confidence: 0.97 },
    {
      field: "phone_code",
      value: baseInfo?.contact?.phoneCode,
      confidence: 0.75,
    },
    {
      field: "phone_number",
      value: baseInfo?.contact?.phoneNumber,
      confidence: 0.78,
    },
    { field: "phone", value: phone, confidence: 0.8 },
    { field: "address", value: baseInfo?.location?.address, confidence: 0.75 },
    { field: "city", value: baseInfo?.location?.city, confidence: 0.75 },
    { field: "state", value: baseInfo?.location?.state, confidence: 0.72 },
    { field: "country", value: baseInfo?.location?.country, confidence: 0.72 },
    {
      field: "postal_code",
      value: baseInfo?.location?.postalCode,
      confidence: 0.72,
    },
    { field: "linkedin", value: baseInfo?.links?.linkedin, confidence: 0.78 },
    { field: "job_title", value: baseInfo?.career?.jobTitle, confidence: 0.7 },
    {
      field: "current_company",
      value: baseInfo?.career?.currentCompany,
      confidence: 0.68,
    },
    { field: "years_exp", value: baseInfo?.career?.yearsExp, confidence: 0.6 },
    {
      field: "desired_salary",
      value: baseInfo?.career?.desiredSalary,
      confidence: 0.62,
    },
    { field: "school", value: baseInfo?.education?.school, confidence: 0.66 },
    { field: "degree", value: baseInfo?.education?.degree, confidence: 0.65 },
    {
      field: "major_field",
      value: baseInfo?.education?.majorField,
      confidence: 0.64,
    },
    {
      field: "graduation_at",
      value: baseInfo?.education?.graduationAt,
      confidence: 0.6,
    },
  ];
  const filled = safeFields
    .filter((f) => Boolean(f.value))
    .map((f) => ({
      field: f.field,
      value: String(f.value ?? ""),
      confidence: f.confidence,
    }));
  return {
    filled,
    suggestions: [],
    blocked: ["EEO", "veteran_status", "disability"],
    actions: [],
  };
}

async function startBrowserSession(session: ApplicationSession) {
  const existing = livePages.get(session.id);
  if (existing) {
    await existing.page.goto(session.url, { waitUntil: "domcontentloaded" });
    await focusFirstField(existing.page);
    return;
  }

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage({
    viewport: { width: 1400, height: 1400 },
  });
  await page.goto(session.url, { waitUntil: "domcontentloaded" });
  await focusFirstField(page);
  livePages.set(session.id, { browser, page });
}

async function stopBrowserSession(sessionId: string) {
  const live = livePages.get(sessionId);
  if (!live) return;
  if (live.interval) clearInterval(live.interval);
  await live.page.close().catch(() => undefined);
  await live.browser.close().catch(() => undefined);
  livePages.delete(sessionId);
}

async function focusFirstField(page: Page) {
  try {
    const locator = page.locator("input, textarea, select").first();
    await locator.scrollIntoViewIfNeeded({ timeout: 4000 });
  } catch {
    // ignore
  }
}

async function broadcastReactionEvent(
  threadId: string,
  messageId: string,
  event: "add" | "remove",
  userId: string,
  emoji: string
) {
  const thread = await findCommunityThreadById(threadId);
  if (!thread) return;

  const payload = {
    type: event === "add" ? "reaction:add" : "reaction:remove",
    threadId,
    messageId,
    userId,
    emoji,
  };

  const memberIds = await listCommunityThreadMemberIds(threadId);
  const allowed = new Set(memberIds);

  communityClients.forEach((client) => {
    if (allowed.has(client.user.id)) {
      sendCommunityPayload(client, payload);
    }
  });
}

async function broadcastMessageEdit(
  threadId: string,
  message: CommunityMessage
) {
  const thread = await findCommunityThreadById(threadId);
  if (!thread) return;

  const payload = {
    type: "message:edited",
    threadId,
    message,
  };

  const memberIds = await listCommunityThreadMemberIds(threadId);
  const allowed = new Set(memberIds);

  communityClients.forEach((client) => {
    if (allowed.has(client.user.id)) {
      sendCommunityPayload(client, payload);
    }
  });
}

async function broadcastMessageDelete(threadId: string, messageId: string) {
  const thread = await findCommunityThreadById(threadId);
  if (!thread) return;

  const payload = {
    type: "message:deleted",
    threadId,
    messageId,
  };

  const memberIds = await listCommunityThreadMemberIds(threadId);
  const allowed = new Set(memberIds);

  communityClients.forEach((client) => {
    if (allowed.has(client.user.id)) {
      sendCommunityPayload(client, payload);
    }
  });
}

bootstrap().catch((err) => {
  app.log.error(err);
  process.exit(1);
});

