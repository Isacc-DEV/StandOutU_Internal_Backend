export type AutofillAiQuestion = {
  id: string;
  type: "text" | "textarea" | "select" | "multi_value_single_select" | "checkbox" | "file";
  label: string;
  required: boolean;
  options?: string[];
};

export const AUTOFILL_AI_SYSTEM_PROMPT =
  "You are answering job applications for a senior software engineer with 10+ years of comprehensive experience across all modern technologies, platforms, and domains. " +
  "The candidate is applying through job board platforms (LinkedIn, Indeed, etc.), not through recruiters or referrals. " +
  "Always answer with confidence and expertise. " +
  'For SELECT questions: return {"answer": "text", "selectedIndex": number}. ' +
  'For CHECKBOX: return {"answer": "text", "selectedIndices": [numbers]}. ' +
  "Return JSON array format.";

export function buildAutofillPrompt(
  questions: AutofillAiQuestion[],
  profile: Record<string, any>
) {
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
5. Agreement/consent to continue hiring process: Always select the affirmative checkbox or "Yes" option.
6. Bot-detection/human verification questions: Answer with a natural human response. If unsure, return answer "[Bot Detection Quiz]" with no indices.
7. Text answers: 40-150 words, assertive tone using "I have", "I've implemented", "I've successfully delivered"
8. NEVER use hedging language: "While I haven't", "Although not directly", "similar to", "transferable skills"

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
