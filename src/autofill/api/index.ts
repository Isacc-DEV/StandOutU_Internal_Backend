import type { FastifyInstance } from "fastify";
import { z } from "zod";
import { forbidObserver } from "../../auth";
import {
  AUTOFILL_AI_SYSTEM_PROMPT,
  buildAutofillPrompt,
  type AutofillAiQuestion,
} from "../prompt";

type ChatCompletionParams = {
  provider: "OPENAI" | "HUGGINGFACE" | "GEMINI";
  model: string;
  apiKey: string;
  systemPrompt?: string;
  userPrompt: string;
  temperature?: number;
  maxTokens?: number;
};

type AutofillApiDeps = {
  callChatCompletion: (params: ChatCompletionParams) => Promise<string | undefined>;
  extractJsonArrayPayload: (input: string) => unknown[] | null;
  trimString: (val: unknown) => string;
};

export const registerAutofillApiRoutes = async (
  app: FastifyInstance,
  deps: AutofillApiDeps,
): Promise<void> => {
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
      }),
    ),
    profile: z.record(z.any()),
    apiKey: z.string().optional(),
    model: z.string().optional(),
    debug: z.boolean().optional(),
  });

  app.post("/autofill/ai", async (request, reply) => {
    if (forbidObserver(reply, request.authUser)) return;
    const body = schema.parse(request.body);
    const apiKey = deps.trimString(body.apiKey) || deps.trimString(process.env.OPENAI_API_KEY);
    if (!apiKey) {
      return reply.status(400).send({ message: "OPENAI_API_KEY is required" });
    }
    const model =
      deps.trimString(body.model) ||
      deps.trimString(process.env.OPENAI_AUTOFILL_MODEL) ||
      deps.trimString(process.env.OPENAI_GREENHOUSE_MODEL) ||
      "gpt-4";
    const prompt = buildAutofillPrompt(
      body.questions as AutofillAiQuestion[],
      body.profile ?? {},
    );

    try {
      if (body.debug) {
        request.log.info(
          {
            questionCount: body.questions.length,
            model,
          },
          "Autofill AI request",
        );
        request.log.info({ prompt }, "Autofill AI prompt");
      }
      const content = await deps.callChatCompletion({
        provider: "OPENAI",
        model,
        apiKey,
        systemPrompt: AUTOFILL_AI_SYSTEM_PROMPT,
        userPrompt: prompt,
        temperature: 0.7,
        maxTokens: 2000,
      });
      if (!content) {
        return reply.status(502).send({ message: "LLM response empty" });
      }
      if (body.debug) {
        request.log.info({ content }, "Autofill AI response");
      }
      const parsed = deps.extractJsonArrayPayload(content);
      if (!parsed) {
        return reply.status(502).send({ message: "LLM response not parseable" });
      }
      return {
        answers: parsed,
        provider: "OPENAI",
        model,
        ...(body.debug ? { prompt, rawResponse: content } : {}),
      };
    } catch (err) {
      request.log.error({ err }, "Autofill AI failed");
      return reply.status(502).send({ message: "Autofill AI failed" });
    }
  });
};
