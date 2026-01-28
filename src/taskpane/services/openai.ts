const OPENAI_API_URL = "https://api.openai.com/v1/chat/completions";

const PATENT_SYSTEM_PROMPT = `You are an expert patent attorney and technical writer assistant. Your role is to help draft, structure, and refine patent applications.

When asked to create a patent structure or draft, follow the standard patent application format:
1. TITLE OF THE INVENTION
2. CROSS-REFERENCE TO RELATED APPLICATIONS (if applicable)
3. FIELD OF THE INVENTION
4. BACKGROUND OF THE INVENTION
5. SUMMARY OF THE INVENTION
6. BRIEF DESCRIPTION OF THE DRAWINGS (if applicable)
7. DETAILED DESCRIPTION OF THE PREFERRED EMBODIMENTS
8. CLAIMS (independent and dependent)
9. ABSTRACT

Guidelines:
- Use precise, technical language appropriate for patent documents
- Claims should be clear, specific, and properly formatted with claim numbers
- Include both independent claims and dependent claims where appropriate
- The description should enable a person skilled in the art to reproduce the invention
- Avoid vague or ambiguous terms
- Use consistent terminology throughout

IMPORTANT: You must ALWAYS respond with a valid JSON object in the following format:
{
  "action": "none" | "insert" | "replace",
  "content": "The text content to insert or replace in the document (if action is insert or replace)",
  "message": "A conversational message to display to the user explaining what you did or answering their question"
}

Action types:
- "none": Use when answering questions, providing explanations, or when no document modification is needed. Only include "message".
- "insert": Use when the user asks you to draft, create, or add new content to the document. Include both "content" and "message".
- "replace": Use when the user has selected text and asks you to modify, improve, rewrite, or replace it. Include both "content" (the replacement text) and "message".

Examples:
- User asks "What are the key parts of a patent?" -> action: "none", message: explains the parts
- User says "Draft a title for an invention about..." -> action: "insert", content: the title, message: "I've drafted a title for your invention."
- User selects text and says "Make this more formal" -> action: "replace", content: improved text, message: "I've made the text more formal."

Always respond with valid JSON only. Do not include any text outside the JSON object.`;

export interface Message {
  role: "user" | "assistant" | "system";
  content: string;
}

export interface StructuredResponse {
  action: "none" | "insert" | "replace";
  content?: string;
  message: string;
}

export interface OpenAIResponse {
  structured: StructuredResponse | null;
  rawContent: string;
  error?: string;
}

export interface FileContext {
  name: string;
  content: string;
}

function parseStructuredResponse(content: string): StructuredResponse | null {
  try {
    const parsed = JSON.parse(content);
    if (parsed.action && parsed.message) {
      return {
        action: parsed.action,
        content: parsed.content,
        message: parsed.message,
      };
    }
  } catch {
    const jsonMatch = content.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      try {
        const parsed = JSON.parse(jsonMatch[0]);
        if (parsed.action && parsed.message) {
          return {
            action: parsed.action,
            content: parsed.content,
            message: parsed.message,
          };
        }
      } catch {
        // Parsing failed
      }
    }
  }
  return null;
}

export async function sendToOpenAI(
  apiKey: string,
  messages: Message[],
  selectedText?: string,
  files?: FileContext[],
  model: string = "gpt-4o"
): Promise<OpenAIResponse> {
  const messagesWithContext = [...messages];
  if (messagesWithContext.length > 0) {
    const lastMessage = messagesWithContext[messagesWithContext.length - 1];
    if (lastMessage.role === "user") {
      let contextPrefix = "";

      if (files && files.length > 0) {
        contextPrefix += "[REFERENCE FILES]\n";
        files.forEach((file) => {
          contextPrefix += `--- ${file.name} ---\n${file.content}\n\n`;
        });
        contextPrefix += "[END REFERENCE FILES]\n\n";
      }

      if (selectedText) {
        contextPrefix += `[SELECTED TEXT FROM DOCUMENT]\n${selectedText}\n[END SELECTED TEXT]\n\n`;
      }

      if (contextPrefix) {
        messagesWithContext[messagesWithContext.length - 1] = {
          ...lastMessage,
          content: `${contextPrefix}User request: ${lastMessage.content}`,
        };
      }
    }
  }

  const allMessages: Message[] = [
    { role: "system", content: PATENT_SYSTEM_PROMPT },
    ...messagesWithContext,
  ];

  try {
    const response = await fetch(OPENAI_API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model,
        messages: allMessages,
        temperature: 0.7,
        max_tokens: 4096,
      }),
    });

    if (!response.ok) {
      const errorData = await response.json();
      return {
        structured: null,
        rawContent: "",
        error: errorData.error?.message || `API error: ${response.status}`,
      };
    }

    const data = await response.json();
    const rawContent = data.choices[0]?.message?.content || "";
    const structured = parseStructuredResponse(rawContent);

    return {
      structured,
      rawContent,
    };
  } catch (error) {
    return {
      structured: null,
      rawContent: "",
      error: error instanceof Error ? error.message : "Unknown error occurred",
    };
  }
}
