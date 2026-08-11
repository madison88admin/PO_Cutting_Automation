import {
    OLLAMA_BASE_URL,
    OLLAMA_MODEL,
    OLLAMA_TIMEOUT_MS,
} from '@/lib/constants';

interface OllamaChatResponse {
    message?: {
        content?: string;
    };
    response?: string;
    error?: string;
}

export async function chatWithOllamaJson(
    systemPrompt: string,
    userPrompt: string
): Promise<string> {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), OLLAMA_TIMEOUT_MS);

    try {
        const response = await fetch(`${OLLAMA_BASE_URL}/api/chat`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                model: OLLAMA_MODEL,
                stream: false,
                format: 'json',
                think: false,
                messages: [
                    { role: 'system', content: systemPrompt },
                    { role: 'user', content: userPrompt },
                ],
                options: {
                    temperature: 0,
                },
            }),
            signal: controller.signal,
        });

        const bodyText = await response.text();
        let data: OllamaChatResponse = {};
        try {
            data = bodyText ? JSON.parse(bodyText) : {};
        } catch {
            if (!response.ok) {
                throw new Error(`Ollama returned HTTP ${response.status}`);
            }
            return bodyText;
        }

        if (!response.ok) {
            throw new Error(data.error || `Ollama returned HTTP ${response.status}`);
        }

        const content = data.message?.content || data.response || '';
        if (!content.trim()) {
            throw new Error('Ollama returned an empty response');
        }
        return content.trim();
    } catch (error) {
        if (error instanceof Error && error.name === 'AbortError') {
            throw new Error(`Ollama request timed out after ${OLLAMA_TIMEOUT_MS / 1000} seconds`);
        }
        throw error;
    } finally {
        clearTimeout(timeout);
    }
}
