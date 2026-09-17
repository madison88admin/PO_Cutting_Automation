export const GROQ_API_KEY = process.env.GROQ_API_KEY || '';
export const GEMINI_API_KEY = process.env.GEMINI_API_KEY || '';
export const OLLAMA_BASE_URL = (process.env.OLLAMA_BASE_URL || 'http://5.223.78.194:11434').replace(/\/+$/, '');
export const OLLAMA_MODEL = process.env.OLLAMA_MODEL || 'qwen3:4b';
export const OLLAMA_TIMEOUT_MS = Math.max(1, Number(process.env.OLLAMA_TIMEOUT || '300')) * 1000;
