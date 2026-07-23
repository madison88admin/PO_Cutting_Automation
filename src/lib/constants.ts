export const GROQ_API_KEY = process.env.GROQ_API_KEY || '';
export const GEMINI_API_KEY = process.env.GEMINI_API_KEY || '';
export const OLLAMA_BASE_URL = (process.env.OLLAMA_BASE_URL || 'http://127.0.0.1:11434').replace(/\/+$/, '');
export const OLLAMA_MODEL = process.env.OLLAMA_MODEL || 'qwen3:14b';
export const OLLAMA_TIMEOUT_MS = Math.max(1, Number(process.env.OLLAMA_TIMEOUT || '300')) * 1000;
