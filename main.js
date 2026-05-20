/**
 * main.js — Electron 메인 프로세스
 *
 * 역할:
 *   Electron 앱의 Node.js 백엔드. 창 생성, 파일 시스템 접근,
 *   AI API 호출, PPTX 파일 처리 등 렌더러(브라우저)가 직접 할 수 없는
 *   작업을 모두 담당한다.
 *
 * 통신 방식:
 *   렌더러(index.html, screen.html)에서 ipcRenderer.invoke/send 로 요청하면
 *   여기서 ipcMain.handle/on 으로 받아 처리 후 결과를 돌려준다.
 *
 * 주요 기능:
 *   1. 설정 파일(settings.json) 읽기/쓰기
 *   2. AI API(Gemini/OpenAI/Claude)로 가사 자동 분리 및 번역
 *   3. 메인 창 + 외부 송출 창 관리
 *   4. PPTX 내보내기 / 가져오기
 */

// ── 외부 라이브러리 ──
const { app, BrowserWindow, ipcMain, Menu, screen, dialog, shell } = require("electron");
const fs   = require("fs");    // 파일 읽기/쓰기
const path = require("path");  // 경로 조합
const { GoogleGenAI } = require("@google/genai");  // Gemini AI SDK
const { DOMParser } = require("@xmldom/xmldom");
const MAX_LINES_PER_SLIDE = 1;

// PPTX 관련 라이브러리는 선택적으로 로드 (없어도 앱 실행 가능)
let AdmZip = null;  // ZIP 파일 파싱 / PPTX 직접 XML 생성
try { AdmZip = require("adm-zip"); } catch(e) { console.warn("adm-zip 미설치:", e.message); }

// ── 자동 업데이트 (electron-updater) ──
// npm install 후 활성화됨. 없으면 조용히 무시.
let autoUpdater = null;
try { ({ autoUpdater } = require("electron-updater")); } catch(e) { console.warn("electron-updater 미설치 — 자동 업데이트 비활성:", e.message); }

/**
 * 자동 업데이트 이벤트를 설정한다.
 * - 개발 모드(electron .)에서는 실행하지 않음 (app.isPackaged === false)
 * - 에러는 console.warn으로만 기록 (앱 동작 방해 없음)
 */
function setupAutoUpdater() {
  if (!autoUpdater) return;
  if (!app.isPackaged) {
    console.log("[updater] 개발 모드 — 자동 업데이트 건너뜀");
    return;
  }

  autoUpdater.autoDownload = true;         // 업데이트 발견 시 자동 다운로드
  autoUpdater.autoInstallOnAppQuit = true; // 앱 종료 시 자동 설치

  autoUpdater.on("update-available", (info) => {
    console.log(`[updater] 새 버전 발견: ${info.version}`);
  });

  autoUpdater.on("update-downloaded", (info) => {
    console.log(`[updater] 다운로드 완료: ${info.version}`);
    // 메인 창이 열려 있으면 렌더러에 알림 전송
    if (mainWin && !mainWin.isDestroyed()) {
      mainWin.webContents.send("update-status", {
        type: "downloaded",
        version: info.version,
        message: `버전 ${info.version} 업데이트가 준비됐습니다. 재시작하면 적용됩니다.`
      });
    }
  });

  autoUpdater.on("error", (err) => {
    console.warn("[updater] 업데이트 확인 오류:", err.message || err);
  });

  // 업데이트 확인 (백그라운드, 에러 발생해도 앱 계속 실행)
  autoUpdater.checkForUpdatesAndNotify().catch((err) => {
    console.warn("[updater] checkForUpdatesAndNotify 실패:", err.message || err);
  });
}

// ── 전역 창 참조 ──
let mainWin   = null;  // 메인 조작창 (index.html)
let secondWin = null;  // 외부 송출 화면 (screen.html)

// ──────────────────────────────────────────────────────────
// 설정 파일 관리
// 저장 위치: OS별 userData 디렉터리 (예: ~/Library/Application Support/WorshipSlide Pro/settings.json)
// ──────────────────────────────────────────────────────────

/** 설정 파일이 저장되는 절대 경로 */
const settingsPath = path.join(app.getPath("userData"), "settings.json");

/**
 * 설정 파일을 읽어 객체로 반환한다.
 * 파일이 없거나 JSON 파싱 실패 시 빈 객체를 반환한다.
 *
 * @returns {Object} 설정 객체 (aiProvider, models, geminiApiKey, aiApiKeys 등 포함)
 */
function loadSettings() {
  try {
    if (!fs.existsSync(settingsPath)) {
      return {};
    }
    const raw = fs.readFileSync(settingsPath, "utf-8");
    return JSON.parse(raw);
  } catch (error) {
    console.error("설정 불러오기 실패:", error);
    return {};
  }
}

const AI_PROVIDERS = ["gemini", "openai", "claude"];
const DEFAULT_MODELS = {
  gemini: "gemini-2.5-flash",
  openai: "gpt-4o-mini",
  claude: "claude-3-5-haiku-latest"
};

function ensureSettingsShape(raw = {}) {
  const settings = (raw && typeof raw === "object") ? raw : {};
  settings.aiProvider = AI_PROVIDERS.includes(settings.aiProvider) ? settings.aiProvider : "gemini";
  settings.aiApiKeys = (settings.aiApiKeys && typeof settings.aiApiKeys === "object") ? settings.aiApiKeys : {};
  settings.models = (settings.models && typeof settings.models === "object") ? settings.models : {};

  for (const provider of AI_PROVIDERS) {
    settings.models[provider] = String(settings.models[provider] || "").trim();
  }

  // 구버전 단일 aiModel 값이 있으면 현재 제공자 슬롯으로 1회 마이그레이션
  if (settings.aiModel && !settings.models[settings.aiProvider]) {
    settings.models[settings.aiProvider] = String(settings.aiModel).trim();
  }
  delete settings.aiModel;

  return settings;
}

/**
 * 설정 객체를 JSON 파일로 저장한다.
 *
 * @param {Object} settings - 저장할 설정 객체
 * @returns {boolean} 저장 성공 여부
 */
function saveSettings(settings) {
  try {
    fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 2), "utf-8");
    return true;
  } catch (error) {
    console.error("설정 저장 실패:", error);
    return false;
  }
}

// ──────────────────────────────────────────────────────────
// IPC: 설정 조회/저장
// ──────────────────────────────────────────────────────────

/**
 * 현재 설정을 렌더러에 전달한다.
 * 보안상 API 키 원문은 보내지 않고 "저장 여부"만 전달한다.
 *
 * 반환값: { hasGeminiApiKey, aiProvider, aiModel, hasAIKey }
 */
ipcMain.handle("settings:get", async () => {
  const settings  = ensureSettingsShape(loadSettings());
  const aiProvider = settings.aiProvider || "gemini";

  return {
    hasGeminiApiKey: !!settings.geminiApiKey,
    aiProvider,
    aiModel: settings.models?.[aiProvider] || "",
    // 현재 선택된 제공자의 API 키가 저장되어 있는지 여부
    hasAIKey: !!(settings.aiApiKeys && settings.aiApiKeys[aiProvider])
  };
});

/**
 * AI 제공자·모델·API 키를 통합 저장한다.
 *
 * @param {Object} data
 *   data.provider - AI 제공자 ('gemini' | 'openai' | 'claude')
 *   data.model    - 모델명 (빈 문자열이면 제공자별 기본 모델 사용)
 *   data.apiKey   - API 키 (빈 문자열이면 기존 키 유지)
 */
async function saveAISettingsCore(data) {
  try {
    const provider = String(data?.provider || "gemini").trim();
    const model    = String(data?.model    || "").trim();
    const apiKey   = String(data?.apiKey   || "").trim();

    if (!AI_PROVIDERS.includes(provider)) {
      return { success: false, error: "지원하지 않는 AI 제공자입니다." };
    }

    const settings = ensureSettingsShape(loadSettings());
    settings.aiProvider = provider;
    settings.models[provider] = model;

    if (apiKey) {
      settings.aiApiKeys[provider] = apiKey;
      // Gemini는 구버전 호환을 위해 geminiApiKey 필드도 함께 저장
      if (provider === "gemini") settings.geminiApiKey = apiKey;
    }

    const saved = saveSettings(settings);
    if (!saved) return { success: false, error: "AI 설정 파일 저장에 실패했습니다." };

    return { success: true };
  } catch (error) {
    console.error("AI 설정 저장 실패:", error);
    return { success: false, error: "AI 설정 저장 중 오류가 발생했습니다." };
  }
}

ipcMain.handle("settings:set-ai-settings", async (event, data) => {
  return await saveAISettingsCore(data);
});

// ──────────────────────────────────────────────────────────
// AI 가사 분리 / 번역 공통 유틸
// ──────────────────────────────────────────────────────────

/**
 * AI에게 전달할 가사 분리 프롬프트를 생성한다.
 * translateTo에 언어 코드('en', 'ja')를 포함하면 번역도 함께 요청한다.
 *
 * 핵심 규칙:
 *   - 한 슬라이드의 lines 배열은 최대 MAX_LINES_PER_SLIDE개
 *   - 원본 가사를 절대 생략/수정/재창작하지 않음
 *   - 후렴/절 표기, 제목 등도 포함
 *   - 번역 요청 시 translations.en, translations.ja 배열로 반환
 *
 * @param {string}   lyrics      - 원본 가사 텍스트
 * @param {string[]} translateTo - 번역 대상 언어 코드 배열 (예: ['en', 'ja'])
 * @returns {string} 완성된 프롬프트 문자열
 */
function getLyricsSplitPrompt(lyrics, translateTo = []) {
  const translationTargets   = Array.isArray(translateTo) ? translateTo : [];
  const translationLanguages = translationTargets
    .map(lang => lang === "en" ? "영어(en)" : lang === "ja" ? "일본어(ja)" : lang)
    .join(", ");

  let prompt = `너는 예배용 슬라이드 송출 앱의 가사 편집 도우미야.

아래 가사를 예배 송출 화면에 넣기 좋게 슬라이드 단위로 나눠줘.

규칙:
- 반드시 JSON 배열만 반환해.
- 설명문, 마크다운 코드블록, 주석 없이 JSON 배열만 반환한다.
- 입력된 내용을 빠뜨리지 말고 순서를 유지해서 JSON 배열로 반환한다.
- 각 슬라이드의 lines 배열에는 원소를 반드시 1개만 넣어라. 절대 2개 이상 넣지 마라.
- 한 문장(한 줄)이 하나의 슬라이드가 된다.
- 여러 문장을 하나의 슬라이드로 합치지 마라. 반드시 문장마다 별도 슬라이드로 분리해라.
- 의미가 어색하게 끊기지 않게 나눠.
- 빈 줄은 무시해.
- 가사 맨 앞에 <제목> 형태로 꺽쇠(<>)로 감싼 텍스트가 있으면, 그것은 노래 제목이다. 반드시 JSON 배열의 첫 번째 슬라이드로 넣어라. 꺽쇠 기호(<>)는 제거하고 안의 텍스트만 lines에 넣어라.
- 사용자가 붙여넣은 내용 중 제목, 곡명, 절 표시, 후렴, 브릿지, 엔딩, 반복 표시, 괄호 안 문구도 의미 있는 내용이면 절대 생략하지 않는다.
- 사용자가 입력한 문장을 요약하거나 삭제하지 않는다.
- 찬양 가사가 아닌 것처럼 보여도 사용자가 붙여넣은 텍스트는 가능한 한 모두 슬라이드에 포함한다.
- 오타를 임의로 고치거나 가사를 새로 창작하지 않는다.
- 중복으로 보이는 가사도 사용자가 입력했다면 임의로 제거하지 않는다.
- 후렴, 1절, 2절 같은 표기는 가사 내용에서 삭제하지 말고, 사용자가 입력한 순서 그대로 하나의 줄 또는 슬라이드로 포함한다.
`;

  // 번역 요청 시 각 슬라이드 객체에 translations 필드 추가
  if (translationLanguages) {
    prompt += `
번역 추가 규칙:
- 각 슬라이드 객체에 translations 객체를 반드시 포함한다.
- 선택된 언어만 translations 안에 넣는다: ${translationLanguages}
- 영어는 translations.en 배열에 넣는다.
- 일본어는 translations.ja 배열에 넣는다.
- 번역도 원본 lines와 같은 의미를 유지하고 각 배열에 원소를 반드시 1개만 넣어라.
- 반환 형식은 반드시 아래와 같아.

[
  {
    "lines": ["원본 줄 1"],
    "translations": {
      "en": ["English line 1"],
      "ja": ["日本語 1行目"]
    }
  }
]
`;
  } else {
    prompt += `
반환 형식은 반드시 아래와 같아.

[
  {
    "lines": ["슬라이드 내용 1"]
  },
  {
    "lines": ["다음 슬라이드 내용"]
  }
]
`;
  }

  prompt += `
사용자가 붙여넣은 원문:

${lyrics}
`;

  return prompt;
}

/**
 * AI 응답에서 마크다운 코드 블록(```json ... ```)을 제거한다.
 * 일부 AI가 JSON을 코드 블록으로 감싸서 반환하는 경우를 처리한다.
 *
 * @param {string} text - AI 응답 원문
 * @returns {string} 순수 JSON 텍스트
 */
function cleanAIJsonText(text) {
  return String(text || "")
    .replace(/```json/g, "")
    .replace(/```/g, "")
    .trim();
}

function normalizeLines(lines) {
  return (Array.isArray(lines) ? lines : [])
    .map(line => String(line || "").trim().replace(/^<(.+)>$/, "$1"))
    .filter(Boolean)
    .slice(0, MAX_LINES_PER_SLIDE);
}

/**
 * AI 응답 텍스트를 슬라이드 배열로 파싱한다.
 * 잘못된 형식의 항목은 필터링하고 최대 2줄까지만 허용한다.
 *
 * @param {string}   text         - AI 응답 텍스트
 * @param {string}   providerName - 오류 메시지에 표시할 AI 제공자명
 * @param {string[]} translateTo  - 번역 언어 코드 배열
 * @returns {Array}  슬라이드 배열. 각 항목: { lines: string[], translations: {} }
 * @throws {Error} JSON 파싱 실패 또는 빈 결과 시
 */
function parseLyricsSlidesFromAI(text, providerName = "AI", translateTo = []) {
  const cleaned = cleanAIJsonText(text);
  let parsed;

  try {
    parsed = JSON.parse(cleaned);
  } catch (parseError) {
    console.error(providerName + " 응답 JSON 파싱 실패:", cleaned);
    throw new Error(providerName + " 응답을 슬라이드 형식으로 읽지 못했습니다. 다시 시도해 주세요.");
  }

  if (!Array.isArray(parsed)) {
    throw new Error(providerName + " 응답이 배열 형식이 아닙니다.");
  }

  const targets = Array.isArray(translateTo) ? translateTo : [];

  const validSlides = parsed
    .filter(item => item && Array.isArray(item.lines))
    .map(item => {
      const lines = normalizeLines(item.lines);

      const translations = {};

      // 영어 번역 처리
      if (targets.includes("en") && Array.isArray(item.translations?.en)) {
        translations.en = normalizeLines(item.translations.en);
      }

      // 일본어 번역 처리
      if (targets.includes("ja") && Array.isArray(item.translations?.ja)) {
        translations.ja = normalizeLines(item.translations.ja);
      }

      return { lines, translations };
    })
    .filter(item => item.lines.length > 0);

  if (validSlides.length === 0) {
    throw new Error(providerName + "가 사용할 수 있는 슬라이드 결과를 만들지 못했습니다.");
  }

  return validSlides;
}

/**
 * 설정에서 특정 AI 제공자의 API 키를 꺼낸다.
 * aiApiKeys 맵 → geminiApiKey(Gemini 전용 구버전 필드) 순서로 탐색한다.
 *
 * @param {Object} settings  - loadSettings() 결과
 * @param {string} provider  - 'gemini' | 'openai' | 'claude'
 * @returns {string} API 키 문자열 (없으면 빈 문자열)
 */
function getAIKey(settings, provider) {
  if (settings.aiApiKeys && settings.aiApiKeys[provider]) {
    return settings.aiApiKeys[provider];
  }
  if (provider === "gemini") {
    return settings.geminiApiKey || "";
  }
  return "";
}

function getModelForProvider(settings, provider) {
  const normalized = ensureSettingsShape(settings);
  return normalized.models[provider] || DEFAULT_MODELS[provider] || "";
}

// ── AI 제공자별 가사 분리 함수 ──

/**
 * Gemini API로 가사를 슬라이드 배열로 분리한다.
 *
 * @param {string}   lyrics      - 원본 가사
 * @param {Object}   settings    - 설정 객체
 * @param {string[]} translateTo - 번역 대상 언어 코드 배열
 * @returns {Promise<Array>} 슬라이드 배열
 */
async function splitLyricsWithGemini(lyrics, settings = loadSettings(), translateTo = []) {
  const inputLyrics = String(lyrics || "").trim();
  if (!inputLyrics) throw new Error("가사 내용이 비어 있습니다.");

  const geminiApiKey = getAIKey(settings, "gemini");
  if (!geminiApiKey) {
    throw new Error("Gemini API 키가 설정되어 있지 않습니다. 설정에서 API 키를 저장해 주세요.");
  }

  const ai = new GoogleGenAI({ apiKey: geminiApiKey });
  const response = await ai.models.generateContent({
    model: getModelForProvider(settings, "gemini"),
    contents: getLyricsSplitPrompt(inputLyrics, translateTo)
  });

  const text = response.text || "";
  return parseLyricsSlidesFromAI(text, "Gemini", translateTo);
}

/**
 * OpenAI API(ChatGPT)로 가사를 슬라이드 배열로 분리한다.
 *
 * @param {string}   lyrics      - 원본 가사
 * @param {Object}   settings    - 설정 객체
 * @param {string[]} translateTo - 번역 대상 언어 코드 배열
 * @returns {Promise<Array>} 슬라이드 배열
 */
async function splitLyricsWithOpenAI(lyrics, settings = loadSettings(), translateTo = []) {
  const inputLyrics = String(lyrics || "").trim();
  if (!inputLyrics) throw new Error("가사 내용이 비어 있습니다.");

  const apiKey = getAIKey(settings, "openai");
  if (!apiKey) {
    throw new Error("OpenAI API 키가 설정되어 있지 않습니다. 설정에서 GPT / OpenAI를 선택하고 API 키를 저장해 주세요.");
  }

  const model = getModelForProvider(settings, "openai");

  // system 프롬프트로 "JSON만 반환"을 강제해 파싱 실패를 줄인다
  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model,
      temperature: 0.2,  // 낮은 temperature = 창의성 억제, 규칙 준수 향상
      messages: [
        { role: "system", content: "You return only valid JSON. No markdown." },
        { role: "user",   content: getLyricsSplitPrompt(inputLyrics, translateTo) }
      ]
    })
  });

  const data = await response.json();
  if (!response.ok) throw new Error(data?.error?.message || "OpenAI API 호출에 실패했습니다.");

  const text = data?.choices?.[0]?.message?.content || "";
  return parseLyricsSlidesFromAI(text, "OpenAI", translateTo);
}

/**
 * Claude API(Anthropic)로 가사를 슬라이드 배열로 분리한다.
 *
 * @param {string}   lyrics      - 원본 가사
 * @param {Object}   settings    - 설정 객체
 * @param {string[]} translateTo - 번역 대상 언어 코드 배열
 * @returns {Promise<Array>} 슬라이드 배열
 */
async function splitLyricsWithClaude(lyrics, settings = loadSettings(), translateTo = []) {
  const inputLyrics = String(lyrics || "").trim();
  if (!inputLyrics) throw new Error("가사 내용이 비어 있습니다.");

  const apiKey = getAIKey(settings, "claude");
  if (!apiKey) {
    throw new Error("Claude API 키가 설정되어 있지 않습니다. 설정에서 Claude / Anthropic을 선택하고 API 키를 저장해 주세요.");
  }

  const model = getModelForProvider(settings, "claude");

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    body: JSON.stringify({
      model,
      max_tokens: 2048,
      temperature: 0.2,
      messages: [{ role: "user", content: getLyricsSplitPrompt(inputLyrics, translateTo) }]
    })
  });

  const data = await response.json();
  if (!response.ok) throw new Error(data?.error?.message || "Claude API 호출에 실패했습니다.");

  // Claude 응답은 content 배열 형태: [{type:"text", text:"..."}]
  const text = Array.isArray(data?.content)
    ? data.content.map(part => part.text || "").join("\n")
    : "";

  return parseLyricsSlidesFromAI(text, "Claude", translateTo);
}

/**
 * 설정에서 선택된 AI 제공자에게 가사 분리를 요청한다.
 * 제공자 선택 로직을 한 곳에 모아 IPC 핸들러들이 공유한다.
 *
 * @param {string}   lyrics      - 원본 가사
 * @param {Object}   settings    - 설정 객체
 * @param {string[]} translateTo - 번역 대상 언어 코드 배열
 * @returns {Promise<Array>} 슬라이드 배열
 */
async function splitLyricsBySelectedProvider(lyrics, settings, translateTo = []) {
  const provider      = settings.aiProvider || "gemini";
  const safeTranslate = Array.isArray(translateTo) ? translateTo : [];

  if (provider === "gemini")  return await splitLyricsWithGemini(lyrics, settings, safeTranslate);
  if (provider === "openai")  return await splitLyricsWithOpenAI(lyrics, settings, safeTranslate);
  if (provider === "claude")  return await splitLyricsWithClaude(lyrics, settings, safeTranslate);

  throw new Error("지원하지 않는 AI 제공자입니다: " + provider);
}

// IPC 핸들러: 구버전 호환용 (Gemini 고정)
ipcMain.handle("gemini:split-lyrics", async (event, lyrics) => {
  return await splitLyricsWithGemini(lyrics, loadSettings(), []);
});

// IPC 핸들러: 설정된 AI 제공자로 가사 분리 (번역 없음)
ipcMain.handle("ai:split-lyrics", async (event, lyrics) => {
  return await splitLyricsBySelectedProvider(lyrics, loadSettings(), []);
});

// IPC 핸들러: 설정된 AI 제공자로 가사 분리 + 번역
// options.translateTo 에 언어 코드 배열을 넘기면 번역 결과도 포함된다
ipcMain.handle("ai:split-lyrics-with-translations", async (event, lyrics, options) => {
  const translateTo = Array.isArray(options?.translateTo) ? options.translateTo : [];
  return await splitLyricsBySelectedProvider(lyrics, loadSettings(), translateTo);
});

// ──────────────────────────────────────────────────────────
// 다언어 가사 동시 분할
// 한국어/영어/일본어 가사를 동시에 입력받아 같은 슬라이드 수로 분할한다.
// ──────────────────────────────────────────────────────────

/**
 * 다언어 가사를 동시에 분할하는 AI 프롬프트를 생성한다.
 * 모든 언어의 슬라이드 수를 동일하게 맞추도록 지시한다.
 *
 * @param {string}  kr    - 한국어 가사 (필수)
 * @param {string}  en    - 영어 가사 (없으면 genEn으로 생성 여부 결정)
 * @param {string}  ja    - 일본어 가사 (없으면 genJa로 생성 여부 결정)
 * @param {boolean} genEn - 영어 가사가 없을 때 AI가 번역 생성 여부
 * @param {boolean} genJa - 일본어 가사가 없을 때 AI가 번역 생성 여부
 * @returns {string} 완성된 프롬프트
 */
function getMultiLangSplitPrompt(kr, en, ja, genEn=false, genJa=false) {
  const hasEn  = !!(en && en.trim());
  const hasJa  = !!(ja && ja.trim());
  const needEn = hasEn || genEn;
  const needJa = hasJa || genJa;

  let prompt = `너는 예배용 슬라이드 앱의 가사 편집 도우미야.

아래 가사를 슬라이드 단위로 나눠줘.

규칙:
- 반드시 JSON 배열만 반환해. 마크다운 코드블록 없음.
- 각 슬라이드의 배열 원소는 반드시 1개만. 절대 2개 이상 넣지 마라.
- 한 문장(한 줄)이 하나의 슬라이드가 된다. 여러 문장을 합치지 마라.
- 반드시 문장마다 별도 슬라이드로 분리해라.
- 빈 줄은 무시해.
- 가사 맨 앞에 <제목> 형태로 꺽쇠(<>)로 감싼 텍스트가 있으면, 그것은 노래 제목이다. 반드시 JSON 배열의 첫 번째 슬라이드로 넣어라. 꺽쇠 기호(<>)는 제거하고 안의 텍스트만 넣어라.
- 내용을 생략, 요약, 수정하지 마.
- 모든 언어의 슬라이드 수를 반드시 동일하게 맞춰.
${needEn || needJa ? "- 영어/일본어는 한국어에 의미적으로 대응하는 내용으로 맞춰." : ""}

한국어:
${kr}
`;

  if (hasEn)      prompt += `\n영어 (아래 텍스트를 한국어와 같은 수의 슬라이드로 분할):\n${en}\n`;
  else if (genEn) prompt += `\n영어: (한국어에서 자연스럽게 번역해서 생성)\n`;

  if (hasJa)      prompt += `\n일본어 (아래 텍스트를 한국어와 같은 수의 슬라이드로 분할):\n${ja}\n`;
  else if (genJa) prompt += `\n일본어: (한국어에서 자연스럽게 번역해서 생성)\n`;

  const enField = needEn ? `"en": ["영어 줄 1"], ` : "";
  const jaField = needJa ? `"ja": ["日本語 1行目"]`   : "";

  prompt += `
반환 형식 (JSON 배열):
[
  {"kr": ["한국어 줄 1"], ${enField}${jaField}},
  {"kr": ["다음 슬라이드 줄 1"], ${enField}${jaField}}
]
`;

  return prompt;
}

/**
 * 다언어 AI 응답을 슬라이드 배열로 파싱한다.
 *
 * @param {string}  text   - AI 응답 텍스트
 * @param {boolean} hasEn  - 영어 포함 여부
 * @param {boolean} hasJa  - 일본어 포함 여부
 * @param {boolean} genEn  - 영어 생성 여부
 * @param {boolean} genJa  - 일본어 생성 여부
 * @returns {Array} 슬라이드 배열. 각 항목: { kr: string[], en: string[], ja: string[] }
 */
function parseMultiLangSlides(text, hasEn, hasJa, genEn=false, genJa=false) {
  const needEn = hasEn || genEn;
  const needJa = hasJa || genJa;
  hasEn = needEn; hasJa = needJa;

  const cleaned = cleanAIJsonText(text);
  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    throw new Error("AI 응답을 슬라이드 형식으로 읽지 못했습니다. 다시 시도해 주세요.");
  }
  if (!Array.isArray(parsed)) throw new Error("AI 응답이 배열 형식이 아닙니다.");

  return parsed
    .filter(item => item && Array.isArray(item.kr) && item.kr.length)
    .map(item => ({
      kr: normalizeLines(item.kr),
      en: hasEn ? normalizeLines(item.en) : [],
      ja: hasJa ? normalizeLines(item.ja) : []
    }))
    .filter(item => item.kr.length > 0);
}

/**
 * 설정된 AI 제공자로 다언어 가사를 동시 분할한다.
 *
 * @param {string}  kr     - 한국어 가사
 * @param {string}  en     - 영어 가사 (없으면 빈 문자열)
 * @param {string}  ja     - 일본어 가사 (없으면 빈 문자열)
 * @param {Object}  settings - 설정 객체
 * @param {boolean} genEn  - 영어 AI 생성 여부
 * @param {boolean} genJa  - 일본어 AI 생성 여부
 * @returns {Promise<Array>} 다언어 슬라이드 배열
 */
async function splitMultiLangByProvider(kr, en, ja, settings, genEn=false, genJa=false) {
  const provider = settings.aiProvider || "gemini";
  const prompt   = getMultiLangSplitPrompt(kr, en, ja, genEn, genJa);
  const hasEn    = !!(en && en.trim());
  const hasJa    = !!(ja && ja.trim());

  let rawText = "";

  if (provider === "gemini") {
    const geminiApiKey = getAIKey(settings, "gemini");
    if (!geminiApiKey) throw new Error("Gemini API 키가 설정되어 있지 않습니다.");
    const ai   = new GoogleGenAI({ apiKey: geminiApiKey });
    const resp = await ai.models.generateContent({ model: getModelForProvider(settings, "gemini"), contents: prompt });
    rawText = resp.text || "";

  } else if (provider === "openai") {
    const apiKey = getAIKey(settings, "openai");
    if (!apiKey) throw new Error("OpenAI API 키가 설정되어 있지 않습니다.");
    const resp = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${apiKey}` },
      body: JSON.stringify({ model: getModelForProvider(settings, "openai"), temperature: 0.2,
        messages: [{ role: "system", content: "Return only valid JSON array. No markdown." }, { role: "user", content: prompt }] })
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data?.error?.message || "OpenAI API 호출 실패");
    rawText = data?.choices?.[0]?.message?.content || "";

  } else if (provider === "claude") {
    const apiKey = getAIKey(settings, "claude");
    if (!apiKey) throw new Error("Claude API 키가 설정되어 있지 않습니다.");
    const resp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
      body: JSON.stringify({ model: getModelForProvider(settings, "claude"), max_tokens: 4096, temperature: 0.2,
        messages: [{ role: "user", content: prompt }] })
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data?.error?.message || "Claude API 호출 실패");
    rawText = Array.isArray(data?.content) ? data.content.map(p => p.text || "").join("\n") : "";

  } else {
    throw new Error("지원하지 않는 AI 제공자: " + provider);
  }

  return parseMultiLangSlides(rawText, hasEn, hasJa, genEn, genJa);
}

/**
 * IPC 핸들러: 한국어/영어/일본어 가사를 동시에 같은 수로 분할한다.
 * genEn/genJa가 true이면 입력이 없어도 AI가 번역 텍스트를 생성한다.
 */
ipcMain.handle("ai:split-multilang", async (event, kr, en, ja, genEn, genJa) => {
  return await splitMultiLangByProvider(
    String(kr || "").trim(),
    String(en || "").trim(),
    String(ja || "").trim(),
    loadSettings(),
    !!genEn,
    !!genJa
  );
});

// ──────────────────────────────────────────────────────────
// 창 생성 및 관리
// ──────────────────────────────────────────────────────────

/**
 * 메인 조작창(index.html)을 생성한다.
 * 창이 닫히면 외부 송출 화면도 함께 닫힌다.
 */
function createMainWindow() {
  mainWin = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 1100,   // 너무 좁으면 레이아웃이 깨지므로 최소 너비 제한
    minHeight: 700,
    title: "WorshipSlide Pro",
    icon: path.join(__dirname, process.platform === "win32" ? "icon.ico" : "icon.png"),
    backgroundColor: "#080810",  // 로딩 전 검은색 배경 (흰색 깜빡임 방지)
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,  // 렌더러가 Node.js에 직접 접근 불가 (보안)
      nodeIntegration: false   // 렌더러에서 require() 사용 불가 (보안)
    }
  });

  mainWin.loadFile("index.html");

  // 개발 중 DevTools를 열고 싶으면 아래 주석 해제
  // mainWin.webContents.openDevTools();

  // 메인 창이 닫히면 송출 화면도 강제 종료
  mainWin.on("closed", () => {
    if (secondWin && !secondWin.isDestroyed()) {
      secondWin.close();
    }
    mainWin = null;
  });
}

/**
 * IPC 핸들러: 외부 송출 화면(screen.html)을 열거나 이미 열려 있으면 전면으로 가져온다.
 *
 * 동작:
 *   1. 연결된 외부 모니터가 있으면 그 모니터에 전체화면으로 열기
 *   2. 외부 모니터가 없으면 현재 모니터에 전체화면으로 열기
 *   3. 이미 열려 있으면 포커스만 이동
 *
 * 반환값: { success: boolean, alreadyOpen: boolean }
 */
ipcMain.handle("open-second-screen", () => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.setMenuBarVisibility(false);
    secondWin.setFullScreen(true);
    secondWin.focus();
    return { success: true, alreadyOpen: true };
  }

  // 외부 모니터 탐색 (없으면 현재 모니터 사용)
  const displays       = screen.getAllDisplays();
  const primaryDisplay = screen.getPrimaryDisplay();
  const external       = displays.find(display => display.id !== primaryDisplay.id);
  const targetDisplay  = external || primaryDisplay;

  const { x, y, width, height } = targetDisplay.bounds;

  secondWin = new BrowserWindow({
    x, y, width, height,
    title: "WorshipSlide Pro Output",
    backgroundColor: "#000000",
    frame: false,           // 타이틀바 없음 (전체화면용)
    fullscreen: true,
    autoHideMenuBar: true,
    skipTaskbar: true,      // 작업 표시줄에 표시하지 않음
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  secondWin.setMenuBarVisibility(false);
  secondWin.setFullScreen(true);
  secondWin.loadFile("screen.html");

  // 송출 화면이 닫히면 조작창에 이벤트 알림
  secondWin.on("closed", () => {
    secondWin = null;
    if (mainWin && !mainWin.isDestroyed()) {
      mainWin.webContents.send("second-screen-closed");
    }
  });

  return { success: true, alreadyOpen: false };
});

/**
 * IPC 핸들러: 슬라이드 데이터를 송출 화면으로 전달한다.
 * index.html → main.js → screen.html 방향으로 중계한다.
 * (렌더러끼리 직접 통신 불가 → 메인 프로세스가 중계)
 */
ipcMain.on("slide-update", (event, data) => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.webContents.send("slide-update", data);
  }
});

/**
 * IPC 핸들러: 블랙아웃(화면 암전)을 켜거나 끈다.
 * on=true이면 검은 오버레이로 화면 전체를 덮는다.
 */
ipcMain.on("blank-screen", (event, on) => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.webContents.send("blank-screen", on);
  }
});

/**
 * IPC 핸들러: 자동재생 진행 바 상태를 송출 화면으로 전달한다.
 * (현재 송출 화면에서는 진행 바를 표시하지 않으므로 사실상 미사용)
 */
ipcMain.on("progress-bar", (event, data) => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.webContents.send("progress-bar", data);
  }
});

/**
 * IPC 핸들러: 송출 화면에서 누른 키를 조작창으로 전달한다.
 * screen.html → main.js → index.html 방향으로 중계.
 *
 * 지원 액션:
 *   'next'       - 다음 슬라이드
 *   'prev'       - 이전 슬라이드
 *   'toggleAuto' - 자동재생 토글
 *   'stopAuto'   - 자동재생 중지
 *   {type:'key'} - 그 외 키 (사용자 정의 단축키 처리)
 */
ipcMain.on("output-control", (event, action) => {
  if (!mainWin || mainWin.isDestroyed()) return;

  if (action === "next" || action === "prev" || action === "toggleAuto" || action === "stopAuto") {
    mainWin.webContents.send("output-control", action);
    return;
  }

  if (action && typeof action === "object" && action.type === "key") {
    mainWin.webContents.send("output-control", action);
  }
});

/**
 * IPC 핸들러: 위치 마커 선 상태를 송출 화면으로 전달한다.
 */
ipcMain.on("marker-line", (event, data) => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.webContents.send("marker-line", data);
  }
});

/**
 * IPC 핸들러: 외부 송출 화면을 닫는다.
 */
ipcMain.handle("close-second-screen", () => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.close();
  }
  return { success: true };
});

// ──────────────────────────────────────────────────────────
// 앱 메뉴
// ──────────────────────────────────────────────────────────

/**
 * 네이티브 앱 메뉴를 생성하고 등록한다.
 * macOS의 복사/붙여넣기/전체선택 등 표준 기능이 기본적으로 동작하려면
 * 메뉴에 해당 role이 포함되어야 한다.
 */
function buildMenu() {
  const template = [
    {
      label: "WorshipSlide",
      submenu: [
        { label: "버전 정보", role: "about" },
        { type: "separator" },
        { label: "종료", role: "quit" }
      ]
    },
    {
      label: "편집",
      submenu: [
        { label: "실행 취소", role: "undo" },
        { label: "다시 실행", role: "redo" },
        { type: "separator" },
        { label: "잘라내기", role: "cut" },
        { label: "복사",     role: "copy" },
        { label: "붙여넣기", role: "paste" },
        { label: "전체 선택", role: "selectAll" }
      ]
    },
    {
      label: "보기",
      submenu: [
        { label: "새로고침",   role: "reload" },
        { label: "전체화면",   role: "togglefullscreen" },
        { type: "separator" },
        { label: "개발자 도구", role: "toggleDevTools" }
      ]
    }
  ];

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);
}

// ── 앱 시작 ──

/** Electron이 초기화를 완료하면 메뉴와 메인 창을 생성한다. */
app.whenReady().then(() => {
  buildMenu();
  createMainWindow();
  setupAutoUpdater(); // 자동 업데이트 초기화 (패키지 빌드 환경에서만 동작)

  // macOS: 독에서 앱 아이콘 클릭 시 창이 모두 닫혀 있으면 새 창 생성
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow();
    }
  });
});

/**
 * 모든 창이 닫히면 non-macOS 플랫폼에서 앱을 종료한다.
 * macOS에서는 표준 동작에 따라 앱을 계속 실행한 채 Dock에 유지한다.
 */
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

// ──────────────────────────────────────────────────────────
// PPTX 내보내기
// ──────────────────────────────────────────────────────────

/**
 * 앱 내 웹폰트 이름을 PPTX 시스템 폰트 이름으로 변환한다.
 */
function korFontToPptx(fontFamily) {
  const map = {
    "Noto Serif KR":    "나눔명조",
    "Noto Sans KR":     "맑은 고딕",
    "Black Han Sans":   "맑은 고딕",
    "Nanum Gothic":     "나눔고딕",
    "Nanum Myeongjo":   "나눔명조",
    "Nanum Pen Script": "나눔손글씨",
    "serif":            "나눔명조",
    "sans-serif":       "맑은 고딕"
  };
  return map[fontFamily] || fontFamily || "맑은 고딕";
}

/**
 * IPC 핸들러: 슬라이드 그룹을 .pptx 파일로 내보낸다.
 *
 * AdmZip으로 PPTX XML을 직접 생성해 pptxgenjs 없이 동작한다.
 * import 코드(parsePptxSlide)가 읽는 XML 구조를 정확히 역산해 생성하므로
 * 가져오기 → 편집 → 내보내기 round-trip이 정확히 유지된다.
 * showMasterSp="0"으로 테마 마스터 배경이 덮어쓰이지 않도록 방어한다.
 */
ipcMain.handle("pptx:export", async (event, { slides, groupName }) => {
  if (!AdmZip) {
    return { ok: false, message: "adm-zip가 설치되지 않았습니다.\n터미널에서 npm install 을 실행해 주세요." };
  }

  const dlg = await dialog.showSaveDialog(mainWin, {
    title: "PPTX로 내보내기",
    defaultPath: (groupName || "슬라이드") + ".pptx",
    filters: [{ name: "PowerPoint 파일", extensions: ["pptx"] }]
  });
  if (dlg.canceled) return { ok: false, message: "cancelled" };

  try {
    // PowerPoint 표준 와이드스크린 (13.33" × 7.5") = 12192000 × 6858000 EMU
    const SW = 12192000, SH = 6858000;

    // import: sz(hundredths-of-pt) / 100 * 4/3 = px  →  export: px * 75 = sz
    const toSz = px => Math.round((px || 36) * 75);
    const pctToEmuX = pct => Math.round((pct || 0) / 100 * SW);
    const pctToEmuY = pct => Math.round((pct || 0) / 100 * SH);

    const xmlEsc = s => String(s || "")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    const hex6 = c => ((c || "000000").replace("#", "")).toUpperCase().padStart(6, "0").slice(0, 6);

    const zip = new AdmZip();
    const mediaMap = {};
    let mediaIdx = 1;

    function registerMedia(dataUrl) {
      if (!dataUrl || !dataUrl.startsWith("data:")) return null;
      const parts = dataUrl.split(",");
      if (parts.length < 2) return null;
      const mime = (parts[0].match(/data:([^;]+)/) || [])[1] || "image/jpeg";
      const extMap = { "image/jpeg": "jpg", "image/png": "png", "image/gif": "gif", "image/bmp": "bmp", "image/webp": "webp" };
      const ext = extMap[mime] || "jpg";
      const rId = "rId" + (200 + mediaIdx);
      mediaMap[rId] = { name: "image" + mediaIdx + "." + ext, buf: Buffer.from(parts[1], "base64"), mime };
      mediaIdx++;
      return rId;
    }

    function buildSlideXml(sl, slideMediaRids) {
      const sz    = toSz(sl.fontSize || 36);
      const fc    = hex6(sl.fontColor);
      const tf    = xmlEsc(korFontToPptx(sl.fontFamily));
      const align = (sl.layoutType === "titleAndBody") ? "l" : "ctr";
      let spTree  = "";
      let shapeId = 2;

      // 메인 가사 텍스트
      if (sl.lyrics && sl.lyrics.trim()) {
        const lines     = String(sl.lyrics).split("\n");
        const lineCount = Math.max(1, lines.length);
        const lineH_emu = Math.round(sz / 100 * 12700 * 1.4);
        const h_emu     = Math.min(lineH_emu * lineCount + 457200, Math.round(SH * 0.88));
        const w_emu     = pctToEmuX(sl.textWidth != null ? sl.textWidth : 86);
        const cx_emu    = pctToEmuX(sl.textX != null ? sl.textX : 50);
        const cy_emu    = pctToEmuY(sl.textY != null ? sl.textY : 50);
        const x_emu     = Math.max(0, cx_emu - w_emu / 2);
        const y_emu     = Math.max(0, Math.min(cy_emu - h_emu / 2, SH - h_emu));

        const paras = lines.map(line =>
          `<a:p><a:pPr algn="${align}"/><a:r>` +
          `<a:rPr lang="ko-KR" sz="${sz}" b="1" dirty="0">` +
          `<a:solidFill><a:srgbClr val="${fc}"/></a:solidFill>` +
          `<a:latin typeface="${tf}"/></a:rPr>` +
          `<a:t>${xmlEsc(line)}</a:t></a:r></a:p>`
        ).join("");

        spTree += `<p:sp>
  <p:nvSpPr><p:cNvPr id="${shapeId++}" name="Lyrics"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${x_emu}" y="${y_emu}"/>
    <a:ext cx="${w_emu}" cy="${h_emu}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
  <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>
    ${paras}
  </p:txBody></p:sp>`;

        // 소제목
        if (sl.sub && sl.sub.trim()) {
          const subSz = Math.max(1000, Math.round(sz * 0.38));
          const subH  = 457200;
          const subX  = Math.round(SW * 0.07);
          const subW  = Math.round(SW * 0.86);
          const subY  = Math.min(y_emu + h_emu + 91440, SH - subH);
          spTree += `<p:sp>
  <p:nvSpPr><p:cNvPr id="${shapeId++}" name="Sub"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${subX}" y="${subY}"/>
    <a:ext cx="${subW}" cy="${subH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
  <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>
    <a:p><a:pPr algn="ctr"/><a:r>
      <a:rPr lang="ko-KR" sz="${subSz}" b="0" dirty="0">
        <a:solidFill><a:srgbClr val="AAAAAA"/></a:solidFill>
        <a:latin typeface="${tf}"/></a:rPr>
      <a:t>${xmlEsc(sl.sub)}</a:t></a:r></a:p>
  </p:txBody></p:sp>`;
        }
      }

      // 추가 텍스트 박스
      for (const box of (sl.textBoxes || [])) {
        if (!box.text || !box.text.trim()) continue;
        const bSz    = toSz(box.fontSize || 24);
        const bFc    = hex6(box.fontColor);
        const bTf    = xmlEsc(korFontToPptx(box.fontFamily));
        const bAlign = box.align === "left" ? "l" : box.align === "right" ? "r" : "ctr";
        const bW     = pctToEmuX(box.width  != null ? box.width  : 80);
        const bH     = Math.max(457200, pctToEmuY(box.height != null ? box.height : 10));
        const bCx    = pctToEmuX(box.x != null ? box.x : 50);
        const bCy    = pctToEmuY(box.y != null ? box.y : 72);
        const bX     = Math.max(0, bCx - bW / 2);
        const bY     = Math.max(0, Math.min(bCy - bH / 2, SH - bH));
        const bParas = String(box.text).split("\n").map(line =>
          `<a:p><a:pPr algn="${bAlign}"/><a:r>` +
          `<a:rPr lang="ko-KR" sz="${bSz}" b="0" dirty="0">` +
          `<a:solidFill><a:srgbClr val="${bFc}"/></a:solidFill>` +
          `<a:latin typeface="${bTf}"/></a:rPr>` +
          `<a:t>${xmlEsc(line)}</a:t></a:r></a:p>`
        ).join("");
        spTree += `<p:sp>
  <p:nvSpPr><p:cNvPr id="${shapeId++}" name="${xmlEsc(box.label || "TextBox")}"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${bX}" y="${bY}"/>
    <a:ext cx="${bW}" cy="${bH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
  <p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>
    ${bParas}
  </p:txBody></p:sp>`;
      }

      // 배경 처리: <p:bg> + 배경 shape 이중 방어 (마스터 테마 배경 덮어쓰기 방지)
      const hasBgImg   = sl.mediaURL && sl.mediaType === "image" && sl.mediaURL.startsWith("data:");
      let bgXml        = "";
      let bgShapeXml   = "";

      if (hasBgImg) {
        const rId = registerMedia(sl.mediaURL);
        if (rId) {
          slideMediaRids[rId] = mediaMap[rId].name;
          const scale = (sl.mediaScale != null ? sl.mediaScale : 100) / 100;
          const iW    = Math.round(SW * scale);
          const iH    = Math.round(SH * scale);
          const iX    = Math.round(pctToEmuX(sl.mediaOffX != null ? sl.mediaOffX : 50) - iW / 2);
          const iY    = Math.round(pctToEmuY(sl.mediaOffY != null ? sl.mediaOffY : 50) - iH / 2);
          bgXml       = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
          bgShapeXml  = `<p:pic>
  <p:nvPicPr><p:cNvPr id="${shapeId++}" name="BgImage"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>
  <p:blipFill><a:blip r:embed="${rId}"/>
    <a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr><a:xfrm><a:off x="${iX}" y="${iY}"/>
    <a:ext cx="${iW}" cy="${iH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
</p:pic>`;
        } else {
          bgXml      = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
          bgShapeXml = `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId++}" name="Bg"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SW}" cy="${SH}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`;
        }
      } else {
        const bgClr  = sl.bgColor ? hex6(sl.bgColor) : "000000";
        bgXml        = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bgClr}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
        bgShapeXml   = `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId++}" name="Bg"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SW}" cy="${SH}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${bgClr}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>`;
      }

      // showMasterSp="0": 슬라이드 마스터 도형(기본 테마 배경 포함)을 숨김
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       showMasterSp="0">
  <p:cSld>
    ${bgXml}
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SW}" cy="${SH}"/>
        <a:chOff x="0" y="0"/><a:chExt cx="${SW}" cy="${SH}"/></a:xfrm></p:grpSpPr>
      ${bgShapeXml}
      ${spTree}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
    }

    // 모든 슬라이드 XML 생성
    const slideMediaRidsArr = [];
    const slideXmls = slides.map(sl => {
      const ridsMap = {};
      slideMediaRidsArr.push(ridsMap);
      return buildSlideXml(sl, ridsMap);
    });

    // 미디어 파일 추가
    for (const [, info] of Object.entries(mediaMap)) {
      zip.addFile("ppt/media/" + info.name, info.buf);
    }

    // 슬라이드 파일 + rels
    slides.forEach((sl, i) => {
      zip.addFile("ppt/slides/slide" + (i + 1) + ".xml", Buffer.from(slideXmls[i], "utf8"));
      const ridsMap = slideMediaRidsArr[i];
      let rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`;
      for (const [rId, name] of Object.entries(ridsMap)) {
        rels += `\n  <Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${name}"/>`;
      }
      rels += "\n</Relationships>";
      zip.addFile("ppt/slides/_rels/slide" + (i + 1) + ".xml.rels", Buffer.from(rels, "utf8"));
    });

    // presentation.xml
    const sldIds = slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("\n    ");
    zip.addFile("ppt/presentation.xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>${sldIds}</p:sldIdLst>
  <p:sldSz cx="${SW}" cy="${SH}" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`, "utf8"));

    // presentation.xml.rels
    const presSlideRels = slides.map((_, i) =>
      `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`
    ).join("\n  ");
    zip.addFile("ppt/_rels/presentation.xml.rels", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  ${presSlideRels}
</Relationships>`, "utf8"));

    // 커스텀 테마 (기본 파란 그라데이션 테마 방지)
    zip.addFile("ppt/theme/theme1.xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="WorshipSlide">
  <a:themeElements>
    <a:clrScheme name="WorshipSlide">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="ffffff"/></a:lt1>
      <a:dk2><a:srgbClr val="1f497d"/></a:dk2>
      <a:lt2><a:srgbClr val="eeece1"/></a:lt2>
      <a:accent1><a:srgbClr val="4f81bd"/></a:accent1>
      <a:accent2><a:srgbClr val="c0504d"/></a:accent2>
      <a:accent3><a:srgbClr val="9bbb59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064a2"/></a:accent4>
      <a:accent5><a:srgbClr val="4bacc6"/></a:accent5>
      <a:accent6><a:srgbClr val="f79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000ff"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="WorshipSlide">
      <a:majorFont><a:latin typeface="나눔명조"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="맑은 고딕"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="WorshipSlide">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`, "utf8"));

    // 슬라이드 마스터 (최소 구성, 기본 배경 없음)
    zip.addFile("ppt/slideMasters/slideMaster1.xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
      <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
            accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr><a:defRPr lang="ko-KR" sz="4400"/></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle><a:lvl1pPr><a:defRPr lang="ko-KR" sz="2800"/></a:lvl1pPr></p:bodyStyle>
    <p:otherStyle><a:lvl1pPr><a:defRPr lang="ko-KR" sz="1800"/></a:lvl1pPr></p:otherStyle>
  </p:txStyles>
</p:sldMaster>`, "utf8"));

    zip.addFile("ppt/slideMasters/_rels/slideMaster1.xml.rels", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`, "utf8"));

    // 슬라이드 레이아웃 (blank)
    zip.addFile("ppt/slideLayouts/slideLayout1.xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
      <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`, "utf8"));

    zip.addFile("ppt/slideLayouts/_rels/slideLayout1.xml.rels", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`, "utf8"));

    // _rels/.rels
    zip.addFile("_rels/.rels", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`, "utf8"));

    // [Content_Types].xml
    const mediaDefaults = [...new Set(Object.values(mediaMap).map(m => {
      const ext = m.name.split(".").pop();
      return `<Default Extension="${ext}" ContentType="${m.mime}"/>`;
    }))].join("\n  ");
    const slideOverrides = slides.map((_, i) =>
      `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    ).join("\n  ");
    zip.addFile("[Content_Types].xml", Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${mediaDefaults}
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  ${slideOverrides}
</Types>`, "utf8"));

    zip.writeZip(dlg.filePath);
    return { ok: true, filePath: dlg.filePath, count: slides.length };
  } catch (err) {
    console.error("PPTX 내보내기 실패:", err);
    return { ok: false, message: "내보내기 실패: " + err.message };
  }
});

// ──────────────────────────────────────────────────────────
// PPTX 가져오기 (텍스트/이미지 추출 방식)
// ──────────────────────────────────────────────────────────

/**
 * IPC 핸들러: .pptx 파일을 열어 슬라이드 데이터로 변환한다.
 *
 * PPTX 파일 구조:
 *   .pptx는 ZIP 아카이브로, 내부에 XML 파일들이 있다.
 *   - ppt/slides/slide1.xml, slide2.xml, ... : 슬라이드 내용
 *   - ppt/slides/_rels/slide1.xml.rels       : 이미지 등 외부 파일 참조 관계
 *   - ppt/media/                             : 이미지 파일들
 *   - ppt/theme/theme1.xml                   : 폰트/색상 테마
 *   - ppt/presentation.xml                   : 슬라이드 크기 등 전체 설정
 */
ipcMain.handle("pptx:import", async (event) => {
  if (!AdmZip) {
    return { ok: false, message: "adm-zip가 설치되지 않았습니다.\n터미널에서 npm install 을 실행해 주세요." };
  }

  const result = await dialog.showOpenDialog(mainWin, {
    title: "PPTX 가져오기",
    filters: [
      { name: "PowerPoint 파일", extensions: ["pptx", "ppt"] },
      { name: "모든 파일",       extensions: ["*"] }
    ],
    properties: ["openFile"]
  });
  if (result.canceled || !result.filePaths.length) return { ok: false, message: "cancelled" };

  // 구버전 .ppt 형식은 ZIP 구조가 달라 지원 불가
  const ext = path.extname(result.filePaths[0]).toLowerCase();
  if (ext === ".ppt") {
    return { ok: false, message: ".ppt 형식은 지원하지 않습니다.\nPowerPoint에서 '다른 이름으로 저장 → .pptx'로 변환 후 다시 시도해 주세요." };
  }

  try {
    const zip = new AdmZip(result.filePaths[0]);

    // 슬라이드 크기 읽기 (단위: EMU, 1인치 = 914400 EMU)
    let slideW = 9144000, slideH = 5143500;  // 기본값: 16:9 (10인치 × 5.625인치)
    try {
      const presXml = zip.readAsText("ppt/presentation.xml");
      const cxM = presXml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"/);
      const cyM = presXml.match(/<p:sldSz\b[^>]*\bcy="(\d+)"/);
      if (cxM) slideW = parseInt(cxM[1]);
      if (cyM) slideH = parseInt(cyM[1]);
    } catch(e) {}

    // 슬라이드 XML 파일 목록을 번호순으로 정렬
    const slideEntries = zip.getEntries()
      .map(e => ({ entry: e, norm: e.entryName.replace(/\\/g, "/") }))
      .filter(({ norm }) => /^ppt\/slides\/slide\d+\.xml$/i.test(norm))
      .sort((a, b) => {
        const n = x => { const m = x.norm.match(/slide(\d+)\.xml$/i); return m ? parseInt(m[1]) : 0; };
        return n(a) - n(b);
      });

    if (slideEntries.length === 0) {
      return { ok: false, message: "슬라이드를 찾을 수 없습니다. 올바른 .pptx 파일인지 확인해 주세요." };
    }

    const theme  = extractPptxTheme(zip);
    const slides = [];

    for (const { entry, norm } of slideEntries) {
      const xml = entry.getData().toString("utf8");

      // 관계 파일(.rels): 이미지 참조 ID와 실제 파일 경로 매핑
      const relsPath  = norm.replace(/\/([^/]+)$/, "/_rels/$1.rels");
      const relsEntry = zipGetEntry(zip, relsPath);
      const relsMap   = relsEntry ? parseRelsXml(relsEntry.getData().toString("utf8")) : {};

      const slide = parsePptxSlide(xml, relsMap, zip, slideW, slideH, theme);
      slides.push(slide);
    }

    const fileName = path.basename(result.filePaths[0], ".pptx");
    return { ok: true, slides, fileName };
  } catch (err) {
    console.error("PPTX 가져오기 실패:", err);
    return { ok: false, message: "가져오기 실패: " + err.message };
  }
});

// ──────────────────────────────────────────────────────────
// PPTX 이미지 변환 가져오기
// 각 슬라이드를 실제로 렌더링해 PNG 이미지로 캡처하는 방식.
// 텍스트 추출이 어려운 복잡한 디자인 슬라이드에 적합하다.
// ──────────────────────────────────────────────────────────

/**
 * 슬라이드 데이터를 HTML 문자열로 변환한다.
 * 화면 캡처를 위해 1920×1080 크기의 HTML 페이지를 생성한다.
 *
 * @param {Object} slide - parsePptxSlide()가 반환한 슬라이드 데이터
 * @returns {string} 완성된 HTML 문자열
 */
function generateSlideHtml(slide) {
  const esc = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  let bgCss = 'background:#000000;';
  if (slide.mediaURL) {
    // cover: 슬라이드 전체를 채워 여백 없이 표시
    bgCss = `background-image:url('${slide.mediaURL}');background-size:cover;background-position:center;background-repeat:no-repeat;background-color:#000000;`;
  } else if (slide.bgColor) {
    bgCss = `background:${slide.bgColor};`;
  }

  const si = (slide.textShadowEnabled !== false) ? Math.max(0, Math.min(1, (slide.textShadowIntensity ?? 80) / 100)) : 0;
  const shadowVal  = si > 0 ? `0 2px ${Math.round(48 * si)}px rgba(0,0,0,${(0.9 * si).toFixed(2)})` : 'none';
  const outlineCss = slide.textOutlineEnabled ? `-webkit-text-stroke:${slide.textOutlineWidth || 2}px rgba(0,0,0,0.85);paint-order:stroke fill;` : '';

  const fs = slide.fontSize || 36;
  const tx = slide.textX ?? 50;
  const ty = slide.textY ?? 50;
  const tw = slide.textWidth ? slide.textWidth + '%' : '80%';
  const textAnchor = slide.textAnchor || 'ctr';
  const textTransY =
    textAnchor === 't' ? '0' :
    textAnchor === 'b' ? '-100%' :
    '-50%';

  // 메인 가사 HTML
  let textHtml = '';
  if (slide.lyrics && slide.lyrics.trim()) {
    textHtml = `<div style="position:absolute;left:${tx}%;top:${ty}%;width:${tw};transform:translate(-50%,${textTransY});font-size:${fs}px;color:${slide.fontColor || '#ffffff'};font-family:'${slide.fontFamily || 'Noto Serif KR'}',serif;font-weight:700;text-shadow:${shadowVal};${outlineCss}text-align:center;line-height:1.7;white-space:pre-wrap;word-break:keep-all;">${esc(slide.lyrics)}</div>`;
    if (slide.sub && slide.sub.trim()) {
      const subFs = Math.round(fs * 0.45);
      const lineCount = Math.max(1, String(slide.lyrics).split(/\n|<br\s*\/?>/i).length);
      textHtml += `<div style="position:absolute;left:${tx}%;top:calc(${ty}% + ${Math.round(fs * 1.7 * lineCount)}px);width:80%;transform:translate(-50%,-50%);font-size:${subFs}px;color:rgba(255,255,255,0.5);font-family:'Noto Sans KR',sans-serif;font-weight:300;text-align:center;letter-spacing:1px;white-space:pre-wrap;">${esc(slide.sub)}</div>`;
    }
  }

  // 도형 HTML (shapeBoxes) — 함수 스코프에서 독립적으로 계산
  const shapeBoxesHtml = (slide.shapeBoxes || []).filter(s => s).map(s => {
    const shapeType = String(s.shapeType || 'rect');
    const borderRadius =
      shapeType === 'ellipse' || shapeType === 'arc' ? '50%' :
      shapeType === 'roundRect' ? '18px' :
      '0';
    const fillColor = s.fillColor || 'transparent';
    const lineColor = s.lineColor || 'transparent';
    const lineWidth = Number.isFinite(Number(s.lineWidth)) ? Number(s.lineWidth) : 0;
    return `<div style="position:absolute;left:${s.x || 50}%;top:${s.y || 50}%;width:${s.w || 1}%;height:${s.h || 1}%;transform:translate(-50%,-50%);background:${fillColor};border:${lineWidth}px solid ${lineColor};border-radius:${borderRadius};box-sizing:border-box;"></div>`;
  }).join('');

  // 추가 텍스트박스 HTML (textBoxes) — shapeBoxes와 분리하여 각 박스를 올바르게 렌더링
  const boxesHtml = (slide.textBoxes || []).filter(b => b && b.text).map(b => {
    const bSi     = (b.textShadowEnabled !== false) ? Math.max(0, Math.min(1, (b.textShadowIntensity ?? 80) / 100)) : 0;
    const bShadow = bSi > 0 ? `0 2px ${Math.round(18 * bSi)}px rgba(0,0,0,${(0.85 * bSi).toFixed(2)})` : 'none';
    const bOutline = b.textOutlineEnabled ? `-webkit-text-stroke:${b.textOutlineWidth || 2}px rgba(0,0,0,0.85);paint-order:stroke fill;` : '';
    const bAnc    = b.textAnchor || 'ctr';
    const bTransY = bAnc === 't' ? '0' : bAnc === 'b' ? '-100%' : '-50%';
    const bFs     = Math.round((b.fontSize || 24) * 1.4);
    const bAlign  = b.align || 'center';
    return `<div style="position:absolute;left:${b.x || 50}%;top:${b.y || 50}%;width:${b.width || 42}%;transform:translate(-50%,${bTransY});font-size:${bFs}px;color:${b.fontColor || '#ffffff'};font-family:'${b.fontFamily || 'Noto Sans KR'}',sans-serif;font-weight:600;text-align:${bAlign};text-shadow:${bShadow};${bOutline}white-space:pre-wrap;word-break:keep-all;">${esc(b.text)}</div>`;
  }).join('');

  // 완성된 HTML 페이지 (1920×1080, 구글 웹폰트 포함)
  return `<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@700;900&family=Noto+Sans+KR:wght@300;400;700&family=Black+Han+Sans&family=Nanum+Gothic&family=Nanum+Myeongjo&display=swap" rel="stylesheet">
<style>*{margin:0;padding:0;box-sizing:border-box;}html,body{width:1920px;height:1080px;overflow:hidden;}</style>
</head><body style="width:1920px;height:1080px;position:relative;${bgCss}">${shapeBoxesHtml}${textHtml}${boxesHtml}</body></html>`;
}

/**
 * IPC 핸들러: .pptx의 각 슬라이드를 PNG 이미지로 변환해 가져온다.
 *
 * 동작 원리:
 *   1. 각 슬라이드를 generateSlideHtml()로 HTML 파일로 변환
 *   2. 숨겨진 오프스크린 BrowserWindow(1920×1080)에 HTML 로드
 *   3. webContents.capturePage()로 PNG 캡처
 *   4. base64 data URL로 변환해 반환
 *
 * 진행 상황: 'pptx:imageProgress' 이벤트로 { current, total } 전달
 */
ipcMain.handle("pptx:importAsImages", async (event) => {
  if (!AdmZip) {
    return { ok: false, message: "adm-zip가 설치되지 않았습니다.\n터미널에서 npm install 을 실행해 주세요." };
  }

  const pickResult = await dialog.showOpenDialog(mainWin, {
    title: "PPTX 이미지로 가져오기",
    filters: [
      { name: "PowerPoint 파일", extensions: ["pptx", "ppt"] },
      { name: "모든 파일",       extensions: ["*"] }
    ],
    properties: ["openFile"]
  });
  if (pickResult.canceled || !pickResult.filePaths.length) return { ok: false, message: "cancelled" };

  const filePath = pickResult.filePaths[0];
  if (path.extname(filePath).toLowerCase() === ".ppt") {
    return { ok: false, message: ".ppt 형식은 지원하지 않습니다.\nPowerPoint에서 '다른 이름으로 저장 → .pptx'로 변환 후 다시 시도해 주세요." };
  }

  let offWin = null;
  try {
    const zip = new AdmZip(filePath);

    // 슬라이드 크기 파싱
    let slideW = 9144000, slideH = 5143500;
    try {
      const presXml = zip.readAsText("ppt/presentation.xml");
      const cxM = presXml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"/);
      const cyM = presXml.match(/<p:sldSz\b[^>]*\bcy="(\d+)"/);
      if (cxM) slideW = parseInt(cxM[1]);
      if (cyM) slideH = parseInt(cyM[1]);
    } catch(e) {}

    const slideEntries = zip.getEntries()
      .map(e => ({ entry: e, norm: e.entryName.replace(/\\/g, "/") }))
      .filter(({ norm }) => /^ppt\/slides\/slide\d+\.xml$/i.test(norm))
      .sort((a, b) => {
        const n = x => { const m = x.norm.match(/slide(\d+)\.xml$/i); return m ? parseInt(m[1]) : 0; };
        return n(a) - n(b);
      });

    if (!slideEntries.length) return { ok: false, message: "슬라이드를 찾을 수 없습니다. 올바른 .pptx 파일인지 확인해 주세요." };
    const theme        = extractPptxTheme(zip);
    const os = require("os");
    // 오프스크린 창: 화면에 보이지 않는 1920×1080 창으로 HTML을 렌더링해 캡처
    offWin = new BrowserWindow({
      width: 1920, height: 1080,
      show: false, frame: false,
      webPreferences: { nodeIntegration: false, contextIsolation: true }
    });

    for (let i = 0; i < slideEntries.length; i++) {
      const { entry, norm } = slideEntries[i];
      // 진행률 이벤트 전송
      event.sender.send("pptx:imageProgress", { current: i + 1, total: slideEntries.length });

      // 메모리 절약을 위해 순차적으로 파싱 및 렌더링 수행
      const xml = entry.getData().toString("utf8");
      const relsPath = norm.replace(/\/([^/]+)$/, "/_rels/$1.rels");
      const relsEntry = zipGetEntry(zip, relsPath);
      const relsMap = relsEntry ? parseRelsXml(relsEntry.getData().toString("utf8")) : {};
      const parsedSlide = parsePptxSlide(xml, relsMap, zip, slideW, slideH, theme);

      const html    = generateSlideHtml(parsedSlide);
      const tmpFile = path.join(os.tmpdir(), `ws_slide_${Date.now()}_${i}.html`);
      
      fs.writeFileSync(tmpFile, html, "utf8");

      try {
        await offWin.loadFile(tmpFile);
        await offWin.webContents.executeJavaScript(`
          Promise.race([
            document.fonts.ready,
            new Promise(r => setTimeout(r, 3000))
          ])
        `);
        const nativeImg = await offWin.webContents.capturePage();
        const b64       = nativeImg.toPNG().toString("base64");
        const slideChunk = {
          mediaURL: `data:image/png;base64,${b64}`,
          mediaType: "image",
          lyrics: "", sub: "", bgColor: null, textBoxes: [],
          layoutType: "centerLyrics",
          fontSize:   parsedSlide.fontSize  || 36,
          fontColor:  parsedSlide.fontColor  || "#ffffff",
          fontFamily: parsedSlide.fontFamily || "Noto Serif KR",
          textX: 50, textY: 50, textWidth: null,
          textShadowEnabled: false, textShadowIntensity: 80,
          textOutlineEnabled: false, textOutlineWidth: 2,
        };
        // 대용량 이미지 데이터는 청크 이벤트로 즉시 전달해 렌더러 프리징을 줄인다.
        event.sender.send("pptx:imageChunk", {
          index: i,
          total: slideEntries.length,
          slide: slideChunk
        });
      } finally {
        try { fs.unlinkSync(tmpFile); } catch(e) {}  // 임시 파일 정리
      }
    }

    offWin.destroy();
    offWin = null;

    const fileName = path.basename(filePath, ".pptx");
    return { ok: true, fileName, total: slideEntries.length, streamed: true };
  } catch (err) {
    if (offWin) { try { offWin.destroy(); } catch(e) {} }
    console.error("PPTX 이미지 변환 오류:", err);
    return { ok: false, message: "변환 실패: " + err.message };
  }
});

// ──────────────────────────────────────────────────────────
// PPTX → 네이티브 이미지 렌더링 (원본 디자인 유지 모드)
// PowerPoint(macOS AppleScript)로 PNG를 직접 저장하고,
// 실패하면 LibreOffice/PDF → Python CoreGraphics PNG 변환으로 폴백한다.
// PowerPoint도 LibreOffice도 없으면 설치 안내 메시지를 반환한다.
// ──────────────────────────────────────────────────────────

function removePptxTextShapeXml(slideXml) {
  return String(slideXml || "").replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (full, body) => {
    return body.includes("<p:txBody>") ? "" : full;
  });
}

function createPptxWithoutTextShapes(filePath, tmpDir) {
  if (!AdmZip) throw new Error("adm-zip가 설치되지 않았습니다.");
  const zip = new AdmZip(filePath);
  const slideEntries = zip.getEntries()
    .filter(entry => /^ppt\/slides\/slide\d+\.xml$/i.test(entry.entryName));
  for (const entry of slideEntries) {
    const xml     = entry.getData().toString("utf8");
    const stripped = removePptxTextShapeXml(xml);
    if (stripped === xml) continue;
    const nextData = Buffer.from(stripped, "utf8");
    if (typeof zip.updateFile === "function") {
      zip.updateFile(entry.entryName, nextData);
    } else {
      zip.deleteFile(entry.entryName);
      zip.addFile(entry.entryName, nextData);
    }
  }
  const outPath = path.join(tmpDir, "no_text_shapes.pptx");
  zip.writeZip(outPath);
  return outPath;
}

/**
 * PPTX 슬라이드를 Electron 내부 렌더러로 PNG 이미지 배열로 변환한다.
 * PowerPoint / LibreOffice 없이도 동작하는 폴백 경로.
 */
async function renderPptxSlidesWithElectron(filePath, event) {
  const zip = new AdmZip(filePath);

  let slideW = 9144000, slideH = 5143500;
  try {
    const presXml = zip.readAsText("ppt/presentation.xml");
    const cxM = presXml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"/);
    const cyM = presXml.match(/<p:sldSz\b[^>]*\bcy="(\d+)"/);
    if (cxM) slideW = parseInt(cxM[1]);
    if (cyM) slideH = parseInt(cyM[1]);
  } catch(e) {}

  const slideEntries = zip.getEntries()
    .map(e => ({ entry: e, norm: e.entryName.replace(/\\/g, "/") }))
    .filter(({ norm }) => /^ppt\/slides\/slide\d+\.xml$/i.test(norm))
    .sort((a, b) => {
      const n = x => { const m = x.norm.match(/slide(\d+)\.xml$/i); return m ? parseInt(m[1]) : 0; };
      return n(a) - n(b);
    });

  if (!slideEntries.length) throw new Error("슬라이드를 찾을 수 없습니다.");

  const theme = extractPptxTheme(zip);
  const parsedSlides = [];
  for (const { entry, norm } of slideEntries) {
    const xml      = entry.getData().toString("utf8");
    const relsPath  = norm.replace(/\/([^/]+)$/, "/_rels/$1.rels");
    const relsEntry = zipGetEntry(zip, relsPath);
    const relsMap   = relsEntry ? parseRelsXml(relsEntry.getData().toString("utf8")) : {};
    parsedSlides.push(parsePptxSlide(xml, relsMap, zip, slideW, slideH, theme));
  }

  const os = require("os");
  const offWin = new BrowserWindow({
    width: 1920, height: 1080,
    show: false, frame: false,
    webPreferences: { nodeIntegration: false, contextIsolation: true }
  });

  const resultSlides = [];
  try {
    for (let i = 0; i < parsedSlides.length; i++) {
      event.sender.send("pptx:imageProgress", { current: i + 1, total: parsedSlides.length });
      const html    = generateSlideHtml(parsedSlides[i]);
      const tmpFile = path.join(os.tmpdir(), `ws_slide_${Date.now()}_${i}.html`);
      fs.writeFileSync(tmpFile, html, "utf8");
      try {
        await offWin.loadFile(tmpFile);
        await offWin.webContents.executeJavaScript(
          `Promise.race([document.fonts.ready, new Promise(r => setTimeout(r, 3000))])`
        );
        const nativeImg = await offWin.webContents.capturePage();
        const slide = {
          mediaURL:  `data:image/png;base64,${nativeImg.toPNG().toString("base64")}`,
          mediaType: "image",
          bgColor:   null
        };
        resultSlides.push(slide);
        event.sender.send("pptx:imageChunk", { index: i, total: parsedSlides.length, slide });
      } finally {
        try { fs.unlinkSync(tmpFile); } catch(e) {}
      }
    }
  } finally {
    try { offWin.destroy(); } catch(e) {}
  }

  return resultSlides;
}

ipcMain.handle("shell:open-external", async (event, url) => {
  await shell.openExternal(url);
});

ipcMain.handle("pptx:select-file", async () => {
  if (!AdmZip) {
    return { ok: false, message: "adm-zip가 설치되지 않았습니다.\n터미널에서 npm install 을 실행해 주세요." };
  }

  const pick = await dialog.showOpenDialog(mainWin, {
    title: "PPTX 이미지로 가져오기",
    filters: [{ name: "PowerPoint 파일", extensions: ["pptx"] }],
    properties: ["openFile"]
  });
  if (pick.canceled || !pick.filePaths.length) return { ok: false, message: "cancelled" };

  const filePath = pick.filePaths[0];
  try {
    const zip = new AdmZip(filePath);
    const slideCount = zip.getEntries()
      .map(e => e.entryName.replace(/\\/g, "/"))
      .filter(name => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
      .length;
    if (!slideCount) return { ok: false, message: "슬라이드를 찾을 수 없습니다. 올바른 .pptx 파일인지 확인해 주세요." };
    return {
      ok: true,
      filePath,
      fileName: path.basename(filePath, path.extname(filePath)),
      slideCount
    };
  } catch (err) {
    return { ok: false, message: "PPTX 파일 확인 실패: " + err.message };
  }
});

// PowerPoint 자동화 권한 사전 확인 (macOS 전용)
// 실제 렌더링 전에 빠른 AppleScript 테스트로 권한 여부를 미리 파악한다.
ipcMain.handle("pptx:check-ppt-permission", async () => {
  const isMac = process.platform === "darwin";
  if (!isMac) return { ok: true, platform: "non-mac" };

  const { spawnSync } = require("child_process");

  let pptAppPath = null;
  const candidatePaths = [
    "/Applications/Microsoft PowerPoint.app",
    "/Applications/Microsoft Office/Microsoft PowerPoint.app",
  ];
  for (const p of candidatePaths) {
    if (fs.existsSync(p) && fs.existsSync(path.join(p, "Contents/MacOS/Microsoft PowerPoint"))) {
      pptAppPath = p;
      break;
    }
  }
  if (!pptAppPath) {
    try {
      const mdf = spawnSync("mdfind", ["kMDItemCFBundleIdentifier == 'com.microsoft.Powerpoint'"], { timeout: 5000, encoding: "utf8" });
      const found = (mdf.stdout || "").split("\n").find(p => p.trim().endsWith(".app"));
      if (found && fs.existsSync(path.join(found.trim(), "Contents/MacOS/Microsoft PowerPoint"))) {
        pptAppPath = found.trim();
      }
    } catch(e) { /* mdfind 없음 */ }
  }

  if (!pptAppPath) {
    return { ok: false, notInstalled: true, message: "Microsoft PowerPoint가 설치되어 있지 않습니다." };
  }

  // 최소 AppleScript로 자동화 권한을 유발 또는 확인한다.
  // macOS는 첫 실행 시 사용자에게 허용/거부 다이얼로그를 자동으로 표시한다.
  const testScript = `tell application ${JSON.stringify(pptAppPath)}\n  get name\nend tell`;
  try {
    const result = spawnSync("osascript", ["-e", testScript], { timeout: 30000, encoding: "utf8" });
    if (result.status === 0) return { ok: true };
    const errMsg = (result.stderr || result.stdout || "");
    if (errMsg.includes("-1743") || errMsg.includes("Not authorized") || errMsg.includes("not allowed to send")) {
      return { ok: false, permissionDenied: true, message: "PowerPoint 자동화 권한이 없습니다." };
    }
    // 기타 오류(PowerPoint가 설치됐지만 응답 불가 등)는 일단 통과시켜 본 렌더에서 처리
    return { ok: true };
  } catch(e) {
    return { ok: true }; // 확인 자체 실패 시 본 렌더에서 처리
  }
});

/**
 * IPC 핸들러: PPTX를 원본 디자인 그대로 PNG 이미지로 변환한다.
 *
 * 처리 순서:
 *   1. PowerPoint(AppleScript, macOS 전용)로 슬라이드를 PNG로 직접 저장
 *   2. 실패 시 PowerPoint PDF 또는 LibreOffice PDF → PNG 변환으로 폴백
 *   3. PNG 파일을 base64 data URL 배열로 반환
 *
 * options.removeTextShapes: true이면 텍스트 shape을 제거한 PPTX로 렌더링 (텍스트 편집 병행 시 사용)
 * options.aspectMode: "16:9" | "original"
 * options.imageFit: "cover" | "contain"
 */
/**
 * PNG 파일들을 180도 회전시켜 base64 data URL로 반환한다.
 * Electron의 오프스크린 BrowserWindow + canvas를 이용해 외부 라이브러리 없이 처리한다.
 * @param {string[]} pngPaths - PNG 파일 절대 경로 배열
 * @returns {Promise<string[]>} base64 data URL 배열
 */
async function rotatePngs180(pngPaths) {
  if (!pngPaths.length) return [];
  const rotWin = new BrowserWindow({
    width: 1920, height: 1080, show: false, frame: false,
    webPreferences: { nodeIntegration: false, contextIsolation: true }
  });
  try {
    await rotWin.loadURL("data:text/html,<canvas id=c></canvas>");
    const results = [];
    for (const p of pngPaths) {
      const buf = fs.readFileSync(p);
      const src = "data:image/png;base64," + buf.toString("base64");
      const rotated = await rotWin.webContents.executeJavaScript(`
        new Promise(resolve => {
          const img = new Image();
          img.onload = () => {
            const c = document.getElementById('c');
            c.width = img.width; c.height = img.height;
            const ctx = c.getContext('2d');
            ctx.clearRect(0, 0, img.width, img.height);
            ctx.save();
            ctx.translate(img.width / 2, img.height / 2);
            ctx.rotate(Math.PI);
            ctx.scale(-1, 1);
            ctx.drawImage(img, -img.width / 2, -img.height / 2);
            ctx.restore();
            resolve(c.toDataURL('image/png'));
          };
          img.onerror = () => resolve(${JSON.stringify(src)});
          img.src = ${JSON.stringify(src)};
        })
      `);
      results.push(rotated);
    }
    return results;
  } finally {
    rotWin.destroy();
  }
}

ipcMain.handle("pptx:render-to-images", async (event, options = {}) => {
  let { filePath } = options || {};

  // filePath가 없으면 파일 선택 다이얼로그 표시
  if (!filePath) {
    const pick = await dialog.showOpenDialog(mainWin, {
      title: "PPTX 원본 디자인으로 가져오기",
      filters: [{ name: "PowerPoint 파일", extensions: ["pptx"] }],
      properties: ["openFile"]
    });
    if (pick.canceled || !pick.filePaths.length) return { ok: false, message: "cancelled" };
    filePath = pick.filePaths[0];
  }

  const aspectMode      = options?.aspectMode === "16:9" ? "16:9" : "original";
  const imageFit        = options?.imageFit === "cover" ? "cover" : "contain";
  const removeTextShapes = options?.removeTextShapes === true;
  const os              = require("os");
  const { execSync, spawnSync } = require("child_process");

  const tmpDir = path.join(os.tmpdir(), "pptx_render_" + Date.now());
  fs.mkdirSync(tmpDir, { recursive: true });
  const collectPngFiles = (rootDir) => {
    const found = [];
    const walk = (dir) => {
      for (const name of fs.readdirSync(dir)) {
        const full = path.join(dir, name);
        const st = fs.statSync(full);
        if (st.isDirectory()) walk(full);
        else if (/\.png$/i.test(name)) found.push(full);
      }
    };
    if (fs.existsSync(rootDir)) walk(rootDir);
    return found.sort((a, b) => {
      const na = parseInt(path.basename(a).match(/\d+/)?.[0] || 0);
      const nb = parseInt(path.basename(b).match(/\d+/)?.[0] || 0);
      return na - nb || a.localeCompare(b);
    });
  };

  let renderFilePath = filePath;
  if (removeTextShapes) {
    try {
      renderFilePath = createPptxWithoutTextShapes(filePath, tmpDir);
    } catch(e) {
      console.warn("텍스트 Shape 제거 실패, 원본으로 폴백:", e.message);
      renderFilePath = filePath;
    }
  }

  const pdfPath = path.join(tmpDir, "presentation.pdf");
  let success   = false;
  let pdfSource = null;
  const renderErrors = [];

  // 1단계: PowerPoint AppleScript로 슬라이드를 PNG 파일로 직접 저장 (macOS 전용)
  const isMac = process.platform === "darwin";
  let permissionDenied = false;

  // PowerPoint 설치 경로 탐색 (기본 경로 + Spotlight 검색)
  let pptAppPath = null;
  if (isMac) {
    const candidatePaths = [
      "/Applications/Microsoft PowerPoint.app",
      "/Applications/Microsoft Office/Microsoft PowerPoint.app",
    ];
    for (const p of candidatePaths) {
      if (fs.existsSync(p) && fs.existsSync(path.join(p, "Contents/MacOS/Microsoft PowerPoint"))) {
        pptAppPath = p;
        break;
      }
    }
    if (!pptAppPath) {
      // Spotlight으로 설치 위치 검색
      try {
        const mdf = spawnSync("mdfind", ["kMDItemCFBundleIdentifier == 'com.microsoft.Powerpoint'"], { timeout: 5000, encoding: "utf8" });
        const found = (mdf.stdout || "").split("\n").find(p => p.trim().endsWith(".app"));
        if (found && fs.existsSync(path.join(found.trim(), "Contents/MacOS/Microsoft PowerPoint"))) {
          pptAppPath = found.trim();
        }
      } catch(e) { /* mdfind 없음 */ }
    }
    if (!pptAppPath) {
      renderErrors.push("Microsoft PowerPoint가 설치되어 있지 않습니다.");
    }
  }

  let isPptLaunchable = false;
  if (pptAppPath) {
    const launchCheck = spawnSync("open", ["-b", "com.microsoft.Powerpoint"], { timeout: 15000, encoding: "utf8" });
    isPptLaunchable = launchCheck.status === 0;
    if (!isPptLaunchable) {
      renderErrors.push("PowerPoint 실행 실패: " + ((launchCheck.stderr || launchCheck.stdout || "").trim() || "PowerPoint를 실행하지 못했습니다."));
    }
  }

  if (isPptLaunchable) {
    const directPngDir = path.join(tmpDir, "powerpoint_pngs");
    fs.mkdirSync(directPngDir, { recursive: true });
    const pngDirForScript = directPngDir.endsWith("/") ? directPngDir : directPngDir + "/";

    // open 후 presentation 객체를 직접 받는 대신 active presentation/front document로 접근
    const script = [
      `tell application ${JSON.stringify(pptAppPath)}`,
      `  activate`,
      `  open POSIX file ${JSON.stringify(renderFilePath)}`,
      `  delay 8`,
      `  try`,
      `    set pres to active presentation`,
      `  on error`,
      `    set pres to presentation 1`,
      `  end try`,
      `  try`,
      `    save pres in POSIX file ${JSON.stringify(pngDirForScript)} as save as PNG`,
      `    delay 5`,
      `    close pres saving no`,
      `  on error e number n`,
      `    try`,
      `      close pres saving no`,
      `    end try`,
      `    error e number n`,
      `  end try`,
      `end tell`,
    ].join("\n");
    const scriptFile = path.join(tmpDir, "export.applescript");
    const scriptBin  = path.join(tmpDir, "export.scpt");
    fs.writeFileSync(scriptFile, script, "utf8");
    try {
      execSync(`osacompile -o "${scriptBin}" "${scriptFile}"`, { timeout: 15000 });
      execSync(`osascript "${scriptBin}"`, { timeout: 300000 });
      const directPngFiles = collectPngFiles(directPngDir);
      if (directPngFiles.length > 0) {
        event.sender.send("pptx:imageProgress", { current: 0, total: directPngFiles.length });
        const rotatedUrls = await rotatePngs180(directPngFiles);
        for (let idx = 0; idx < rotatedUrls.length; idx++) {
          event.sender.send("pptx:imageProgress", { current: idx + 1, total: rotatedUrls.length });
          event.sender.send("pptx:imageChunk", {
            index: idx, total: rotatedUrls.length,
            slide: { mediaURL: rotatedUrls[idx], mediaType: "image", bgColor: null }
          });
        }
        fs.rmSync(tmpDir, { recursive: true, force: true });
        return { ok: true, fileName: path.basename(filePath, path.extname(filePath)), total: rotatedUrls.length, source: "powerpoint-png", streamed: true };
      }
      renderErrors.push("PowerPoint PNG 저장 후 파일이 생성되지 않았습니다.");
    } catch(e) {
      const errMsg = e.message || "";
      if (errMsg.includes("-1743") || errMsg.includes("Not authorized") || errMsg.includes("not allowed to send")) {
        permissionDenied = true;
        renderErrors.push("PowerPoint 자동화 권한 없음(-1743): 시스템 설정에서 이 앱의 Automation 권한을 허용해야 합니다.");
      } else {
        renderErrors.push("PowerPoint PNG 저장 실패: " + errMsg.slice(0, 300));
      }
      console.warn("PowerPoint PNG 저장 실패:", errMsg?.slice(0, 200));
    }
  }

  // 2단계: PowerPoint PDF 폴백 (macOS)
  if (!success && isPptLaunchable) {
    const script = [
      `tell application ${JSON.stringify(pptAppPath)}`,
      `  activate`,
      `  open POSIX file ${JSON.stringify(renderFilePath)}`,
      `  delay 8`,
      `  try`,
      `    set pres to active presentation`,
      `  on error`,
      `    set pres to presentation 1`,
      `  end try`,
      `  try`,
      `    save pres in POSIX file ${JSON.stringify(pdfPath)} as (save as PDF)`,
      `  on error e number n`,
      `    try`,
      `      close pres saving no`,
      `    end try`,
      `    error e number n`,
      `  end try`,
      `  delay 2`,
      `  close pres saving no`,
      `end tell`,
    ].join("\n");
    const scriptFile = path.join(tmpDir, "export_pdf.applescript");
    const scriptBin  = path.join(tmpDir, "export_pdf.scpt");
    fs.writeFileSync(scriptFile, script, "utf8");
    try {
      execSync(`osacompile -o "${scriptBin}" "${scriptFile}"`, { timeout: 15000 });
      execSync(`osascript "${scriptBin}"`, { timeout: 300000 });
      if (fs.existsSync(pdfPath) && fs.statSync(pdfPath).size > 0) {
        success = true;
        pdfSource = "powerpoint-pdf";
      } else {
        renderErrors.push("PowerPoint PDF 저장 후 파일이 생성되지 않았습니다.");
      }
    } catch(e) {
      const errMsg = e.message || "";
      if (errMsg.includes("-1743") || errMsg.includes("Not authorized") || errMsg.includes("not allowed to send")) {
        permissionDenied = true;
        renderErrors.push("PowerPoint 자동화 권한 없음(-1743)");
      } else {
        renderErrors.push("PowerPoint PDF 저장 실패: " + errMsg.slice(0, 300));
      }
      console.warn("PowerPoint PDF 저장 실패:", errMsg?.slice(0, 200));
    }
  }

  // 3단계: Windows PowerPoint COM 자동화 (PowerShell, 비동기)
  const isWindows = process.platform === "win32";
  if (!success && isWindows) {
    const winPngDir = path.join(tmpDir, "powerpoint_pngs");
    fs.mkdirSync(winPngDir, { recursive: true });

    const psScript = `
$ErrorActionPreference = "Continue"
$pptPath = ${JSON.stringify(renderFilePath)}
$outDir  = ${JSON.stringify(winPngDir)}

$ppt = $null
$presentation = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    # $true/$false 사용 (Office PIA 없이도 동작)
    $ppt.Visible = $true
    $ppt.DisplayAlerts = 0   # ppAlertsNone

    # 파일 열기 (ReadOnly=False, Untitled=False, WithWindow=True)
    $presentation = $ppt.Presentations.Open($pptPath, $false, $false, $true)

    # 파일이 완전히 로드될 때까지 대기
    Start-Sleep -Seconds 3

    # Protected View(보호된 보기)로 열린 경우 편집 모드로 전환
    try {
        if ($ppt.ProtectedViewWindows.Count -gt 0) {
            $pvw = $ppt.ProtectedViewWindows.Item(1)
            $presentation = $pvw.Edit()
            Start-Sleep -Seconds 2
        }
    } catch {}

    $slideCount = $presentation.Slides.Count
    Write-Host "PAGES:$slideCount"
    [Console]::Out.Flush()

    $sw = $presentation.PageSetup.SlideWidth
    $sh = $presentation.PageSetup.SlideHeight
    if ($sh -gt 0) {
        $exportH = [int]([Math]::Round(1920 * ($sh / $sw)))
    } else {
        $exportH = 1080
    }

    for ($i = 1; $i -le $slideCount; $i++) {
        $outPath = [System.IO.Path]::Combine($outDir, ("slide{0:D4}.png" -f $i))
        $presentation.Slides.Item($i).Export($outPath, "PNG", 1920, $exportH)
        Write-Host "DONE:\${i}:\${slideCount}"
        [Console]::Out.Flush()
    }

    try { $presentation.Close() } catch {}
    try { $ppt.Quit() } catch {}
    Write-Host "SUCCESS"
} catch {
    $errMsg = $_.Exception.Message
    Write-Host "ERROR:$errMsg"
    if ($presentation -ne $null) { try { $presentation.Close() } catch {} }
    if ($ppt -ne $null) { try { $ppt.Quit() } catch {} }
    exit 1
} finally {
    if ($presentation -ne $null) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null } catch {}
        $presentation = $null
    }
    if ($ppt -ne $null) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
        $ppt = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
`.trim();

    const psFile = path.join(tmpDir, "export_ppt.ps1");
    fs.writeFileSync(psFile, psScript, "utf8");
    try {
      // spawnSync 대신 비동기 spawn 사용 → 메인 스레드 블록 방지 + 실시간 진행 이벤트
      await new Promise((resolve, reject) => {
        const { spawn } = require("child_process");
        const child = spawn(
          "powershell",
          ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", psFile],
          { stdio: ["ignore", "pipe", "pipe"] }
        );
        let stdout = "", stderr = "", totalSlides = 0;

        child.stdout.on("data", chunk => {
          const text = chunk.toString("utf8");
          stdout += text;
          for (const raw of text.split("\n")) {
            const line = raw.trim();
            if (line.startsWith("PAGES:")) {
              totalSlides = parseInt(line.slice(6)) || 0;
              event.sender.send("pptx:imageProgress", { current: 0, total: totalSlides });
            } else if (line.startsWith("DONE:")) {
              const parts = line.slice(5).split(":");
              event.sender.send("pptx:imageProgress", {
                current: parseInt(parts[0]) || 0,
                total:   parseInt(parts[1]) || totalSlides
              });
            }
          }
        });
        child.stderr.on("data", chunk => { stderr += chunk.toString("utf8"); });
        child.on("error", reject);
        child.on("close", code => {
          if (code === 0 || stdout.includes("SUCCESS")) resolve({ stdout, stderr });
          else reject(new Error((stderr || stdout || "PowerShell 실패").slice(0, 500)));
        });
      });

      const winPngFiles = collectPngFiles(winPngDir);
      if (winPngFiles.length > 0) {
        // Windows PowerPoint는 정방향으로 내보내므로 회전 불필요
        for (let idx = 0; idx < winPngFiles.length; idx++) {
          event.sender.send("pptx:imageProgress", { current: idx + 1, total: winPngFiles.length });
          event.sender.send("pptx:imageChunk", {
            index: idx, total: winPngFiles.length,
            slide: {
              mediaURL:  `data:image/png;base64,${fs.readFileSync(winPngFiles[idx]).toString("base64")}`,
              mediaType: "image", bgColor: null
            }
          });
        }
        fs.rmSync(tmpDir, { recursive: true, force: true });
        return { ok: true, fileName: path.basename(filePath, path.extname(filePath)), total: winPngFiles.length, source: "powerpoint-png", streamed: true };
      }
      renderErrors.push("Windows PowerPoint PNG 저장 후 파일이 생성되지 않았습니다.");
    } catch(e) {
      const errMsg = (e.message || "").slice(0, 300);
      if (errMsg.includes("0x80040154") || errMsg.includes("retrieving the COM") || errMsg.includes("80080005") || errMsg.includes("PowerPoint.Application")) {
        renderErrors.push("Microsoft PowerPoint가 설치되지 않았거나 접근할 수 없습니다.");
      } else {
        renderErrors.push("Windows PowerPoint COM 실패: " + errMsg);
      }
      console.warn("Windows PowerPoint COM 실패:", errMsg);
    }
  }

  // 4단계: LibreOffice PPTX → PNG 직접 변환 (크로스플랫폼 폴백)
  // PNG 직접 변환을 먼저 시도하고, 실패 시 PDF → PNG 경로로 폴백
  if (!success) {
    const libreOfficePaths = [
      "/Applications/LibreOffice.app/Contents/MacOS/soffice",
      "/usr/local/bin/soffice", "/usr/bin/soffice",
      "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
      "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ];
    for (const soffice of libreOfficePaths) {
      if (!fs.existsSync(soffice)) continue;

      // 1차 시도: PPTX → PNG 직접 변환 (PDF 중간 단계 없음)
      const loPngDir = path.join(tmpDir, "lo_pngs");
      fs.mkdirSync(loPngDir, { recursive: true });
      try {
        await new Promise((resolve, reject) => {
          const { spawn } = require("child_process");
          const child = spawn(soffice, [
            "--headless", "--convert-to", "png",
            "--outdir", loPngDir, renderFilePath
          ], { stdio: ["ignore", "pipe", "pipe"] });
          let stderr = "";
          child.stderr.on("data", d => { stderr += d.toString(); });
          child.on("error", reject);
          child.on("close", code => {
            if (code === 0) resolve();
            else reject(new Error(stderr || "LibreOffice 실패"));
          });
        });
        const loPngFiles = collectPngFiles(loPngDir);
        if (loPngFiles.length > 0) {
          // macOS LibreOffice는 회전 보정 적용, Windows는 정방향이므로 그대로
          const urls = isMac
            ? await rotatePngs180(loPngFiles)
            : loPngFiles.map(f => `data:image/png;base64,${fs.readFileSync(f).toString("base64")}`);
          for (let idx = 0; idx < urls.length; idx++) {
            event.sender.send("pptx:imageProgress", { current: idx + 1, total: urls.length });
            event.sender.send("pptx:imageChunk", {
              index: idx, total: urls.length,
              slide: { mediaURL: urls[idx], mediaType: "image", bgColor: null }
            });
          }
          fs.rmSync(tmpDir, { recursive: true, force: true });
          return { ok: true, fileName: path.basename(filePath, path.extname(filePath)), total: urls.length, source: "libreoffice-png", streamed: true };
        }
      } catch(e) {
        renderErrors.push("LibreOffice PNG 변환 실패: " + (e.message || "").slice(0, 100));
      }

      // 2차 시도: PPTX → PDF (이후 PDF → PNG 단계에서 처리)
      // spawn 사용: 인자 배열 → shell escape 불필요, 한글/공백 경로에서 안전
      try {
        await new Promise((resolve, reject) => {
          const { spawn } = require("child_process");
          const child = spawn(soffice, [
            "--headless", "--convert-to", "pdf",
            "--outdir", tmpDir, renderFilePath
          ], { stdio: ["ignore", "pipe", "pipe"] });
          let stderr = "";
          const timer = setTimeout(() => {
            try { child.kill(); } catch(_) {}
            reject(new Error("LibreOffice PDF 변환 타임아웃(120s)"));
          }, 120000);
          child.stderr.on("data", d => { stderr += d.toString(); });
          child.on("error", err => { clearTimeout(timer); reject(err); });
          child.on("close", code => {
            clearTimeout(timer);
            if (code === 0) resolve();
            else reject(new Error(stderr || "LibreOffice PDF 변환 실패"));
          });
        });
        const pdfs = fs.readdirSync(tmpDir).filter(f => /\.pdf$/i.test(f));
        if (pdfs.length > 0) {
          fs.renameSync(path.join(tmpDir, pdfs[0]), pdfPath);
          success = true;
          pdfSource = "libreoffice-pdf";
          break;
        }
      } catch(e) {
        renderErrors.push("LibreOffice PDF 변환 실패: " + (e.message || "").slice(0, 100));
        console.warn("LibreOffice PDF 실패:", (e.message || "").slice(0, 100));
      }
    }
  }

  if (!success) {
    fs.rmSync(tmpDir, { recursive: true, force: true });
    if (permissionDenied) {
      return { ok: false, permissionDenied: true, message: "PowerPoint 자동화 권한이 없습니다." };
    }
    let hint = "";
    if (isWindows) {
      const pptNotInstalled = renderErrors.some(e => e.includes("설치되지 않"));
      const libreOfficeFailed = renderErrors.some(e => e.includes("LibreOffice"));
      if (!renderErrors.length || (pptNotInstalled && !libreOfficeFailed)) {
        hint = "Microsoft PowerPoint 또는 LibreOffice가 필요합니다.\n\n" +
          "① Microsoft PowerPoint가 설치된 경우: 앱을 처음 실행 시 PowerPoint가 잠시 열립니다.\n" +
          "② LibreOffice를 설치하면 PowerPoint 없이도 변환 가능합니다.\n" +
          "   다운로드: https://www.libreoffice.org/download/";
      } else {
        hint = "PowerPoint 또는 LibreOffice 변환 중 오류가 발생했습니다.\n" +
          "PowerPoint가 실행 중이라면 완전히 닫은 후 다시 시도해 주세요.";
      }
    } else if (!pptAppPath) {
      hint = "Microsoft PowerPoint가 설치되어 있지 않습니다.\nLibreOffice를 설치하면 대체 변환이 가능합니다.";
    } else {
      hint = "PowerPoint 변환에 실패했습니다.\n자동화(Automation) 권한이 허용되어 있는지 확인해 주세요.";
    }
    const detail = renderErrors.length ? "\n\n상세:\n- " + renderErrors.join("\n- ") : "";
    return { ok: false, message: hint + detail };
  }

  // 6단계: PDF → PNG 변환 (플랫폼별)
  const pngDir = path.join(tmpDir, "pngs");
  fs.mkdirSync(pngDir, { recursive: true });
  let pdfPngSuccess = false;

  if (isMac) {
    // macOS: Python + CoreGraphics 프레임워크
    const pythonScript = `
import ctypes, os, sys

cf = ctypes.CDLL('/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation')
cg = ctypes.CDLL('/System/Library/Frameworks/CoreGraphics.framework/CoreGraphics')

class CGRect(ctypes.Structure):
    _fields_ = [('x', ctypes.c_double), ('y', ctypes.c_double),
                ('width', ctypes.c_double), ('height', ctypes.c_double)]

cf.CFStringCreateWithCString.restype = ctypes.c_void_p
cf.CFStringCreateWithCString.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_uint32]
cf.CFURLCreateWithFileSystemPath.restype = ctypes.c_void_p
cf.CFURLCreateWithFileSystemPath.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_long, ctypes.c_bool]
cf.CFRelease.argtypes = [ctypes.c_void_p]
cg.CGPDFDocumentCreateWithURL.restype = ctypes.c_void_p
cg.CGPDFDocumentCreateWithURL.argtypes = [ctypes.c_void_p]
cg.CGPDFDocumentGetNumberOfPages.restype = ctypes.c_size_t
cg.CGPDFDocumentGetNumberOfPages.argtypes = [ctypes.c_void_p]
cg.CGPDFDocumentGetPage.restype = ctypes.c_void_p
cg.CGPDFDocumentGetPage.argtypes = [ctypes.c_void_p, ctypes.c_size_t]
cg.CGPDFPageGetBoxRect.restype = CGRect
cg.CGPDFPageGetBoxRect.argtypes = [ctypes.c_void_p, ctypes.c_int]
cg.CGColorSpaceCreateDeviceRGB.restype = ctypes.c_void_p
cg.CGColorSpaceRelease.argtypes = [ctypes.c_void_p]
cg.CGBitmapContextCreate.restype = ctypes.c_void_p
cg.CGBitmapContextCreate.argtypes = [ctypes.c_void_p, ctypes.c_size_t, ctypes.c_size_t,
    ctypes.c_size_t, ctypes.c_size_t, ctypes.c_void_p, ctypes.c_uint32]
cg.CGContextRelease.argtypes = [ctypes.c_void_p]
cg.CGContextDrawPDFPage.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
cg.CGContextTranslateCTM.argtypes = [ctypes.c_void_p, ctypes.c_double, ctypes.c_double]
cg.CGContextScaleCTM.argtypes = [ctypes.c_void_p, ctypes.c_double, ctypes.c_double]
cg.CGContextSetRGBFillColor.argtypes = [ctypes.c_void_p, ctypes.c_double, ctypes.c_double, ctypes.c_double, ctypes.c_double]
cg.CGContextFillRect.argtypes = [ctypes.c_void_p, CGRect]

def _save_bgra_png(filename, width, height, data):
    import zlib, struct
    def make_chunk(tag, d):
        return struct.pack(">I", len(d)) + tag + d + struct.pack(">I", zlib.crc32(tag + d) & 0xffffffff)
    sig = b'\\x89PNG\\r\\n\\x1a\\n'
    ihdr = make_chunk(b'IHDR', struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0))
    mv = memoryview(data)
    rgba = bytearray(width * height * 4)
    for i in range(0, len(mv), 4):
        b, g, r, a = mv[i:i+4]
        rgba[i] = r; rgba[i+1] = g; rgba[i+2] = b; rgba[i+3] = a
    scanlines = bytearray()
    for y in range(height):
        scanlines.append(0)
        scanlines.extend(rgba[y * width * 4 : (y + 1) * width * 4])
    idat = make_chunk(b'IDAT', zlib.compress(scanlines))
    iend = make_chunk(b'IEND', b'')
    with open(filename, 'wb') as f:
        f.write(sig + ihdr + idat + iend)

pdf_path, out_dir = sys.argv[1], sys.argv[2]
target_w, target_h = 1920, 1080
aspect_mode = ${JSON.stringify(aspectMode)}
image_fit = ${JSON.stringify(imageFit)}

path_str = cf.CFStringCreateWithCString(None, pdf_path.encode('utf-8'), 134217984)
url = cf.CFURLCreateWithFileSystemPath(None, path_str, 0, False)
cf.CFRelease(path_str)
doc = cg.CGPDFDocumentCreateWithURL(url)
cf.CFRelease(url)
if not doc:
    print("ERROR:PDF open failed"); sys.exit(1)

page_count = cg.CGPDFDocumentGetNumberOfPages(doc)
print(f"PAGES:{page_count}", flush=True)

for i in range(1, page_count + 1):
    page = cg.CGPDFDocumentGetPage(doc, i)
    if not page: continue
    box = cg.CGPDFPageGetBoxRect(page, 0)
    if aspect_mode == "16:9":
        rw, rh = target_w, target_h
        scale = max(target_w / box.width, target_h / box.height) if image_fit == "cover" else min(target_w / box.width, target_h / box.height)
        dx = (target_w - box.width * scale) / 2.0
        dy = (target_h - box.height * scale) / 2.0
    else:
        scale = min(target_w / box.width, target_h / box.height)
        rw, rh = int(box.width * scale), int(box.height * scale)
        dx, dy = 0, 0
    colorspace = cg.CGColorSpaceCreateDeviceRGB()
    bpr = rw * 4
    buf = ctypes.create_string_buffer(bpr * rh)
    ctx = cg.CGBitmapContextCreate(buf, rw, rh, 8, bpr, colorspace, ctypes.c_uint32(0x2002))
    cg.CGColorSpaceRelease(colorspace)
    if not ctx: continue
    cg.CGContextSetRGBFillColor(ctx, 1.0, 1.0, 1.0, 1.0)
    cg.CGContextFillRect(ctx, CGRect(0, 0, rw, rh))
    cg.CGContextTranslateCTM(ctx, dx, rh - dy)
    cg.CGContextScaleCTM(ctx, scale, -scale)
    cg.CGContextTranslateCTM(ctx, -box.x, -box.y)
    cg.CGContextDrawPDFPage(ctx, page)
    cg.CGContextRelease(ctx)
    raw = ctypes.string_at(buf, bpr * rh)
    out_path = os.path.join(out_dir, f'slide{i:04d}.png')
    _save_bgra_png(out_path, rw, rh, raw)
    print(f"DONE:{i}:{page_count}", flush=True)
`;
    const pyFile = path.join(tmpDir, "pdf2png.py");
    fs.writeFileSync(pyFile, pythonScript, "utf8");
    try {
      await new Promise((resolve, reject) => {
        const { spawn } = require("child_process");
        const child = spawn("python3", [pyFile, pdfPath, pngDir], { stdio: ["ignore", "pipe", "pipe"] });
        let stdout = "", stderr = "";
        child.stdout.on("data", chunk => {
          const text = chunk.toString("utf8");
          stdout += text;
          for (const raw of text.split("\n")) {
            const line = raw.trim();
            if (line.startsWith("DONE:")) {
              const [, cur, tot] = line.split(":");
              event.sender.send("pptx:imageProgress", { current: parseInt(cur), total: parseInt(tot) });
            }
            if (line.startsWith("ERROR:")) reject(new Error(line.slice(6)));
          }
        });
        child.stderr.on("data", d => { stderr += d.toString(); });
        child.on("error", reject);
        child.on("close", code => {
          if (code === 0) resolve({ stdout, stderr });
          else reject(new Error(stderr || stdout || "python3 실패"));
        });
      });
      pdfPngSuccess = true;
    } catch(e) {
      renderErrors.push("PDF→PNG(Python) 실패: " + (e.message || "").slice(0, 200));
    }
  }

  if (!pdfPngSuccess && isWindows) {
    // Windows: LibreOffice PDF → PNG 변환
    const libreOfficePaths = [
      "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
      "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ];
    for (const soffice of libreOfficePaths) {
      if (!fs.existsSync(soffice)) continue;
      try {
        await new Promise((resolve, reject) => {
          const { spawn } = require("child_process");
          const child = spawn(soffice, [
            "--headless", "--convert-to", "png",
            "--outdir", pngDir, pdfPath
          ], { stdio: ["ignore", "pipe", "pipe"] });
          let stderr = "";
          child.stderr.on("data", d => { stderr += d.toString(); });
          child.on("error", reject);
          child.on("close", code => {
            if (code === 0) resolve();
            else reject(new Error(stderr || "LibreOffice PDF→PNG 실패"));
          });
        });
        if (collectPngFiles(pngDir).length > 0) { pdfPngSuccess = true; break; }
      } catch(e) {
        renderErrors.push("LibreOffice PDF→PNG 실패: " + (e.message || "").slice(0, 100));
      }
    }
  }

  if (!pdfPngSuccess && isWindows) {
    // Windows: PowerShell + Windows.Data.Pdf (Windows 10/11 내장 API)
    const winPdfPngScript = `
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]

function Await($wrt) {
    $task = $wrt.AsTask()
    $task.Wait()
    return $task.Result
}

$pdfPath = ${JSON.stringify(pdfPath)}
$outDir  = ${JSON.stringify(pngDir)}

try {
    $file = Await([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdfPath))
    $doc  = Await([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file))
    $pageCount = $doc.PageCount
    Write-Host "PAGES:$pageCount"

    for ($i = 0; $i -lt $pageCount; $i++) {
        $page = $doc.GetPage($i)
        $outPath = [System.IO.Path]::Combine($outDir, ("slide{0:D4}.png" -f ($i + 1)))
        $stream  = [Windows.Storage.Streams.InMemoryRandomAccessStream]::new()
        $opts    = [Windows.Data.Pdf.PdfPageRenderOptions]::new()
        $opts.DestinationWidth = 1920
        Await($page.RenderToStreamAsync($stream, $opts)) | Out-Null
        $bytes = [byte[]]::new($stream.Size)
        $ibr   = $stream.GetInputStreamAt(0)
        $reader = [Windows.Storage.Streams.DataReader]::new($ibr)
        Await($reader.LoadAsync($stream.Size)) | Out-Null
        $reader.ReadBytes($bytes)
        [System.IO.File]::WriteAllBytes($outPath, $bytes)
        $stream.Dispose()
        Write-Host "DONE:$($i+1):$pageCount"
    }
    Write-Host "SUCCESS"
} catch {
    Write-Host "ERROR:$_"
    exit 1
}
`.trim();
    const wpsFile = path.join(tmpDir, "pdf2png_win.ps1");
    fs.writeFileSync(wpsFile, winPdfPngScript, "utf8");
    try {
      await new Promise((resolve, reject) => {
        const { spawn } = require("child_process");
        const child = spawn("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", wpsFile], {
          stdio: ["ignore", "pipe", "pipe"]
        });
        let stdout = "", stderr = "", total = 0;
        child.stdout.on("data", chunk => {
          const text = chunk.toString("utf8");
          stdout += text;
          for (const raw of text.split("\n")) {
            const line = raw.trim();
            if (line.startsWith("PAGES:")) total = parseInt(line.slice(6)) || 0;
            else if (line.startsWith("DONE:")) {
              const parts = line.slice(5).split(":");
              event.sender.send("pptx:imageProgress", { current: parseInt(parts[0]) || 0, total: parseInt(parts[1]) || total });
            }
          }
        });
        child.stderr.on("data", d => { stderr += d.toString(); });
        child.on("error", reject);
        child.on("close", code => {
          if (code === 0 || stdout.includes("SUCCESS")) resolve({ stdout, stderr });
          else reject(new Error((stderr || stdout || "PowerShell PDF→PNG 실패").slice(0, 300)));
        });
      });
      if (collectPngFiles(pngDir).length > 0) pdfPngSuccess = true;
    } catch(e) {
      renderErrors.push("Windows.Data.Pdf PDF→PNG 실패: " + (e.message || "").slice(0, 100));
    }
  }

  if (!pdfPngSuccess) {
    fs.rmSync(tmpDir, { recursive: true, force: true });
    const detail = renderErrors.length ? "\n\n상세:\n- " + renderErrors.join("\n- ") : "";
    return { ok: false, message: "PDF→PNG 변환에 실패했습니다." + detail };
  }

  try {
    const pngFiles = collectPngFiles(pngDir);
    if (!pngFiles.length) {
      fs.rmSync(tmpDir, { recursive: true, force: true });
      return { ok: false, message: "PNG 파일이 생성되지 않았습니다." };
    }

    // macOS는 회전 보정, Windows는 정방향이므로 그대로
    const finalUrls = isMac
      ? await rotatePngs180(pngFiles)
      : pngFiles.map(f => `data:image/png;base64,${fs.readFileSync(f).toString("base64")}`);

    for (let idx = 0; idx < finalUrls.length; idx++) {
      event.sender.send("pptx:imageProgress", { current: idx + 1, total: finalUrls.length });
      event.sender.send("pptx:imageChunk", {
        index: idx,
        total: finalUrls.length,
        slide: { mediaURL: finalUrls[idx], mediaType: "image", bgColor: null }
      });
    }

    fs.rmSync(tmpDir, { recursive: true, force: true });
    return { ok: true, fileName: path.basename(filePath, path.extname(filePath)), total: finalUrls.length, source: pdfSource || "pdf", streamed: true };
  } catch(err) {
    fs.rmSync(tmpDir, { recursive: true, force: true });
    return { ok: false, message: "이미지 읽기 실패: " + err.message };
  }
});

// ──────────────────────────────────────────────────────────
// PPTX XML 파싱 헬퍼 함수들
// ──────────────────────────────────────────────────────────

/**
 * PPTX 관계 파일(.rels)의 XML을 파싱해 ID → 파일경로 매핑을 반환한다.
 * PPTX에서 이미지나 다른 파일을 참조할 때 ID를 사용하므로
 * 실제 파일 경로를 찾으려면 이 매핑이 필요하다.
 *
 * @param {string} xml - .rels 파일의 XML 텍스트
 * @returns {Object} { rId: targetPath } 형태의 매핑 객체
 */
function parseRelsXml(xml) {
  const map = {};
  const re  = /<Relationship\s[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;
  let m;
  // Windows에서 생성된 PPTX는 백슬래시를 사용하므로 슬래시로 정규화
  while ((m = re.exec(xml)) !== null) map[m[1]] = m[2].replace(/\\/g, "/");
  return map;
}

/**
 * PPTX 테마 파일에서 색상 팔레트와 폰트 스킴을 추출한다.
 * 슬라이드의 색상/폰트가 테마를 참조할 때 실제 값을 구하는 데 사용한다.
 *
 * @param {AdmZip} zip - 열린 PPTX ZIP 객체
 * @returns {Object} { colors: {dk1, lt1, ...}, majorFont, minorFont, ... }
 */
function extractPptxTheme(zip) {
  const theme = { colors: {}, majorFont: '', minorFont: '', majorFontEa: '', minorFontEa: '', fillStyles: [] };
  try {
    let themeEntry = zipGetEntry(zip, 'ppt/theme/theme1.xml');
    if (!themeEntry) {
      // theme1.xml이 없으면 다른 번호의 테마 파일 탐색
      themeEntry = zip.getEntries().find(e => /^ppt\/theme\/theme\d+\.xml$/i.test(e.entryName.replace(/\\/g, '/'))) || null;
    }
    if (!themeEntry) return theme;
    const xml = themeEntry.getData().toString('utf8');

    // 색상 스킴 추출 (dk1=어두운색1, lt1=밝은색1, accent1~6=강조색 등)
    for (const name of ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink']) {
      const re2  = new RegExp('<a:' + name + '>([\\s\\S]{0,300}?)<\\/a:' + name + '>');
      const m2   = xml.match(re2);
      if (!m2) continue;
      const srgb = m2[1].match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
      const sys  = m2[1].match(/<a:sysClr\s+[^>]*lastClr="([0-9A-Fa-f]{6})"/);
      if (srgb) theme.colors[name] = srgb[1];
      else if (sys) theme.colors[name] = sys[1];
    }

    // 폰트 스킴 (major=제목용, minor=본문용, Ea=동아시아 폰트)
    const mjLa = xml.match(/<a:majorFont>[\s\S]*?<a:latin\s+typeface="([^"]+)"/);
    const mnLa = xml.match(/<a:minorFont>[\s\S]*?<a:latin\s+typeface="([^"]+)"/);
    const mjEa = xml.match(/<a:majorFont>[\s\S]*?<a:ea\s+typeface="([^"]+)"/);
    const mnEa = xml.match(/<a:minorFont>[\s\S]*?<a:ea\s+typeface="([^"]+)"/);
    if (mjLa) theme.majorFont   = mjLa[1];
    if (mnLa) theme.minorFont   = mnLa[1];
    if (mjEa) theme.majorFontEa = mjEa[1];
    if (mnEa) theme.minorFontEa = mnEa[1];

    // fillStyleLst를 읽어 fillRef/bgRef idx 역참조에 사용
    const fillLstM = xml.match(/<a:fillStyleLst>([\s\S]*?)<\/a:fillStyleLst>/);
    if (fillLstM) {
      const fillChunk = fillLstM[1];
      const fillRe = /<(a:solidFill|a:gradFill|a:pattFill)\b([\s\S]*?)<\/\1>/g;
      let fm;
      while ((fm = fillRe.exec(fillChunk)) !== null) {
        const part = fm[0];
        const srgb = part.match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
        const sys = part.match(/<a:sysClr\s+[^>]*?(?:val|lastClr)="([0-9A-Fa-f]{6})"/);
        const sch = part.match(/<a:schemeClr\s+val="([^"]+)"/);
        if (srgb) theme.fillStyles.push("#" + srgb[1]);
        else if (sys) theme.fillStyles.push("#" + sys[1]);
        else if (sch && theme.colors[sch[1]]) theme.fillStyles.push("#" + theme.colors[sch[1]]);
        else theme.fillStyles.push(null);
      }
    }
  } catch(e) {}
  return theme;
}

/**
 * 대소문자를 구분하지 않고 ZIP 내 항목을 찾는다.
 * macOS APFS 같은 대소문자 구분 파일시스템 환경에서도 올바르게 동작한다.
 *
 * @param {AdmZip} zip       - 열린 ZIP 객체
 * @param {string} entryPath - 찾을 항목의 경로
 * @returns {ZipEntry|null} 찾은 항목, 없으면 null
 */
function zipGetEntry(zip, entryPath) {
  const norm  = entryPath.replace(/\\/g, "/");
  let entry   = zip.getEntry(norm);
  if (entry) return entry;
  const lower = norm.toLowerCase();
  return zip.getEntries().find(e => e.entryName.replace(/\\/g, "/").toLowerCase() === lower) || null;
}

/**
 * XML에서 HTML 특수문자 엔티티를 원래 문자로 복원한다.
 * PPTX XML 텍스트에 &amp; &lt; 등이 포함될 수 있다.
 *
 * @param {string} t - 엔티티가 포함된 텍스트
 * @returns {string} 복원된 텍스트
 */
function xmlDecodeText(t) {
  return String(t || "")
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}

/**
 * 관계 ID(rId)에 해당하는 이미지를 찾아 base64 data URL로 반환한다.
 * 경로 규칙이 다양하므로 여러 후보 경로를 순서대로 시도한다.
 *
 * @param {string} rId      - 관계 ID (예: 'rId1')
 * @param {Object} relsMap  - parseRelsXml() 결과
 * @param {AdmZip} zip      - 열린 ZIP 객체
 * @param {string} baseDir  - 기본 디렉터리 (예: 'ppt/')
 * @returns {string|null} base64 data URL, 또는 null
 */
function getImageBase64(rId, relsMap, zip, baseDir) {
  let target = relsMap[rId];
  if (!target) return null;
  target = target.replace(/\\/g, "/");

  // 후보 경로 생성 (PPTX 내 상대 경로 규칙이 파일마다 다를 수 있음)
  const tries = [];
  if (target.startsWith("../")) {
    tries.push(baseDir + target.slice(3));
    tries.push("ppt/" + target.slice(3));
  } else if (target.startsWith("/")) {
    tries.push(target.slice(1));
  } else if (/^https?:\/\//.test(target)) {
    return null;  // 외부 URL은 지원 안 함
  } else {
    tries.push(baseDir + target);
    tries.push("ppt/" + target);
    tries.push(target);
  }
  // 파일명만으로 ppt/media/ 하위 재탐색 (경로가 틀려도 파일명으로 찾기)
  const basename = target.split("/").pop();
  if (basename) tries.push("ppt/media/" + basename);

  let entry = null;
  for (const p of tries) {
    entry = zipGetEntry(zip, p);
    if (entry) break;
  }
  if (!entry) return null;

  // MIME 타입 결정
  const ext   = path.extname(entry.entryName).slice(1).toLowerCase();
  const mimes = { jpg:"image/jpeg", jpeg:"image/jpeg", png:"image/png", gif:"image/gif", bmp:"image/bmp", webp:"image/webp", tiff:"image/tiff", svg:"image/svg+xml", emf:"image/x-emf", wmf:"image/x-wmf" };
  const mime  = mimes[ext] || "image/jpeg";
  return "data:" + mime + ";base64," + entry.getData().toString("base64");
}

function resolveOpenXmlTargetPath(currentDir, targetPath) {
  let t = String(targetPath || "").replace(/\\/g, "/");
  const cur = String(currentDir || "").replace(/\\/g, "/");
  if (!t) return "";
  if (t.startsWith("/")) return t.slice(1);
  if (/^https?:\/\//i.test(t)) return "";
  const baseParts = cur.split("/").filter(Boolean);
  const targetParts = t.split("/").filter(Boolean);
  while (targetParts.length && targetParts[0] === "..") {
    targetParts.shift();
    if (baseParts.length) baseParts.pop();
  }
  while (targetParts.length && targetParts[0] === ".") {
    targetParts.shift();
  }
  return [...baseParts, ...targetParts].join("/");
}

/**
 * PPTX 슬라이드 XML을 파싱해 앱 슬라이드 객체로 변환한다.
 *
 * 파싱 순서:
 *   1. 배경: <p:bg> 배경 이미지/색상 → 슬라이드 레이아웃 → 슬라이드 마스터 순으로 탐색
 *   2. <p:pic> 이미지 요소 중 가장 큰 것을 배경으로 사용
 *   3. <p:sp> 텍스트 도형들을 파싱해 폰트/위치/내용 추출
 *   4. placeholder 타입, 폰트 크기 등을 기준으로 메인 가사 텍스트 식별
 *   5. 나머지 텍스트는 textBoxes 배열에 추가
 *
 * @param {string} xml     - 슬라이드 XML 텍스트
 * @param {Object} relsMap - 관계 ID → 파일경로 매핑
 * @param {AdmZip} zip     - 열린 ZIP 객체
 * @param {number} slideW  - 슬라이드 너비 (EMU)
 * @param {number} slideH  - 슬라이드 높이 (EMU)
 * @param {Object} theme   - extractPptxTheme() 결과
 * @returns {Object} 앱 슬라이드 객체
 */
function parsePptxSlide(xml, relsMap, zip, slideW, slideH, theme = {}) {
  const slide = {
    lyrics: "", sub: "",
    fontSize: 36, fontColor: "#ffffff", fontFamily: "Noto Serif KR",
    textX: 50, textY: 50, textWidth: null,
    textShadowEnabled: true, textShadowIntensity: 80,
    textOutlineEnabled: false, textOutlineWidth: 2,
    layoutType: "centerLyrics",
    mediaURL: null, mediaType: null,
    bgColor: null,
    textBoxes: [],
    shapeBoxes: []
  };

  const baseDir = "ppt/";
  const emuToPct = (value, total) => {
    const num = Number(value);
    const den = Number(total);
    if (!Number.isFinite(num) || !Number.isFinite(den) || den <= 0) return 0;
    return Math.round((num / den) * 100 * 10) / 10;
  };
  const resolveSchemeColor = val => (theme.colors && theme.colors[val]) || null;
  const resolveFillRefColor = (idxRaw) => {
    const idx = parseInt(String(idxRaw || ""), 10);
    if (!Number.isFinite(idx) || !Array.isArray(theme.fillStyles) || !theme.fillStyles.length) return null;
    // PPTX fillRef/bgRef idx는 문서에 따라 1-based 혹은 100x/1000x 계열로 나타날 수 있어 보정
    const candidates = [
      idx,
      idx - 1000,
      idx - 1001,
      idx - 100,
      idx - 1
    ].filter(n => Number.isFinite(n) && n >= 1 && n <= theme.fillStyles.length);
    for (const c of candidates) {
      const color = theme.fillStyles[c - 1];
      if (color) return color;
    }
    return null;
  };
  const readFillColorHex = (chunk = "") => {
    if (!chunk) return null;
    const srgb = chunk.match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
    if (srgb) return "#" + srgb[1];
    const sys = chunk.match(/<a:sysClr\s+[^>]*?(?:val|lastClr)="([0-9A-Fa-f]{6})"/);
    if (sys) return "#" + sys[1];
    const grad = chunk.match(/<a:gradFill[\s\S]*?<a:gs\b[^>]*>[\s\S]*?(?:<a:srgbClr\s+val="([0-9A-Fa-f]{6})"|<a:sysClr\s+[^>]*?(?:val|lastClr)="([0-9A-Fa-f]{6})"|<a:schemeClr\s+val="([^"]+)")/);
    if (grad) {
      if (grad[1]) return "#" + grad[1];
      if (grad[2]) return "#" + grad[2];
      if (grad[3]) {
        const resolved = resolveSchemeColor(grad[3]);
        if (resolved) return "#" + resolved.replace(/^#/, "");
      }
    }
    const sch = chunk.match(/<a:schemeClr\s+val="([^"]+)"/);
    if (sch) {
      const resolved = resolveSchemeColor(sch[1]);
      if (resolved) return "#" + resolved.replace(/^#/, "");
    }
    const fillRef = chunk.match(/<(?:a:fillRef|a:bgRef|p:bgRef)\s+[^>]*\bidx="(\d+)"/);
    if (fillRef) {
      const refColor = resolveFillRefColor(fillRef[1]);
      if (refColor) return refColor;
    }
    return null;
  };

  // ── 1. 배경 처리 ──
  // 우선순위: 슬라이드 직접 지정 > 슬라이드 레이아웃 > 슬라이드 마스터

  // (a) <p:bg> 안의 배경 이미지
  const bgBlipM = xml.match(/<p:bg>[\s\S]*?<a:blip\s[^>]*r:embed="([^"]+)"/);
  if (bgBlipM) {
    const dataUrl = getImageBase64(bgBlipM[1], relsMap, zip, baseDir);
    if (dataUrl) { slide.mediaURL = dataUrl; slide.mediaType = "image"; }
  }

  // (a2) <p:bg> 안의 단색 배경
  if (!slide.mediaURL) {
    const bgChunkM = xml.match(/<p:bg>([\s\S]*?)<\/p:bg>/);
    const bgColor = bgChunkM ? readFillColorHex(bgChunkM[1]) : null;
    if (bgColor) slide.bgColor = bgColor;
  }

  // (a3) 슬라이드 레이아웃 → 슬라이드 마스터 배경 확인 (슬라이드 자체 배경이 없을 때)
  if (!slide.mediaURL && !slide.bgColor) {
    const layoutRel = Object.entries(relsMap).find(([, t]) => t.includes("slideLayout"));
    if (layoutRel) {
      const layoutTarget = layoutRel[1].replace(/\\/g, "/");
      const layoutPath   = resolveOpenXmlTargetPath("ppt/slides", layoutTarget);
      try {
        const layoutEntry = zipGetEntry(zip, layoutPath);
        if (layoutEntry) {
          const lxml      = layoutEntry.getData().toString("utf8");
          const lRelsPath  = layoutPath.replace(/\/([^/]+)$/, "/_rels/$1.rels");
          const lRelsEntry = zipGetEntry(zip, lRelsPath);
          const lRelsMap   = lRelsEntry ? parseRelsXml(lRelsEntry.getData().toString("utf8")) : {};

          const lBgBlip = lxml.match(/<p:bg>[\s\S]*?<a:blip\s[^>]*r:embed="([^"]+)"/);
          if (lBgBlip) {
            const lDataUrl = getImageBase64(lBgBlip[1], lRelsMap, zip, "ppt/");
            if (lDataUrl) { slide.mediaURL = lDataUrl; slide.mediaType = "image"; }
          }
          if (!slide.mediaURL) {
            const lBgChunkM = lxml.match(/<p:bg>([\s\S]*?)<\/p:bg>/);
            const lBgColor = lBgChunkM ? readFillColorHex(lBgChunkM[1]) : null;
            if (lBgColor) slide.bgColor = lBgColor;
          }

          // 레이아웃에도 배경이 없으면 슬라이드 마스터 확인
          if (!slide.mediaURL && !slide.bgColor) {
            const masterRel = Object.entries(lRelsMap).find(([, t]) => t.includes("slideMaster"));
            if (masterRel) {
              const masterTarget = masterRel[1].replace(/\\/g, "/");
              const layoutDir = layoutPath.replace(/\/[^/]+$/, "");
              const masterPath = resolveOpenXmlTargetPath(layoutDir, masterTarget);
              try {
                const masterEntry = zipGetEntry(zip, masterPath);
                if (masterEntry) {
                  const mxml      = masterEntry.getData().toString("utf8");
                  const mRelsPath  = masterPath.replace(/\/([^/]+)$/, "/_rels/$1.rels");
                  const mRelsEntry = zipGetEntry(zip, mRelsPath);
                  const mRelsMap   = mRelsEntry ? parseRelsXml(mRelsEntry.getData().toString("utf8")) : {};

                  const mBgBlip = mxml.match(/<p:bg>[\s\S]*?<a:blip\s[^>]*r:embed="([^"]+)"/);
                  if (mBgBlip) {
                    const mDataUrl = getImageBase64(mBgBlip[1], mRelsMap, zip, "ppt/");
                    if (mDataUrl) { slide.mediaURL = mDataUrl; slide.mediaType = "image"; }
                  }
                  if (!slide.mediaURL) {
                    const mBgChunkM = mxml.match(/<p:bg>([\s\S]*?)<\/p:bg>/);
                    const mBgColor = mBgChunkM ? readFillColorHex(mBgChunkM[1]) : null;
                    if (mBgColor) slide.bgColor = mBgColor;
                  }
                }
              } catch(e) {}
            }
          }
        }
      } catch(e) {}
    }
  }

  // (a4) 배경이 <p:bg>가 아니라 "큰 도형 채우기"로 만든 케이스 보정
  // 텍스트가 없는 도형 중 면적이 큰 사각형의 solidFill을 배경으로 사용한다.
  if (!slide.mediaURL && !slide.bgColor) {
    const shapeRe = /<p:sp>([\s\S]*?)<\/p:sp>/g;
    let sm;
    let bestShape = null;
    while ((sm = shapeRe.exec(xml)) !== null) {
      const sp = sm[1];
      if (sp.includes("<p:txBody>")) continue;
      const offM = sp.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"/);
      const extM = sp.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
      if (!offM || !extM) continue;
      const x = parseInt(offM[1]);
      const y = parseInt(offM[2]);
      const cx = parseInt(extM[1]);
      const cy = parseInt(extM[2]);
      if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) continue;
      const fill = readFillColorHex(sp);
      const blipEmbed = (sp.match(/<a:blip\b[^>]*r:embed="([^"]+)"/) || [])[1] || null;
      const fillImage = blipEmbed ? getImageBase64(blipEmbed, relsMap, zip, baseDir) : null;
      if (!fill && !fillImage) continue;

      const area = cx * cy;
      const coverRatio = area / (slideW * slideH);
      if (coverRatio < 0.45) continue; // 배경 후보는 최소 절반 가까이 덮어야 함

      const centerX = x + cx / 2;
      const centerY = y + cy / 2;
      const centerBias =
        Math.abs(centerX - slideW / 2) / slideW +
        Math.abs(centerY - slideH / 2) / slideH;
      const score = coverRatio - centerBias * 0.25;

      if (!bestShape || score > bestShape.score) {
        bestShape = { fill, fillImage, score };
      }
    }
    if (bestShape) {
      if (bestShape.fillImage) {
        slide.mediaURL = bestShape.fillImage;
        slide.mediaType = "image";
      } else {
        slide.bgColor = bestShape.fill;
      }
    }
  }

  // (b) <p:pic> 이미지 요소 — 가장 큰 이미지를 배경으로 사용
  const allPics = [];
  const picRe2  = /<p:pic>([\s\S]*?)<\/p:pic>/g;
  let picM2;
  while ((picM2 = picRe2.exec(xml)) !== null) {
    const picXml = picM2[1];
    const embedM = picXml.match(/r:embed="([^"]+)"/);
    const offM   = picXml.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"/);
    const extM   = picXml.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
    if (!embedM) continue;
    const x  = offM ? parseInt(offM[1]) : 0;
    const y  = offM ? parseInt(offM[2]) : 0;
    const cx = extM ? parseInt(extM[1]) : slideW;
    const cy = extM ? parseInt(extM[2]) : slideH;
    allPics.push({ rId: embedM[1], x, y, cx, cy, area: cx * cy });
  }
  allPics.sort((a, b) => b.area - a.area);  // 면적 큰 순서로 정렬
  if (!slide.mediaURL && allPics.length) {
    const bg      = allPics[0];
    const dataUrl = getImageBase64(bg.rId, relsMap, zip, baseDir);
    if (dataUrl) {
      slide.mediaURL  = dataUrl;
      slide.mediaType = "image";
      // 위치/크기를 % 단위로 저장 (앱의 mOffX/Y/Scale 필드)
      slide.mediaOffX  = Math.round((bg.x + bg.cx / 2) / slideW * 100 * 10) / 10;
      slide.mediaOffY  = Math.round((bg.y + bg.cy / 2) / slideH * 100 * 10) / 10;
      slide.mediaScale = Math.round(Math.max(bg.cx / slideW, bg.cy / slideH) * 100 * 10) / 10;
    }
  }

  // ── 2. 슬라이드 레이아웃에서 placeholder 위치 상속 ──
  const layoutPhs = getPptxLayoutPlaceholders(relsMap, zip);

  // 슬라이드 높이를 pt로 변환 (폰트 크기 비율 계산에 사용)
  const slideH_pt = (slideH / 914400) * 72;  // EMU → 인치 → pt

  const pickThemeFont = (tf, preferEa = false) => {
    const val = String(tf || "");
    if (!val) return "";
    if (val === "+mj-lt" || val === "+mj-ea") return (preferEa && theme.majorFontEa) || theme.majorFont || "";
    if (val === "+mn-lt" || val === "+mn-ea") return (preferEa && theme.minorFontEa) || theme.minorFont || "";
    return (!val.startsWith("+") && !val.startsWith("@")) ? val : "";
  };

  const parseShapeColor = (shapeXml = "", fallback = null) => {
    if (!shapeXml) return fallback;
    const srgb = shapeXml.match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
    if (srgb) return "#" + srgb[1];
    const sys = shapeXml.match(/<a:sysClr\s+[^>]*?(?:val|lastClr)="([0-9A-Fa-f]{6})"/);
    if (sys) return "#" + sys[1];
    const sch = shapeXml.match(/<a:schemeClr\s+val="([^"]+)"/);
    if (sch) {
      const resolved = resolveSchemeColor(sch[1]);
      if (resolved) return "#" + String(resolved).replace(/^#/, "");
    }
    return fallback;
  };

  const extractDecorShapesWithDom = () => {
    try {
      const doc = new DOMParser().parseFromString(xml, "text/xml");
      const spNodes = doc.getElementsByTagName("p:sp");
      const result = [];

      const getFirst = (el, tag) => {
        const list = el.getElementsByTagName(tag);
        return list && list.length ? list[0] : null;
      };
      const toInt = (v, d = 0) => {
        const n = parseInt(String(v ?? ""), 10);
        return Number.isFinite(n) ? n : d;
      };

      for (let i = 0; i < spNodes.length; i++) {
        const spEl = spNodes[i];
        const txBodyEl = getFirst(spEl, "p:txBody");
        if (txBodyEl) continue;

        const xfrm = getFirst(spEl, "a:xfrm");
        const off = xfrm ? getFirst(xfrm, "a:off") : null;
        const ext = xfrm ? getFirst(xfrm, "a:ext") : null;
        if (!off || !ext) continue;

        const sx = toInt(off.getAttribute("x"), 0);
        const sy = toInt(off.getAttribute("y"), 0);
        const sw = toInt(ext.getAttribute("cx"), 0);
        const sh = toInt(ext.getAttribute("cy"), 0);
        if (sw <= 0 || sh <= 0) continue;

        const spXml = String(spEl.toString ? spEl.toString() : "");
        const noFill = /<a:noFill\b/.test(spXml);
        const fillColor = noFill ? "transparent" : (parseShapeColor(spXml, null) || "transparent");
        const lineColor = parseShapeColor((spXml.match(/<a:ln\b[\s\S]*?<\/a:ln>/) || [])[0] || "", null);
        const lineW = (() => {
          const lw = (spXml.match(/<a:ln\b[^>]*\bw="(\d+)"/) || [])[1];
          if (!lw) return 0;
          const pt = parseInt(lw, 10) / 12700;
          return Number.isFinite(pt) ? Math.max(0, Math.min(24, Math.round(pt * 10) / 10)) : 0;
        })();
        const prstGeom = (spXml.match(/<a:prstGeom\b[^>]*\bprst="([^"]+)"/) || [])[1] || "rect";

        // 채우기/선이 전혀 없는 보조 도형은 제외
        if (fillColor === "transparent" && !lineColor && lineW <= 0) continue;

        result.push({
          x: Math.max(0, Math.min(100, emuToPct(sx + sw / 2, slideW))),
          y: Math.max(0, Math.min(100, emuToPct(sy + sh / 2, slideH))),
          w: Math.max(0.2, Math.min(100, emuToPct(sw, slideW))),
          h: Math.max(0.2, Math.min(100, emuToPct(sh, slideH))),
          fillColor,
          lineColor: lineColor || "transparent",
          lineWidth: lineW,
          shapeType: prstGeom
        });
      }
      return result;
    } catch (e) {
      return [];
    }
  };

  const extractShapesWithDom = () => {
    try {
      const doc = new DOMParser().parseFromString(xml, "text/xml");
      const spNodes = doc.getElementsByTagName("p:sp");
      const domShapes = [];

      const getFirst = (el, tag) => {
        const list = el.getElementsByTagName(tag);
        return list && list.length ? list[0] : null;
      };
      const toInt = (v, d = 0) => {
        const n = parseInt(String(v ?? ""), 10);
        return Number.isFinite(n) ? n : d;
      };

      for (let i = 0; i < spNodes.length; i++) {
        const spEl = spNodes[i];
        const txBodyEl = getFirst(spEl, "p:txBody");
        if (!txBodyEl) continue;

        const phEl = getFirst(spEl, "p:ph");
        const phType = phEl ? (phEl.getAttribute("type") || "body") : "";
        const phIdx = phEl ? (phEl.getAttribute("idx") || "1") : "";

        const xfrm = getFirst(spEl, "a:xfrm");
        const off = xfrm ? getFirst(xfrm, "a:off") : null;
        const ext = xfrm ? getFirst(xfrm, "a:ext") : null;
        const lph = layoutPhs[phIdx] || layoutPhs[phType] || null;

        const sx = off ? toInt(off.getAttribute("x"), lph ? lph.x : 0) : (lph ? lph.x : 0);
        const sy = off ? toInt(off.getAttribute("y"), lph ? lph.y : 0) : (lph ? lph.y : 0);
        const sw = ext ? toInt(ext.getAttribute("cx"), lph ? lph.w : slideW) : (lph ? lph.w : slideW);
        const sh = ext ? toInt(ext.getAttribute("cy"), lph ? lph.h : slideH) : (lph ? lph.h : slideH);
        if (sw <= 0 || sh <= 0) continue;

        const cxPct = emuToPct(sx + sw / 2, slideW);
        const wPct = emuToPct(sw, slideW);
        const hPct = emuToPct(sh, slideH);

        const bodyPr = getFirst(txBodyEl, "a:bodyPr");
        const anchor = bodyPr ? (bodyPr.getAttribute("anchor") || "t") : "t";
        const cyPct =
          anchor === "ctr" ? emuToPct(sy + sh / 2, slideH) :
          anchor === "b" ? emuToPct(sy + sh, slideH) :
          emuToPct(sy, slideH);

        const pNodes = txBodyEl.getElementsByTagName("a:p");
        let mainFontSize = 0;
        let mainFontColor = "ffffff";
        let mainFontFace = "";
        let mainAlign = "center";
        let alignFixed = false;
        const lines = [];

        for (let pIdx = 0; pIdx < pNodes.length; pIdx++) {
          const p = pNodes[pIdx];
          if (!alignFixed) {
            const algn = p.getAttribute("algn");
            if (algn) {
              alignFixed = true;
              if (algn === "l") mainAlign = "left";
              else if (algn === "r") mainAlign = "right";
              else mainAlign = "center";
            }
          }

          const rNodes = p.getElementsByTagName("a:r");
          let lineText = "";
          for (let rIdx = 0; rIdx < rNodes.length; rIdx++) {
            const r = rNodes[rIdx];
            const rPr = getFirst(r, "a:rPr");
            if (rPr) {
              const runSz = toInt(rPr.getAttribute("sz"), 0);
              if (runSz > 0) mainFontSize = Math.max(mainFontSize, Math.round(runSz / 100));
              const srgb = getFirst(rPr, "a:srgbClr");
              const sys = getFirst(rPr, "a:sysClr");
              const sch = getFirst(rPr, "a:schemeClr");
              if (srgb && srgb.getAttribute("val")) {
                mainFontColor = srgb.getAttribute("val");
              } else if (sys && (sys.getAttribute("val") || sys.getAttribute("lastClr"))) {
                mainFontColor = sys.getAttribute("val") || sys.getAttribute("lastClr");
              } else if (sch && sch.getAttribute("val")) {
                const resolved = resolveSchemeColor(sch.getAttribute("val"));
                if (resolved) mainFontColor = resolved.replace(/^#/, "");
              }

              const ea = getFirst(rPr, "a:ea");
              const latin = getFirst(rPr, "a:latin");
              const runFont = pickThemeFont(ea ? ea.getAttribute("typeface") : "", true) ||
                pickThemeFont(latin ? latin.getAttribute("typeface") : "", false);
              if (runFont) mainFontFace = runFont;
            }
            const tNodes = r.getElementsByTagName("a:t");
            if (tNodes.length) lineText += xmlDecodeText(tNodes[0].textContent || "");
          }
          if (lineText.trim()) lines.push(lineText.trimEnd());
        }

        if (!lines.length) continue;

        domShapes.push({
          text: lines.join("\n").trim(),
          cxPct, cyPct, wPct, hPct,
          fontSize: mainFontSize || 24,
          fontColor: "#" + String(mainFontColor || "ffffff").replace(/^#/, ""),
          fontFace: mainFontFace,
          align: mainAlign,
          anchor,
          phType,
          area: wPct * hPct
        });
      }

      return domShapes;
    } catch (e) {
      return [];
    }
  };

  // ── 3. 텍스트 도형(<p:sp>) 추출 ──
  const shapes = extractShapesWithDom();
  if (!shapes.length) {
    // TODO: DOM 파서가 일부 shape만 반환해도 non-empty면 여기 fallback이 비활성화되어 누락 shape를 보완하지 못할 수 있다.
    const spRe   = /<p:sp>([\s\S]*?)<\/p:sp>/g;
    let spM;
    while ((spM = spRe.exec(xml)) !== null) {
    const sp = spM[1];
    if (!sp.includes("<p:txBody>")) continue;

    // placeholder 타입과 인덱스 (body=본문, title=제목, ctrTitle=중앙제목 등)
    const phTagM  = sp.match(/<p:ph\b([^>]*)>/);
    const phTypeM = phTagM ? phTagM[1].match(/\btype="([^"]+)"/) : null;
    const phIdxM  = phTagM ? phTagM[1].match(/\bidx="([^"]+)"/)  : null;
    const phType  = phTypeM ? phTypeM[1] : (phTagM ? "body" : "");
    const phIdx   = phIdxM  ? phIdxM[1]  : (phTagM ? "1"   : "");

    // 위치/크기: 명시되어 있으면 사용, 없으면 레이아웃 placeholder에서 상속
    const offM = sp.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"/);
    const extM = sp.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
    const lph  = layoutPhs[phIdx] || layoutPhs[phType] || null;

    const sx = offM ? parseInt(offM[1]) : (lph ? lph.x : 0);
    const sy = offM ? parseInt(offM[2]) : (lph ? lph.y : 0);
    const sw = extM ? parseInt(extM[1]) : (lph ? lph.w : slideW);
    const sh = extM ? parseInt(extM[2]) : (lph ? lph.h : slideH);

    // 위치/크기를 % 단위로 변환 (중심점 기준)
    const cxPct = emuToPct(sx + sw / 2, slideW);
    const wPct  = emuToPct(sw, slideW);
    const hPct  = emuToPct(sh, slideH);

    // 수직 앵커 (PPTX 기본: 't'=상단, 'ctr'=중앙, 'b'=하단)
    const bprAncM = sp.match(/<a:bodyPr\b[^>]*\banchor="([^"]+)"/);
    const anchor  = bprAncM ? bprAncM[1] : 't';
    const cyPct   =
      anchor === 'ctr' ? emuToPct(sy + sh / 2, slideH) :
      anchor === 'b'   ? emuToPct(sy + sh, slideH) :
                         emuToPct(sy, slideH);

    // txBody 레벨 기본 서식 (run/단락 서식이 없을 때 폴백으로 사용)
    const bodyDefSzM   = sp.match(/<a:defRPr\b[^>]*\bsz="(\d+)"/);
    const bodyDefSize  = bodyDefSzM ? Math.round(parseInt(bodyDefSzM[1]) / 100) : 0;
    const bodyDefClrM  = sp.match(/<a:defRPr\b[^>]*>[\s\S]*?<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
    const bodyDefSchM  = !bodyDefClrM && sp.match(/<a:defRPr\b[^>]*>[\s\S]*?<a:schemeClr\s+val="([^"]+)"/);
    const bodyDefColor = bodyDefClrM ? bodyDefClrM[1] : (bodyDefSchM ? resolveSchemeColor(bodyDefSchM[1]) : null);

    let mainFontSize  = 0, mainFontColor = "ffffff", mainFontFace = "";
    let mainAlign     = "center";
    let alignFixed    = false;
    const lines       = [];

    // 단락(<a:p>) 순서대로 텍스트 추출
    const paraRe = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
    let paraM;
    while ((paraM = paraRe.exec(sp)) !== null) {
      const para = paraM[1];

      // 첫 번째 명시 정렬만 사용 (l=왼쪽, r=오른쪽, 그 외=중앙)
      if (!alignFixed) {
        const algM = para.match(/\balgn="([^"]+)"/);
        if (algM) {
          alignFixed = true;
          if      (algM[1] === "l") mainAlign = "left";
          else if (algM[1] === "r") mainAlign = "right";
          else                      mainAlign = "center";
        }
      }

      // 단락 기본 폰트 크기/색상 (run 서식 없을 때 폴백)
      const pDefSzM   = para.match(/<a:defRPr\b[^>]*\bsz="(\d+)"/);
      const pDefSize  = pDefSzM ? Math.round(parseInt(pDefSzM[1]) / 100) : bodyDefSize;
      const pDefClrM  = para.match(/<a:defRPr\b[^>]*>[\s\S]*?<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
      const pDefSchM  = !pDefClrM && para.match(/<a:defRPr\b[^>]*>[\s\S]*?<a:schemeClr\s+val="([^"]+)"/);
      const pDefColor = pDefClrM ? pDefClrM[1] : (pDefSchM ? resolveSchemeColor(pDefSchM[1]) : bodyDefColor);

      let lineText = "";
      // run(<a:r>): 실제 텍스트 데이터를 담는 최소 단위
      const runRe = /<a:r>([\s\S]*?)<\/a:r>/g;
      let runM;
      while ((runM = runRe.exec(para)) !== null) {
        const run = runM[1];

        // 폰트 크기 (PPTX 단위: 1/100pt → pt로 변환)
        const szM2  = run.match(/\bsz="(\d+)"/);
        const runSz = szM2 ? Math.round(parseInt(szM2[1]) / 100) : pDefSize;
        if (runSz > mainFontSize) mainFontSize = runSz;

        // 폰트 색상 (srgbClr/sysClr 우선, schemeClr 폴백)
        const clrM = run.match(/(?:<a:srgbClr|<a:sysClr)\s+[^>]*?(?:val|lastClr)="([0-9A-Fa-f]{6})"/);
        if (clrM) {
          mainFontColor = clrM[1];
        } else {
          const schM     = run.match(/<a:schemeClr\s+val="([^"]+)"/);
          const resolved = schM ? resolveSchemeColor(schM[1]) : null;
          if (resolved) mainFontColor = resolved;
          else if (pDefColor && mainFontColor === "ffffff") mainFontColor = pDefColor;
        }

        // 폰트 이름 (한글 폰트: <a:ea> 우선, 라틴 폰트: <a:latin> 폴백)
        const eaTf   = (run.match(/<a:ea\b[^>]*\btypeface="([^"]+)"/) || [])[1] || '';
        const laTf   = (run.match(/<a:latin\b[^>]*\btypeface="([^"]+)"/) || [])[1] || '';
        const runFont = pickThemeFont(eaTf, true) || pickThemeFont(laTf, false);
        if (runFont) mainFontFace = runFont;

        // 텍스트 내용
        const tM = run.match(/<a:t>([\s\S]*?)<\/a:t>/);
        if (tM) lineText += xmlDecodeText(tM[1]);
      }
      if (lineText.trim()) lines.push(lineText.trimEnd());
    }

    if (!lines.length) continue;

      shapes.push({
      text: lines.join("\n").trim(),
      cxPct, cyPct, wPct, hPct,
      fontSize:  mainFontSize || bodyDefSize || 24,
      fontColor: "#" + mainFontColor,
      fontFace:  mainFontFace,
      align:     mainAlign,
      anchor,
      phType,
      area: wPct * hPct
      });
    }
  }

  // ── 4. 메인 텍스트 식별 ──
  // placeholder 타입 'body' + 폰트 크기 + 면적을 종합해 메인 가사를 선정
  shapes.sort((a, b) => {
    const scoreA = (a.phType === "body" ? 10000 : 0) + a.fontSize * 100 + a.area;
    const scoreB = (b.phType === "body" ? 10000 : 0) + b.fontSize * 100 + b.area;
    return scoreB - scoreA;
  });

  // PPTX pt → 앱 px 변환 (1080p 기준 비율 보존)
  const ptToAppPx = pt => Math.max(8, Math.min(200, Math.round(pt * 1080 / (slideH_pt * 1.4))));
  slide.importedTextBoxes = shapes.map(s => ({
    text: s.text,
    x: Math.max(0, Math.min(100, s.cxPct)),
    y: Math.max(0, Math.min(100, s.cyPct)),
    w: Math.min(99, s.wPct),
    h: Math.max(4, s.hPct),
    width: Math.min(99, s.wPct),
    height: Math.max(4, s.hPct),
    fontSize: ptToAppPx(s.fontSize),
    color: s.fontColor,
    fontColor: s.fontColor,
    fontFamily: pptxFontToKor(s.fontFace),
    align: s.align,
    textAnchor: s.anchor || "t",
    role: s.phType || "custom"
  }));

  const mainShape = shapes[0];
  if (mainShape) {
    slide.lyrics     = mainShape.text;
    slide.fontSize   = ptToAppPx(mainShape.fontSize);
    slide.fontColor  = mainShape.fontColor;
    slide.fontFamily = pptxFontToKor(mainShape.fontFace);
    slide.textX      = Math.max(0, Math.min(100, mainShape.cxPct));
    slide.textY      = Math.max(0, Math.min(100, mainShape.cyPct));
    slide.textWidth  = Math.min(99, mainShape.wPct);
    slide.textAnchor = mainShape.anchor || 't';
    slide.layoutType = "centerLyrics";
  }

  // title/ctrTitle placeholder → 소제목(sub)으로 설정
  const titleShape = shapes.find(s => s.phType === "title" || s.phType === "ctrTitle");
  if (titleShape && titleShape !== mainShape) {
    slide.sub = titleShape.text.replace(/\n/g, " ").slice(0, 80);
  }

  // 나머지 텍스트 도형들은 추가 텍스트박스로 추가
  const mkId = () => Math.random().toString(36).slice(2, 10);
  for (const s of shapes) {
    if (s === mainShape || s === titleShape) continue;
    if (!s.text.trim()) continue;
    slide.textBoxes.push({
      id:        mkId(),
      label:     s.text.slice(0, 20),
      text:      s.text,
      role:      "custom",
      x:         Math.max(0, Math.min(100, s.cxPct)),
      y:         Math.max(0, Math.min(100, s.cyPct)),
      w:         Math.min(99, s.wPct),
      h:         Math.max(4, s.hPct),
      width:     Math.min(99, s.wPct),
      height:    Math.max(4, s.hPct),
      fontSize:  ptToAppPx(s.fontSize),
      color:     s.fontColor,
      fontColor: s.fontColor,
      fontFamily: pptxFontToKor(s.fontFace),
      align:     s.align,
      textAnchor: s.anchor || 't'
    });
  }

slide.shapeBoxes = extractDecorShapesWithDom();

const bigBgShape = (slide.shapeBoxes || [])
  .filter(s => s && s.fillColor && s.fillColor !== 'transparent')
  .sort((a, b) => ((b.w || 0) * (b.h || 0)) - ((a.w || 0) * (a.h || 0)))[0];

if (!slide.mediaURL && !slide.bgColor && bigBgShape && bigBgShape.w > 80 && bigBgShape.h > 80) {
  slide.bgColor = bigBgShape.fillColor;
}

return slide;
}

/**
 * 슬라이드 레이아웃 XML에서 placeholder의 위치/크기를 추출한다.
 * 슬라이드 자체에 위치가 없을 때 레이아웃에서 상속받는 값을 제공한다.
 *
 * @param {Object} relsMap - 슬라이드의 관계 ID 매핑
 * @param {AdmZip} zip     - 열린 ZIP 객체
 * @returns {Object} { idx: {x, y, w, h}, type: {x, y, w, h} } 형태의 매핑
 */
function getPptxLayoutPlaceholders(relsMap, zip) {
  const result = {};
  try {
    const layoutRel = Object.entries(relsMap).find(([, t]) => t.includes("slideLayout"));
    if (!layoutRel) return result;
    const layoutTarget = layoutRel[1].replace(/\\/g, "/");
    const layoutPath   = resolveOpenXmlTargetPath("ppt/slides", layoutTarget);
    const layoutEntry = zipGetEntry(zip, layoutPath);
    if (!layoutEntry) return result;
    const lxml = layoutEntry.getData().toString("utf8");

    const spRe = /<p:sp>([\s\S]*?)<\/p:sp>/g;
    let spM;
    while ((spM = spRe.exec(lxml)) !== null) {
      const sp     = spM[1];
      const phTagM = sp.match(/<p:ph\b([^>]*)>/);
      if (!phTagM) continue;
      const typeM  = phTagM[1].match(/\btype="([^"]+)"/);
      const idxM   = phTagM[1].match(/\bidx="([^"]+)"/);
      const type   = typeM ? typeM[1] : "body";
      const idx    = idxM  ? idxM[1]  : "1";

      const offM = sp.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"/);
      const extM = sp.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
      if (offM && extM) {
        const phData = {
          x: parseInt(offM[1]), y: parseInt(offM[2]),
          w: parseInt(extM[1]), h: parseInt(extM[2])
        };
        result[idx]  = phData;  // 인덱스로 조회
        result[type] = phData;  // 타입으로도 조회 가능
      }
    }
  } catch(e) {}
  return result;
}

/**
 * PPTX에서 추출한 폰트 이름을 앱에서 사용 가능한 폰트 이름으로 변환한다.
 * 시스템 한글 폰트를 구글 웹폰트 이름으로 매핑한다.
 *
 * @param {string} fontFace - PPTX에서 추출한 폰트 이름
 * @returns {string} 앱 폰트 이름 (매핑 없으면 원본 그대로)
 */
function pptxFontToKor(fontFace) {
  if (!fontFace) return "Noto Serif KR";
  const map = {
    // 한글 명조 계열 → Noto Serif KR
    "나눔명조": "Noto Serif KR",   "NanumMyeongjo": "Noto Serif KR",
    "바탕":     "Noto Serif KR",   "Batang":        "Noto Serif KR",
    "궁서":     "Noto Serif KR",   "Gungsuh":       "Noto Serif KR",
    "HY신명조": "Noto Serif KR",   "HY헤드라인M":   "Noto Serif KR",
    // 한글 고딕 계열 → Noto Sans KR
    "나눔고딕":            "Nanum Gothic",
    "NanumGothic":         "Nanum Gothic",
    "나눔손글씨 붓":       "Nanum Pen Script",
    "맑은 고딕":           "Noto Sans KR",   "Malgun Gothic": "Noto Sans KR",
    "굴림":                "Noto Sans KR",   "Gulim":         "Noto Sans KR",
    "돋움":                "Noto Sans KR",   "Dotum":         "Noto Sans KR",
    "HY고딕B":             "Noto Sans KR",   "HY강B":         "Noto Sans KR",
    // macOS 기본 한글 폰트
    "Apple SD Gothic Neo": "Noto Sans KR",
    "AppleGothic":         "Noto Sans KR",
    "Apple Myungjo":       "Noto Serif KR",
    // 영문 폰트 → 한글 폰트로 대체
    "Calibri":         "Noto Sans KR",   "Calibri Light":   "Noto Sans KR",
    "Arial":           "Noto Sans KR",   "Arial Narrow":    "Noto Sans KR",
    "Helvetica":       "Noto Sans KR",   "Helvetica Neue":  "Noto Sans KR",
    "Times New Roman": "Noto Serif KR",
    "Georgia":         "Noto Serif KR",
  };
  return map[fontFace] || fontFace;  // 매핑 없으면 원본 유지
}
