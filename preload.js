/**
 * preload.js — Electron 프리로드 스크립트
 *
 * 역할: Electron의 보안 모델(contextIsolation)을 지키면서
 *       메인 프로세스(main.js)와 렌더러(index.html/screen.html)를
 *       안전하게 연결하는 다리 역할을 한다.
 *
 * 동작 원리:
 *   - contextBridge.exposeInMainWorld(이름, 객체) 로
 *     window.이름 = 객체 형태로 렌더러에서 접근 가능하게 만든다.
 *   - ipcRenderer.invoke / ipcRenderer.send 로 main.js에 요청을 보낸다.
 *   - 렌더러는 Node.js API에 직접 접근할 수 없고,
 *     여기서 노출한 API만 사용할 수 있으므로 보안이 유지된다.
 */

const { contextBridge, ipcRenderer } = require("electron");

const _listenerRegistry = new Map();

function registerSingleListener(channel, cb, mapper) {
  if (typeof cb !== "function") return () => {};
  const existing = _listenerRegistry.get(channel);
  if (existing) ipcRenderer.removeListener(channel, existing);
  const listener = (event, payload) => cb(mapper ? mapper(event, payload) : payload);
  _listenerRegistry.set(channel, listener);
  ipcRenderer.on(channel, listener);
  return () => {
    const current = _listenerRegistry.get(channel);
    if (current === listener) {
      ipcRenderer.removeListener(channel, listener);
      _listenerRegistry.delete(channel);
    }
  };
}

function unregisterSingleListener(channel) {
  const existing = _listenerRegistry.get(channel);
  if (existing) {
    ipcRenderer.removeListener(channel, existing);
    _listenerRegistry.delete(channel);
  }
}

// ──────────────────────────────────────────────────────────
// Gemini API (하위 호환용)
// AI 제공자를 Gemini로 고정하던 구버전 코드와의 호환을 위해 남겨둔다.
// 신규 코드는 아래 aiAPI를 사용한다.
// ──────────────────────────────────────────────────────────
contextBridge.exposeInMainWorld("geminiAPI", {
  // 가사를 AI로 분리해서 슬라이드 배열로 반환
  splitLyrics: (lyrics) => ipcRenderer.invoke("gemini:split-lyrics", lyrics),
  // 가사를 분리하면서 지정한 언어(en/ja)로 번역도 함께 반환
  splitLyricsWithTranslations: (lyrics, options) =>
    ipcRenderer.invoke("ai:split-lyrics-with-translations", lyrics, options)
});

// ──────────────────────────────────────────────────────────
// AI API 라우터
// 설정에서 선택한 AI 제공자(Gemini / OpenAI / Claude)로
// 자동으로 요청을 전달한다.
// ──────────────────────────────────────────────────────────
contextBridge.exposeInMainWorld("aiAPI", {
  // 가사 → 슬라이드 배열 변환 (번역 없음)
  splitLyrics: (lyrics) => ipcRenderer.invoke("ai:split-lyrics", lyrics),
  // 가사 → 슬라이드 배열 + 번역 텍스트박스 포함
  splitLyricsWithTranslations: (lyrics, options) =>
    ipcRenderer.invoke("ai:split-lyrics-with-translations", lyrics, options),
  // 한국어/영어/일본어를 동시에 슬라이드로 분할
  // genEn/genJa가 true이면 입력이 없어도 AI가 번역 생성
  splitMultiLang: (kr, en, ja, genEn, genJa) =>
    ipcRenderer.invoke("ai:split-multilang", kr, en, ja, genEn, genJa)
});

// ──────────────────────────────────────────────────────────
// 설정 API
// API 키, AI 제공자, 모델 등을 로컬 파일(settings.json)에 저장/불러오기
// ──────────────────────────────────────────────────────────
contextBridge.exposeInMainWorld("settingsAPI", {
  // 현재 저장된 설정 조회 (API 키 원문은 포함하지 않고 저장 여부만 반환)
  getSettings: () => ipcRenderer.invoke("settings:get"),
  // AI 제공자·모델·API 키 통합 저장
  setAISettings: (settings) => ipcRenderer.invoke("settings:set-ai-settings", settings),
});

// ──────────────────────────────────────────────────────────
// Electron 창 제어 및 IPC 통신 API
// 메인 조작창(index.html)과 외부 송출 화면(screen.html)을
// 열고 닫거나 데이터를 주고받는다.
// ──────────────────────────────────────────────────────────
contextBridge.exposeInMainWorld("electronAPI", {
  // 외부 모니터(또는 현재 모니터)에 송출 화면 열기
  openSecondScreen: () => ipcRenderer.invoke("open-second-screen"),
  // 외부 송출 화면 닫기
  closeSecondScreen: () => ipcRenderer.invoke("close-second-screen"),

  // 현재 슬라이드 데이터를 송출 화면으로 전달 (두 이름 모두 동일 채널)
  sendSlide: (data) => ipcRenderer.send("slide-update", data),
  sendSlideUpdate: (data) => ipcRenderer.send("slide-update", data),

  // 화면 블랙아웃(on=true) / 해제(on=false)
  blankScreen: (on) => ipcRenderer.send("blank-screen", on),

  // 자동재생 진행 바 상태 전달 (action: start/stop/tick, secs: 남은 시간)
  sendProgressBar: (data) => ipcRenderer.send("progress-bar", data),

  // 송출 화면이 닫혔을 때 콜백 등록
  onSecondScreenClosed: (cb) =>
    registerSingleListener("second-screen-closed", cb, () => undefined),

  // 슬라이드 데이터 수신 콜백 등록 (screen.html에서 사용)
  onSlideUpdate: (cb) =>
    registerSingleListener("slide-update", cb),
  // 블랙아웃 상태 수신 콜백 등록
  onBlankScreen: (cb) =>
    registerSingleListener("blank-screen", cb),
  // 진행 바 상태 수신 콜백 등록
  onProgressBar: (cb) =>
    registerSingleListener("progress-bar", cb),
  removeSlideUpdateListener: () => unregisterSingleListener("slide-update"),
  removeBlankScreenListener: () => unregisterSingleListener("blank-screen"),
  removeProgressBarListener: () => unregisterSingleListener("progress-bar"),
  removeSecondScreenClosedListener: () => unregisterSingleListener("second-screen-closed"),

  // 외부 URL 또는 시스템 설정 열기
  openExternal: (url) => ipcRenderer.invoke("shell:open-external", url),

  // 송출 화면에서 키 입력 시 메인 조작창으로 전달 (next/prev/toggleAuto 등)
  sendOutputControl: (action) => ipcRenderer.send("output-control", action),
  // 메인 조작창에서 output-control 이벤트 수신 콜백 등록
  onOutputControl: (callback) =>
    registerSingleListener("output-control", callback),
  removeOutputControlListener: () => unregisterSingleListener("output-control"),

  // 위치 마커 선: { visible, y } 형태로 송출 화면에 전달
  sendMarkerLine: (data) => ipcRenderer.send("marker-line", data),
  onMarkerLine: (cb) => registerSingleListener("marker-line", cb),
  removeMarkerLineListener: () => unregisterSingleListener("marker-line"),
});

// ──────────────────────────────────────────────────────────
// PPTX 내보내기 / 가져오기 API
// ──────────────────────────────────────────────────────────

// 이미지 변환 진행률 콜백을 저장하는 내부 변수
let _pptImageProgressCb = null;
let _pptImageProgressListener = null;
let _pptImageChunkCb = null;
let _pptImageChunkListener = null;

contextBridge.exposeInMainWorld("pptxAPI", {
  // 현재 그룹 슬라이드를 .pptx 파일로 내보내기
  exportPPTX: (data) => ipcRenderer.invoke("pptx:export", data),
  // .pptx 파일을 앱 슬라이드로 가져오기 (텍스트/이미지 추출)
  importPPTX: () => ipcRenderer.invoke("pptx:import"),
  selectPPTXFile: () => ipcRenderer.invoke("pptx:select-file"),
  // .pptx 파일의 각 슬라이드를 PNG 이미지로 변환해 가져오기 (Electron 렌더링)
  importAsImagesPPTX: () => ipcRenderer.invoke("pptx:importAsImages"),
  // .pptx를 PowerPoint/LibreOffice로 실제 렌더링해 PNG로 가져오기 (원본 디자인 유지)
  renderToImages: (opts) => ipcRenderer.invoke("pptx:render-to-images", opts),
  // PowerPoint 자동화 권한 사전 확인 (macOS 전용, 렌더링 전 미리 호출)
  checkPptPermission: () => ipcRenderer.invoke("pptx:check-ppt-permission"),

  // 이미지 변환 진행률 수신 콜백 등록
  // 변환 중 { current: 현재 슬라이드 번호, total: 전체 수 } 형태로 수신
  onImageProgress: (cb) => {
    _pptImageProgressCb = cb;
    if (_pptImageProgressListener) {
      ipcRenderer.removeListener("pptx:imageProgress", _pptImageProgressListener);
    }
    _pptImageProgressListener = (_, data) => {
      if (_pptImageProgressCb) _pptImageProgressCb(data);
    };
    ipcRenderer.on("pptx:imageProgress", _pptImageProgressListener);
  },
  // 이미지 변환 진행률 수신 콜백 해제
  removeImageProgress: () => {
    _pptImageProgressCb = null;
    if (_pptImageProgressListener) {
      ipcRenderer.removeListener("pptx:imageProgress", _pptImageProgressListener);
      _pptImageProgressListener = null;
    }
  },
  onImageChunk: (cb) => {
    _pptImageChunkCb = cb;
    if (_pptImageChunkListener) {
      ipcRenderer.removeListener("pptx:imageChunk", _pptImageChunkListener);
    }
    _pptImageChunkListener = (_, data) => {
      if (_pptImageChunkCb) _pptImageChunkCb(data);
    };
    ipcRenderer.on("pptx:imageChunk", _pptImageChunkListener);
  },
  removeImageChunk: () => {
    _pptImageChunkCb = null;
    if (_pptImageChunkListener) {
      ipcRenderer.removeListener("pptx:imageChunk", _pptImageChunkListener);
      _pptImageChunkListener = null;
    }
  },
});
