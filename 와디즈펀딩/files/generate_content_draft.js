/**
 * 훼이맵 콘텐츠 초안 자동 생성기
 * ------------------------------------------------------
 * 상담이력 시트를 태그로 필터링하여 블로그 포스트 "초안 골격"을
 * 마크다운으로 자동 생성합니다. (제목 후보, 도입부 훅, 본문 구조,
 * 처방 섹션까지 뼈대를 잡아주며, 최종 다듬기는 사람이 합니다.)
 *
 * 사용법:
 *   node generate_content_draft.js [태그1] [태그2] ...
 *   예) node generate_content_draft.js TP-사업
 *       node generate_content_draft.js OH-水 HS-07
 *   태그 없이 실행하면 전체 세션을 대상으로 합니다.
 *
 * 출력: ./content_drafts/ 폴더에 세션당 .md 파일 하나씩 생성
 */

const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const excelPath = "훼이맵_내담자관리.xlsx";
const outDir = "./content_drafts";
const filterTags = process.argv.slice(2); // CLI로 받은 태그들

if (!fs.existsSync(excelPath)) {
  console.error(`엑셀 파일을 찾을 수 없습니다: ${excelPath}`);
  process.exit(1);
}
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

const wb = XLSX.readFile(excelPath);
const sessions = XLSX.utils.sheet_to_json(wb.Sheets["상담이력"], { range: 3, defval: "" })
  .filter((r) => String(r["고객ID"]).trim() !== "");

const TAG_COLS = ["오행태그", "하우스태그", "주제태그", "처방태그", "형식태그"];

function matchesFilter(session) {
  if (filterTags.length === 0) return true;
  const sessionTags = TAG_COLS.map((c) => String(session[c] || "").trim());
  return filterTags.every((t) => sessionTags.includes(t));
}

function splitLines(text) {
  if (!text) return [];
  return String(text)
    .split(/\n|,\s*(?=[가-힣A-Za-z])/)
    .map((s) => s.trim())
    .filter(Boolean);
}

// 태그 → 후크 문구 매핑 (제목/도입부 톤 힌트)
const HOOK_BY_TAG = {
  "TP-사업": "일과 방향에 대한 고민",
  "TP-관계": "관계에서 반복되는 패턴",
  "TP-건강": "몸과 마음의 신호",
  "TP-시기운": "지금 이 시기가 갖는 의미",
  "TP-자기이해": "나도 몰랐던 나의 기질",
  "TP-위기관리": "흔들리는 시기를 지나는 법",
  "TP-영성성장": "내면 깊은 곳의 목소리",
};

function buildDraft(session) {
  const tags = TAG_COLS.map((c) => session[c]).filter(Boolean);
  const topicTag = tags.find((t) => t.startsWith("TP-")) || "";
  const hook = HOOK_BY_TAG[topicTag] || "삶의 흐름을 읽는 법";
  const keyPoints = splitLines(session["핵심내용"]);
  const nextSteps = splitLines(session["다음단계/과제"]);

  const titleCandidates = [
    `${hook}, 사주는 어떻게 답할까`,
    `${session["상담주제"]} — 흐름을 읽는 시선`,
    `왜 지금 이 시기가 중요한가: ${session["상담주제"]}`,
  ];

  const md = `---
draft: true
source_session: "${session["고객ID"]}_${session["상담회차"]}회차_${session["상담일자"]}"
tags: [${tags.join(", ")}]
status: "AI 초안 - 검수 필요"
---

# [제목 후보 - 택1 또는 재작성]
${titleCandidates.map((t) => `- ${t}`).join("\n")}

## 도입부 (훅)
> [ TODO: "${hook}"을 소재로 독자의 공감을 끄는 1~2문단 작성.
> 예: 이런 고민을 해보신 적 있나요? 로 시작하거나, 짧은 일화로 시작 ]

## 본문 1 — 무엇이 보이는가
${keyPoints.length ? keyPoints.map((p) => `- ${p}`).join("\n") : "- [핵심 내용 없음 - 상담이력에 내용 추가 필요]"}

[ TODO: 위 포인트들을 자연스러운 문단으로 풀어 쓰기. 전문용어(하우스, 오행 등)는
       독자가 이해할 수 있도록 짧게 풀이 추가 ]

## 본문 2 — 처방과 실천
${session["처방(아로마 등)"] ? `현재 제안하는 처방: **${session["처방(아로마 등)"]}**` : "[ 처방 정보 없음 ]"}

[ TODO: 이 처방이 왜 이 사람/상황에 맞는지 1문단으로 설명 ]

## 다음 단계
${nextSteps.length ? nextSteps.map((s, i) => `${i + 1}. ${s}`).join("\n") : "- [다음 단계 없음]"}

## 마무리
[ TODO: 독자에게 남기는 메시지 또는 행동 유도 문장 1~2문장 ]

---
*(참고용 내부 메타) 원본 상담: ${session["상담주제"]} / ${session["상담일자"]} / 태그: ${tags.join(", ")}*
`;
  return md;
}

let count = 0;
for (const session of sessions) {
  if (!matchesFilter(session)) continue;
  const md = buildDraft(session);
  const fname = `draft_${session["고객ID"]}_${session["상담회차"]}회차_${session["상담일자"]}.md`;
  fs.writeFileSync(path.join(outDir, fname), md, "utf-8");
  console.log(`초안 생성됨: ${fname}`);
  count++;
}

console.log(`\n총 ${count}개 초안 생성 완료 → ${outDir}`);
if (filterTags.length) console.log(`(필터: ${filterTags.join(", ")})`);
