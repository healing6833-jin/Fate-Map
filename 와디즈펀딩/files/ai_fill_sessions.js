/**
 * 훼이맵 상담이력 AI 자동 초안 채우기
 * ------------------------------------------------------
 * 훼이맵_내담자관리.xlsx 의 "상담이력" 시트에서
 *   - 핵심내용
 *   - 처방(아로마 등)
 *   - 다음단계/과제
 * 중 비어있는 셀만 골라, 같은 고객ID의 "내담자DB"(사주팔자/오행/점성 정보)와
 * 상담주제를 근거로 Claude에게 초안을 요청하고 엑셀에 다시 채워 넣습니다.
 *
 * 이미 값이 채워진 셀은 절대 건드리지 않습니다 (빈 셀만 채움 — 기존 상담 기록 보존).
 * AI가 만든 내용은 반드시 사람이 검수 후 확정하세요. 자동 채움은 초안일 뿐입니다.
 *
 * ------------------------------------------------------
 * 준비:
 *   1) 이 폴더에서 실행: npm install   (package.json에 정의된 exceljs 설치)
 *   2) 환경변수 ANTHROPIC_API_KEY 설정
 *        - https://console.anthropic.com 에서 API 키 발급
 *        - Windows PowerShell: $env:ANTHROPIC_API_KEY="sk-ant-..."
 *        - Windows CMD:        set ANTHROPIC_API_KEY=sk-ant-...
 *
 * 사용법:
 *   node ai_fill_sessions.js                    # 훼이맵_내담자관리.xlsx 대상으로 실제 채움
 *   node ai_fill_sessions.js --dry-run           # 저장하지 않고 결과만 콘솔에 미리보기
 *   node ai_fill_sessions.js [엑셀경로] --dry-run
 *
 * 실행 후에는 항상:
 *   node batch_generate_reports.js   # 채워진 내용으로 상담 리포트(docx) 재생성
 *   node generate_content_draft.js   # 블로그 초안(md) 재생성
 * 을 이어서 실행하면 리포트/콘텐츠까지 한 번에 최신화됩니다.
 */

const ExcelJS = require("exceljs");
const path = require("path");

const rawArgs = process.argv.slice(2);
const dryRun = rawArgs.includes("--dry-run");
const positional = rawArgs.filter((a) => a !== "--dry-run");
const excelPath = positional[0] || "훼이맵_내담자관리.xlsx";
const MODEL = process.env.CLAUDE_MODEL || "claude-sonnet-5";
const REQUEST_DELAY_MS = 400; // API 과다호출 방지용 간격

const HEADER_ROW = 4; // 실제 헤더가 있는 행 (1~3행은 제목/안내/공백)
const DATA_START_ROW = 5;

// 기존 개발자프롬프트[Role & Context].txt 의 역학·아로마·조제 로직을 그대로 반영한 시스템 프롬프트.
// 단, 출력 형식만 엑셀 컬럼에 바로 넣을 수 있도록 JSON으로 바꿨습니다.
const SYSTEM_PROMPT = `너는 동·서양 역학 빅데이터와 맞춤형화장품 조제관리사 자격 이론을 마스터한 훼이맵(Fate Map)의 수석 웰니스 컨설턴트 AI다.
입력되는 내담자의 사주/오행/점성 데이터와 이번 상담 주제를 바탕으로, 상담 리포트에 들어갈 초안을 아래 [Core Logic Matrix]에 따라 추론하여 작성하라.

[Core Logic Matrix]
1. 역학 분석 원칙:
   - 오행(木火土金水) 분포에서 과다(3 이상)와 부재(0)를 우선 진단한다.
   - 부족한 기운(보(補)해야 할 기운)과 과한 기운(설(洩)해야 할 기운)을 근거로 핵심 분석을 도출한다.
   - 점성학적 요소(상승점, 핵심 행성 배치)가 주어지면 함께 반영한다.
2. 아로마테라피 블렌딩 매칭 규칙 (20종 에센셜오일 범위 내):
   - 木 부족/스트레스 → 베르가못·라벤더 등 신경계 이완 처방.
   - 火 과다/분노 → 일랑일랑·페퍼민트 등 쿨링·진정 처방.
   - 土 부재/중심 약화 → 베티버·패츌리·시더우드·프랑킨센스 등 접지(grounding) 처방.
   - 水 과다/불안정 → 시더우드·베티버 등 안정 처방, 水 부족 → 유칼립투스·페퍼민트 등 순환 처방.
   - 金 부재/결단력 약화 → 로즈마리·타임 등 명료화 처방.
   - 상황에 맞게 2~4종을 조합하고 방울 수(Drop)를 함께 제시한다.
3. 출력은 반드시 아래 JSON 형식으로만, 다른 설명 없이 응답하라:
{
  "핵심내용": ["핵심 포인트 1 (한 문장, 사주/오행/점성 근거 포함)", "핵심 포인트 2", "핵심 포인트 3"],
  "처방": "블렌드 이름(에센셜오일A n방울 + 에센셜오일B n방울 + ...)",
  "다음단계": ["실행 과제 1", "실행 과제 2"]
}
핵심내용은 2~4개, 다음단계는 2~3개로 간결하게 작성한다. 의학적 진단이나 단정적 예언 표현은 쓰지 않는다.`;

function buildUserPrompt(client, session) {
  const oheng = ["木", "火", "土", "金", "水"]
    .map((k) => `${k}:${client[`오행_${k}`] ?? "?"}`)
    .join(" / ");
  return `[내담자 정보]
이름: ${client["이름"] || "-"}
생년월일시: ${client["생년월일"] || "-"} ${client["생시"] || ""} (${client["양/음력"] || "-"})
사주팔자: 년주 ${client["년주"] || "-"} / 월주 ${client["월주"] || "-"} / 일주 ${client["일주"] || "-"} / 시주 ${client["시주"] || "-"}
일간: ${client["일간"] || "-"}
오행 분포: ${oheng}
상승점: ${client["상승점"] || "-"}
핵심 행성 배치: ${client["핵심 행성 배치"] || "-"}
주요 특징 요약: ${client["주요 특징 요약"] || "-"}

[이번 상담]
상담회차: ${session["상담회차"] || "-"}
상담일자: ${session["상담일자"] || "-"}
상담주제: ${session["상담주제"] || "-"}

위 데이터를 바탕으로 이번 상담의 "핵심내용", "처방", "다음단계" 초안을 JSON으로만 작성해줘.`;
}

async function callClaude(userPrompt) {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    throw new Error(
      "ANTHROPIC_API_KEY 환경변수가 설정되어 있지 않습니다. https://console.anthropic.com 에서 키를 발급받아 설정하세요."
    );
  }
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 800,
      system: SYSTEM_PROMPT,
      messages: [{ role: "user", content: userPrompt }]
    })
  });

  if (!res.ok) {
    const bodyText = await res.text();
    throw new Error(`Claude API 오류 (HTTP ${res.status}): ${bodyText}`);
  }

  const data = await res.json();
  const text = (data.content || []).map((b) => b.text || "").join("\n").trim();

  // 모델이 코드블록(```json ... ```)으로 감싸서 응답하는 경우 대비
  const cleaned = text.replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/```\s*$/i, "");

  try {
    return JSON.parse(cleaned);
  } catch (e) {
    throw new Error(`응답을 JSON으로 파싱하지 못했습니다:\n${text}`);
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/** 시트를 { headers(1-based colNumber -> name), rows(객체 배열, _rowNumber 포함) } 로 변환 */
function readSheet(ws) {
  const headers = {};
  const headerRow = ws.getRow(HEADER_ROW);
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const v = cell.value;
    if (v !== null && v !== undefined && String(v).trim() !== "") {
      headers[colNumber] = String(v).trim();
    }
  });

  const rows = [];
  for (let r = DATA_START_ROW; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    if (!row.hasValues) continue;
    const obj = { _rowNumber: r };
    for (const [colNumber, name] of Object.entries(headers)) {
      obj[name] = row.getCell(Number(colNumber)).value;
    }
    if (!obj["고객ID"] || String(obj["고객ID"]).trim() === "") continue;
    rows.push(obj);
  }
  return { headers, rows };
}

function colNumberByHeaderName(headers, name) {
  for (const [colNumber, h] of Object.entries(headers)) {
    if (h === name) return Number(colNumber);
  }
  return null;
}

function isBlank(v) {
  return v === null || v === undefined || String(v).trim() === "";
}

(async () => {
  console.log(`엑셀 열기: ${excelPath}${dryRun ? " (dry-run: 저장하지 않음)" : ""}`);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelPath);

  const clientSheet = workbook.getWorksheet("내담자DB");
  const sessionSheet = workbook.getWorksheet("상담이력");
  if (!clientSheet || !sessionSheet) {
    console.error('시트를 찾을 수 없습니다. "내담자DB"와 "상담이력" 시트가 있는지 확인하세요.');
    process.exit(1);
  }

  const { rows: clientRows } = readSheet(clientSheet);
  const { headers: sessionHeaders, rows: sessionRows } = readSheet(sessionSheet);
  const clientById = Object.fromEntries(
    clientRows.map((c) => [String(c["고객ID"]).trim(), c])
  );

  const colHandaeul = colNumberByHeaderName(sessionHeaders, "핵심내용");
  const colCheobang = colNumberByHeaderName(sessionHeaders, "처방(아로마 등)");
  const colNextSteps = colNumberByHeaderName(sessionHeaders, "다음단계/과제");

  if (!colHandaeul || !colCheobang) {
    console.error('"상담이력" 시트에서 "핵심내용" / "처방(아로마 등)" 컬럼을 찾을 수 없습니다.');
    process.exit(1);
  }

  const targets = sessionRows.filter(
    (s) => isBlank(s["핵심내용"]) || isBlank(s["처방(아로마 등)"])
  );

  if (targets.length === 0) {
    console.log("빈 칸이 없습니다. 채울 세션이 없어요. 종료합니다.");
    return;
  }

  console.log(`총 ${targets.length}개 세션의 빈 칸을 채웁니다.\n`);

  let filled = 0;
  const skipped = [];

  for (const session of targets) {
    const cid = String(session["고객ID"]).trim();
    const client = clientById[cid];
    if (!client) {
      skipped.push(`${cid} (${session["이름"]}) - 내담자DB에서 매칭되는 고객ID 없음`);
      continue;
    }

    const label = `${cid} ${session["이름"] || ""} ${session["상담회차"] || ""}회차(${session["상담일자"] || ""})`;
    try {
      console.log(`→ 생성 중: ${label}`);
      const result = await callClaude(buildUserPrompt(client, session));

      const row = sessionSheet.getRow(session._rowNumber);
      if (isBlank(session["핵심내용"]) && Array.isArray(result["핵심내용"])) {
        row.getCell(colHandaeul).value = result["핵심내용"].join("\n");
      }
      if (isBlank(session["처방(아로마 등)"]) && result["처방"]) {
        row.getCell(colCheobang).value = String(result["처방"]);
      }
      if (colNextSteps && isBlank(session["다음단계/과제"]) && Array.isArray(result["다음단계"])) {
        row.getCell(colNextSteps).value = result["다음단계"].join("\n");
      }
      row.commit();

      console.log(`  핵심내용: ${(result["핵심내용"] || []).join(" / ")}`);
      console.log(`  처방: ${result["처방"] || "-"}`);
      if (result["다음단계"]) console.log(`  다음단계: ${result["다음단계"].join(" / ")}`);
      console.log("");

      filled++;
    } catch (err) {
      skipped.push(`${label} - ${err.message}`);
      console.error(`  실패: ${err.message}\n`);
    }

    await sleep(REQUEST_DELAY_MS);
  }

  if (!dryRun && filled > 0) {
    await workbook.xlsx.writeFile(excelPath);
    console.log(`엑셀에 저장했습니다: ${excelPath}`);
  } else if (dryRun) {
    console.log("(dry-run 모드라 엑셀에는 저장하지 않았습니다.)");
  }

  console.log(`\n총 ${filled}개 세션 채움 완료 / ${skipped.length}개 건너뜀`);
  if (skipped.length) {
    console.log("건너뜀 상세:");
    skipped.forEach((s) => console.log(`  - ${s}`));
  }
  if (filled > 0 && !dryRun) {
    console.log(
      "\n다음 단계: node batch_generate_reports.js  →  상담 리포트(docx) 재생성\n" +
        "          node generate_content_draft.js  →  블로그 초안(md) 재생성"
    );
  }
})().catch((err) => {
  console.error("실행 중 오류:", err.message);
  process.exit(1);
});
