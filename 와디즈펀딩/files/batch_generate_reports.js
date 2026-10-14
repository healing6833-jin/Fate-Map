/**
 * 훼이맵 상담 리포트 "일괄" 자동 생성기
 * ------------------------------------------------------
 * 훼이맵_내담자관리.xlsx 파일의
 *   - 내담자DB 시트 (사주/오행 정보)
 *   - 상담이력 시트 (세션별 상담 기록)
 * 두 시트를 고객ID로 연결하여, 상담이력의 행 하나당
 * 리포트 파일 하나를 자동 생성합니다.
 *
 * 사용법:
 *   node batch_generate_reports.js [엑셀파일경로] [출력폴더]
 *   기본값: 훼이맵_내담자관리.xlsx → ./reports
 *
 * 새 상담을 추가하려면 엑셀의 "상담이력" 시트에 행만 추가하고
 * 이 스크립트를 다시 실행하면 됩니다 (기존 파일은 자동 재생성/덮어쓰기).
 */

const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow,
  TableCell, WidthType, ShadingType, BorderStyle, AlignmentType,
  Header, Footer
} = require("docx");

const BRAND = "훼이맵";
const COLOR_PRIMARY = "5B4636";
const COLOR_ACCENT = "8A6D3B";
const COLOR_GREY = "6B7280";

const excelPath = process.argv[2] || "훼이맵_내담자관리.xlsx";
const outDir = process.argv[3] || "./reports";

if (!fs.existsSync(excelPath)) {
  console.error(`엑셀 파일을 찾을 수 없습니다: ${excelPath}`);
  process.exit(1);
}
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

// ---------- 1. 엑셀 읽기 ----------
const wb = XLSX.readFile(excelPath);
const clientRows = XLSX.utils.sheet_to_json(wb.Sheets["내담자DB"], { range: 3, defval: "" })
  .filter((r) => String(r["고객ID"]).trim() !== "");
const sessionRows = XLSX.utils.sheet_to_json(wb.Sheets["상담이력"], { range: 3, defval: "" })
  .filter((r) => String(r["고객ID"]).trim() !== "");

const clientById = Object.fromEntries(clientRows.map((c) => [String(c["고객ID"]).trim(), c]));

function splitLines(text) {
  // 셀 안에 줄바꿈이나 " / ", ", " 로 구분된 여러 항목을 리스트로 변환
  if (!text) return [];
  return String(text)
    .split(/\n|,\s*(?=[가-힣A-Za-z])/)
    .map((s) => s.trim())
    .filter(Boolean);
}

// ---------- 2. 문서 빌더 (기존 스타일 재사용) ----------
function heading(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    border: { bottom: { color: COLOR_ACCENT, space: 4, style: BorderStyle.SINGLE, size: 6 } },
    children: [new TextRun({ text, bold: true, color: COLOR_PRIMARY, size: 26, font: "Malgun Gothic" })]
  });
}
function bodyText(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, size: 22, font: "Malgun Gothic", ...opts })]
  });
}
function bulletList(items) {
  return items.map(
    (t) => new Paragraph({
      spacing: { after: 90 },
      bullet: { level: 0 },
      children: [new TextRun({ text: t, size: 21, font: "Malgun Gothic" })]
    })
  );
}
function infoRow(label, value) {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 2200, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: "F3EFE9" },
        margins: { top: 100, bottom: 100, left: 150, right: 150 },
        children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 20, font: "Malgun Gothic", color: COLOR_PRIMARY })] })]
      }),
      new TableCell({
        width: { size: 6800, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 150, right: 150 },
        children: [new Paragraph({ children: [new TextRun({ text: String(value), size: 20, font: "Malgun Gothic" })] })]
      })
    ]
  });
}

function buildOhengRow(client) {
  const order = ["木", "火", "土", "金", "水"];
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 2200, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: "F3EFE9" },
        margins: { top: 100, bottom: 100, left: 150, right: 150 },
        children: [new Paragraph({ children: [new TextRun({ text: "오행 분포", bold: true, size: 20, font: "Malgun Gothic", color: COLOR_PRIMARY })] })]
      }),
      new TableCell({
        width: { size: 6800, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 150, right: 150 },
        children: [
          new Paragraph({
            children: order.flatMap((k, i) => {
              const val = client[`오행_${k}`];
              const runs = [
                new TextRun({ text: `${k} `, bold: true, size: 20, font: "Malgun Gothic" }),
                new TextRun({ text: `${val}${val === 0 ? " (부재)" : ""}`, size: 20, font: "Malgun Gothic", color: val === 0 ? "B91C1C" : "000000" })
              ];
              if (i < order.length - 1) runs.push(new TextRun({ text: "   /   ", size: 20, color: COLOR_GREY }));
              return runs;
            })
          })
        ]
      })
    ]
  });
}

function buildReportDoc(client, session) {
  const pillarsText = ["년주", "월주", "일주", "시주"].map((k) => `${k} ${client[k]}`).join("   ");
  const keyPoints = splitLines(session["핵심내용"]);
  const nextSteps = splitLines(session["다음단계/과제"]);

  return new Document({
    sections: [
      {
        properties: {
          page: { size: { width: 11907, height: 16840 }, margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } }
        },
        headers: {
          default: new Header({
            children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: BRAND, size: 18, color: COLOR_GREY, font: "Malgun Gothic" })] })]
          })
        },
        footers: {
          default: new Footer({
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "본 리포트는 개인 상담을 위한 참고 자료이며, 의학적·법적 조언을 대신하지 않습니다.", size: 16, color: COLOR_GREY, font: "Malgun Gothic" })] })]
          })
        },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: `${BRAND} 상담 리포트`, bold: true, size: 40, color: COLOR_PRIMARY, font: "Malgun Gothic" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: `제 ${session["상담회차"]}회 상담  |  ${session["상담일자"]}`, size: 20, color: COLOR_GREY, font: "Malgun Gothic" })] }),

          heading("내담자 정보"),
          new Table({
            width: { size: 9000, type: WidthType.DXA },
            columnWidths: [2200, 6800],
            rows: [
              infoRow("이름", `${client["이름"]} (${client["고객ID"]})`),
              infoRow("생년월일시", `${client["생년월일"]} ${client["생시"]} (${client["양/음력"]})`),
              infoRow("사주팔자", pillarsText),
              infoRow("일간", client["일간"]),
              buildOhengRow(client)
            ]
          }),

          heading("이번 상담 주제"),
          bodyText(String(session["상담주제"] || ""), { bold: true, color: COLOR_PRIMARY }),

          heading("핵심 분석 내용"),
          ...(keyPoints.length ? bulletList(keyPoints) : [bodyText("-")]),

          heading("처방"),
          bodyText(String(session["처방(아로마 등)"] || "-")),

          heading("다음 단계 / 과제"),
          ...(nextSteps.length ? bulletList(nextSteps) : [bodyText("-")]),

          new Paragraph({ spacing: { before: 300 }, border: { top: { color: "D0D5DD", space: 4, style: BorderStyle.SINGLE, size: 4 } }, children: [] }),
          bodyText(`다음 상담 예정일: ${session["다음상담예정일"] || "-"}`, { bold: true, color: COLOR_PRIMARY })
        ]
      }
    ]
  });
}

// ---------- 3. 실행: 상담이력 각 행 → 리포트 파일 ----------
(async () => {
  let count = 0;
  const skipped = [];

  for (const session of sessionRows) {
    const cid = String(session["고객ID"]).trim();
    const client = clientById[cid];
    if (!client) {
      skipped.push(`${cid} (${session["이름"]}) - 내담자DB에서 매칭되는 고객ID 없음`);
      continue;
    }
    const doc = buildReportDoc(client, session);
    const buffer = await Packer.toBuffer(doc);
    const fname = `${BRAND}_상담리포트_${client["이름"]}_${session["상담회차"]}회차_${session["상담일자"]}.docx`;
    fs.writeFileSync(path.join(outDir, fname), buffer);
    count++;
    console.log(`생성됨: ${fname}`);
  }

  console.log(`\n총 ${count}개 리포트 생성 완료 → ${outDir}`);
  if (skipped.length) {
    console.log(`\n건너뜀 (${skipped.length}건):`);
    skipped.forEach((s) => console.log(`  - ${s}`));
  }
})();
