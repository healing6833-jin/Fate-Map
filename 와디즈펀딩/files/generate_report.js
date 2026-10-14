/**
 * 훼이맵 상담 리포트 자동 생성기
 * ------------------------------------------
 * 사용법: 아래 REPORT_DATA 객체만 내담자별로 수정한 뒤 실행하세요.
 *   node generate_report.js
 * 출력: ./output_report.docx
 */

const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow,
  TableCell, WidthType, ShadingType, BorderStyle, AlignmentType,
  Header, Footer, PageNumber, NumberFormat
} = require("docx");
const fs = require("fs");

// ============ 여기만 내담자별로 수정하세요 ============
const REPORT_DATA = {
  brand: "훼이맵",
  clientName: "조영진",
  clientId: "C001",
  birthInfo: "1971년 5월 17일(양력) 오전 5시경",
  sajuPillars: { 년주: "辛亥", 월주: "癸巳", 일주: "壬寅", 시주: "壬寅" },
  ilgan: "壬水(임수)",
  oheng: { 木: 2, 火: 1, 土: 0, 金: 1, 水: 3 },
  sessionDate: "2026년 1월 15일",
  sessionNumber: 1,
  topic: "전체 운세 및 훼이맵 사업 수익화 시점 상담",
  keyPoints: [
    "타고난 성향: '은둔하는 건설가' — 겉은 견고(토성), 속은 예민(12하우스 금성·수성)",
    "생애 흐름: 30-40대 고난기를 지나 50대 이후 결실을 맺는 토성 사이클",
    "천직 방향: 북노드 10하우스 — 치유자·교육자·상담자로서 소수에게 깊은 영향을 주는 역할",
    "사업 시점: 2025-2028년은 토성이 12하우스(양자리)를 통과하는 기반 구축기",
    "실질 수익 확립기: 2028-2030년 토성 1하우스 진입 및 토성 리턴(58-59세)",
    "협업의 중요성: 7하우스 궁수자리 목성 강세 → 파트너십이 조기 수익의 핵심 열쇠"
  ],
  prescription: {
    name: "대지의 뿌리 블렌드 (토성/土 보완용)",
    recipe: "베티버 3방울 + 패츌리 2방울 + 시더우드 3방울 + 프랑킨센스 2방울 (호호바 오일 30ml)",
    usage: "총 10방울 기준 약 1.67% 농도 — 일반 마사지에 적합. 아침 명상 또는 업무 시작 전 손목·명치에 소량 도포"
  },
  nextSteps: [
    "콘텐츠·상담 사례 체계화 착수",
    "관련 자격 강화 방안 검토",
    "7하우스 목성 기반 협업 파트너 후보 리스트업",
    "행정·재무 관리 시스템 별도 구축 (토 부재 보완)"
  ],
  nextSessionDate: "2026년 4월 15일",
  counselorNote: "대규모 투자나 성급한 대중 노출은 지양하고, 기반을 다지는 데 집중할 시기입니다."
};
// ======================================================

const COLOR_PRIMARY = "5B4636"; // 브랜드 톤 (흙빛/우드톤)
const COLOR_ACCENT = "8A6D3B";
const COLOR_GREY = "6B7280";

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
    (t) =>
      new Paragraph({
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
        children: [new Paragraph({ children: [new TextRun({ text: value, size: 20, font: "Malgun Gothic" })] })]
      })
    ]
  });
}

const ohengOrder = ["木", "火", "土", "金", "水"];
const ohengRow = new TableRow({
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
          children: ohengOrder.flatMap((k, i) => {
            const val = REPORT_DATA.oheng[k];
            const runs = [
              new TextRun({ text: `${k} `, bold: true, size: 20, font: "Malgun Gothic" }),
              new TextRun({ text: `${val}${val === 0 ? " (부재)" : ""}`, size: 20, font: "Malgun Gothic", color: val === 0 ? "B91C1C" : "000000" })
            ];
            if (i < ohengOrder.length - 1) runs.push(new TextRun({ text: "   /   ", size: 20, color: COLOR_GREY }));
            return runs;
          })
        })
      ]
    })
  ]
});

const pillarsText = Object.entries(REPORT_DATA.sajuPillars)
  .map(([k, v]) => `${k} ${v}`)
  .join("   ");

const doc = new Document({
  sections: [
    {
      properties: {
        page: {
          size: { width: 11907, height: 16840 }, // A4
          margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 }
        }
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({ text: REPORT_DATA.brand, size: 18, color: COLOR_GREY, font: "Malgun Gothic" })]
            })
          ]
        })
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "본 리포트는 개인 상담을 위한 참고 자료이며, 의학적·법적 조언을 대신하지 않습니다.", size: 16, color: COLOR_GREY, font: "Malgun Gothic" })
              ]
            })
          ]
        })
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 60 },
          children: [new TextRun({ text: `${REPORT_DATA.brand} 상담 리포트`, bold: true, size: 40, color: COLOR_PRIMARY, font: "Malgun Gothic" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 300 },
          children: [new TextRun({ text: `제 ${REPORT_DATA.sessionNumber}회 상담  |  ${REPORT_DATA.sessionDate}`, size: 20, color: COLOR_GREY, font: "Malgun Gothic" })]
        }),

        heading("내담자 정보"),
        new Table({
          width: { size: 9000, type: WidthType.DXA },
          columnWidths: [2200, 6800],
          rows: [
            infoRow("이름", `${REPORT_DATA.clientName} (${REPORT_DATA.clientId})`),
            infoRow("생년월일시", REPORT_DATA.birthInfo),
            infoRow("사주팔자", pillarsText),
            infoRow("일간", REPORT_DATA.ilgan),
            ohengRow
          ]
        }),

        heading("이번 상담 주제"),
        bodyText(REPORT_DATA.topic, { bold: true, color: COLOR_PRIMARY }),

        heading("핵심 분석 내용"),
        ...bulletList(REPORT_DATA.keyPoints),

        heading("처방"),
        bodyText(REPORT_DATA.prescription.name, { bold: true }),
        bodyText(`레시피: ${REPORT_DATA.prescription.recipe}`),
        bodyText(`사용법: ${REPORT_DATA.prescription.usage}`, { color: COLOR_GREY }),

        heading("다음 단계 / 과제"),
        ...bulletList(REPORT_DATA.nextSteps),

        heading("상담자 코멘트"),
        bodyText(REPORT_DATA.counselorNote),

        new Paragraph({ spacing: { before: 300 }, border: { top: { color: "D0D5DD", space: 4, style: BorderStyle.SINGLE, size: 4 } }, children: [] }),
        bodyText(`다음 상담 예정일: ${REPORT_DATA.nextSessionDate}`, { bold: true, color: COLOR_PRIMARY })
      ]
    }
  ]
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("/home/claude/output_report.docx", buffer);
  console.log("saved");
});
