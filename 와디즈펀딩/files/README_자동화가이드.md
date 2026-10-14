# 훼이맵 상담 리포트 자동화 가이드

`훼이맵/와디즈펀딩/files` 폴더에 있는 스크립트들을 하나의 파이프라인으로 쓰는 방법입니다.

## 처음 한 번만 하는 준비

1. 이 폴더(`files`)에서 명령 프롬프트(또는 PowerShell)를 열고 패키지를 설치합니다.

   ```
   npm install
   ```

   (`docx`, `xlsx`, `exceljs` 세 패키지가 설치됩니다. `xlsx`, `docx`는 기존 스크립트가 이미 쓰고 있던 것이고, `exceljs`는 이번에 추가된 `ai_fill_sessions.js`가 엑셀 서식/수식을 보존하며 저장하기 위해 씁니다.)

2. Claude API 키를 발급받아 환경변수로 등록합니다. (`ai_fill_sessions.js`에서만 필요합니다.)
   - https://console.anthropic.com 에서 키 발급
   - PowerShell: `$env:ANTHROPIC_API_KEY="sk-ant-..."` (그 창에서만 유지되며, 매번 새 창에서 다시 설정해야 합니다. 영구 등록하려면 Windows 환경변수 설정에 추가하세요.)

## 평소 쓰는 흐름

1. `훼이맵_내담자관리.xlsx`의 "상담이력" 시트에 이번 회차 행을 추가합니다. (고객ID, 이름, 상담일자, 상담회차, 상담주제까지만 채우고 **핵심내용/처방/다음단계는 비워둬도 됩니다.**)
2. 아래 순서로 실행합니다.

   ```
   node ai_fill_sessions.js       (빈 핵심내용/처방/다음단계를 AI가 초안으로 채워 엑셀에 저장)
   node batch_generate_reports.js (상담이력 전체를 기준으로 리포트 docx 재생성 → ./reports)
   node generate_content_draft.js (블로그 초안 md 재생성 → ./content_drafts)
   ```

   또는 `npm run pipeline` 한 번으로 세 단계를 순서대로 실행할 수 있습니다.

3. **반드시 검수하세요.** `ai_fill_sessions.js`가 채운 핵심내용/처방/다음단계는 초안입니다. 엑셀을 열어 내용을 확인·수정한 뒤(필요하면), `batch_generate_reports.js`만 다시 실행해 리포트를 재생성하면 됩니다.

## 각 스크립트 요약

| 스크립트 | 역할 | 입력 | 출력 |
|---|---|---|---|
| `ai_fill_sessions.js` | 빈 상담 기록에 AI 초안(핵심내용/처방/다음단계) 채우기 | `훼이맵_내담자관리.xlsx` | 같은 엑셀 파일에 저장 |
| `batch_generate_reports.js` | 상담이력 전체 → 회차별 상담 리포트 docx 일괄 생성 | `훼이맵_내담자관리.xlsx` | `./reports/*.docx` |
| `generate_content_draft.js` | 상담이력 → 블로그 초안 md 골격 생성 (태그 필터 가능) | `훼이맵_내담자관리.xlsx` | `./content_drafts/*.md` |
| `generate_report.js` | 리포트 1건만 수동으로 만들 때 (REPORT_DATA 직접 수정) | 스크립트 내 하드코딩 데이터 | `./output_report.docx` |

## 안전장치

- `ai_fill_sessions.js`는 **이미 값이 있는 셀은 절대 덮어쓰지 않습니다.** 핵심내용/처방/다음단계 중 비어있는 것만 채웁니다.
- 오행통계 시트의 수식(`=SUM(...)` 등)은 건드리지 않습니다.
- `--dry-run` 옵션을 붙이면 엑셀에 저장하지 않고 콘솔에서 결과만 미리 볼 수 있습니다: `node ai_fill_sessions.js --dry-run`
