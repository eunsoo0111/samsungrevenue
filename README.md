# samsungrevenue
삼성전기(009150) 매출을 추정하기 위한 저장소 (SMIC 54기 3팀 매출교육 과제)

- 결과물: [`삼성전기_매출추정.xlsx`](삼성전기_매출추정.xlsx)
- 정리 노트: [`docs/매출추정_정리.md`](docs/매출추정_정리.md) — raw data 출처, 컨센, breakdown, 라인별 추정 논리
- 기준일: 2026-11-10 (3Q26 잠정실적·IR 반영, 3Q26 분기보고서는 미공시)

## 워크북 구성
| 시트 | 내용 |
|---|---|
| 삼성전기_REV | 최종 매출 추정 Table (2023~2028E, 1Q26~4Q26E), 컨센 대비, 환율 민감도 |
| Breakdown_논리 | 과제 2·3번: 라인별 쪼갠 근거 / Q·P 드라이버 / 핵심 가정 / 크로스체크 |
| 매출추정_컴포넌트·패키지·광학 | 가정(파란 글씨) → 분기 USD 매출 → P×Q×FX → 최종표 (수식 연동) |
| 컨센 | FnGuide 컨센(네이버 제공), 증권사 5곳 부문별 추정(10/29) |
| FX | USD/KRW 분기 평균, 추정 환율 가정 |
| R_IS / R_REV / R_CAPA / R_공시 | DART 원자료: 손익, 부문·수출·지역·고객 매출, Capa·가동률·ASP·Capex, 2026 공급계약·투자 공시 |

## 재현
```bash
python scripts/build_workbook.py data/raw_data.json 삼성전기_매출추정.xlsx
```
`data/raw_data.json`은 DART 웹 뷰어에서 받은 보고서를 `scripts/fetch_dart.py` → `parse_reports.py`·`parse_is.py` → `assemble_data.py` 순서로 가공한 결과다.
