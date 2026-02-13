# SpendTagger (macOS)

은행/카드 지출내역 엑셀 파일을 통합하고, 카테고리/세부 카테고리를 태깅한 뒤 분석 차트를 생성하는 데스크탑 앱입니다.

## 주요 기능

- 여러 엑셀 파일(`.xls`, `.xlsx`) 통합
- 컬럼 자동 인식 + 인식 실패 시 사용자 매핑 팝업
- 카테고리/세부 카테고리 자동 태깅(룰 기반 + 선택적 GPT)
- 불확실 항목 수동 검토
- 온톨로지(카테고리/세부 카테고리) 추가/삭제 편집
- 결과물 생성
  - 통합 지출내역 엑셀
  - 월별 총지출 엑셀
  - 카테고리/세부카테고리/월별 파이차트

## 프로젝트 파일

- `spending_tagger_app.py`: 메인 macOS GUI 앱
- `merge_transactions.py`: 은행/카드 엑셀 통합 및 정규화 로직
- `analyze_spending_charts.py`: 카테고리/세부카테고리 차트 스크립트
- `analyze_monthly_category_pies.py`: 월별 카테고리 차트 스크립트
- `analyze_overall_subcategory_pie.py`: 전체 세부카테고리 차트 스크립트
- `export_monthly_spending.py`: 월별 총지출 엑셀 생성 스크립트
- `ontology.json`: 기본 온톨로지
- `requirements_macos_app.txt`: 실행 의존성

## 빠른 시작

```bash
conda activate tmp
pip install -r requirements_macos_app.txt
python spending_tagger_app.py
```

## 더블클릭 실행(macOS)

- `run_spend_tagger.command`를 더블클릭하면 앱이 실행됩니다.
- 우선순위:
  1. `~/anaconda3/envs/tmp/bin/python`
  2. `~/miniconda3/envs/tmp/bin/python`
  3. 시스템 `python3`/`python`

## macOS 앱 번들 빌드

```bash
./build_macos_app.sh
```

빌드 결과: `dist/SpendTagger.app`

## GitHub 업로드 시 개인정보 보호

`.gitignore`에 아래 항목들이 포함되어 있어 기본적으로 업로드에서 제외됩니다.

- 개인 거래 원본/결과 엑셀(`*.xls`, `*.xlsx`)
- 차트/결과 출력물(`charts/`, `app_output/`)
- 로컬 학습 메모리(`tagging_memory.csv`)
- 로컬 컬럼 매핑 프로필(`column_mapping_profiles.json`)
- 빌드 산출물(`dist/`, `build/`, `*.app/`)

## 상세 문서

- `README_mac_app.md`: 앱 기능과 사용 흐름 상세 설명
