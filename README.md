# Excel 중복 제거 + 날짜 정렬 (Python)

코딩 초보자도 **복붙해서 바로 실행**할 수 있도록 만든 예제입니다.

## 1) 준비물 (Windows)

1. Python 3.10 이상 설치
2. 명령 프롬프트(cmd) 또는 PowerShell 사용

## 2) 설치 방법

프로젝트 폴더에서 아래 명령어를 순서대로 실행하세요.

```bash
python -m venv .venv
.venv\Scripts\activate
pip install pandas openpyxl
```

> `openpyxl`은 `.xlsx` 파일 읽기/쓰기용 라이브러리입니다.

## 3) 전체 코드

파일명: `excel_dedup_sort.py`

```python
import argparse
from pathlib import Path
import sys

import pandas as pd


def process_excel(
    input_path: Path,
    output_path: Path,
    sheet_name: str,
    date_column: str,
    keep: str,
    ascending: bool,
) -> None:
    # 1) 엑셀 읽기
    df = pd.read_excel(input_path, sheet_name=sheet_name)

    # 2) 날짜 컬럼을 datetime으로 변환 (변환 실패 시 NaT)
    df[date_column] = pd.to_datetime(df[date_column], errors="coerce")

    # 3) 전체 행 기준 중복 제거
    #    keep='first' -> 첫 번째만 남김, keep='last' -> 마지막만 남김
    df = df.drop_duplicates(keep=keep)

    # 4) 날짜순 정렬
    df = df.sort_values(by=date_column, ascending=ascending)

    # 5) 저장
    # index=False: 엑셀에 인덱스 번호 열을 추가하지 않음
    df.to_excel(output_path, index=False)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="엑셀 파일에서 중복 행을 제거하고 날짜순으로 정렬합니다."
    )
    parser.add_argument("--input", required=True, help="입력 엑셀 파일 경로 (예: input.xlsx)")
    parser.add_argument(
        "--output",
        default="output_cleaned.xlsx",
        help="출력 엑셀 파일 경로 (기본값: output_cleaned.xlsx)",
    )
    parser.add_argument(
        "--sheet",
        default=0,
        help="읽을 시트 이름 또는 인덱스 (기본값: 첫 번째 시트)",
    )
    parser.add_argument(
        "--date-column",
        required=True,
        help="정렬 기준이 되는 날짜 컬럼명 (예: 주문일)",
    )
    parser.add_argument(
        "--keep",
        choices=["first", "last"],
        default="first",
        help="중복 시 어떤 행을 남길지 선택 (기본값: first)",
    )
    parser.add_argument(
        "--descending",
        action="store_true",
        help="이 옵션을 넣으면 최신 날짜 -> 오래된 날짜 순으로 정렬",
    )

    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        print(f"[오류] 입력 파일을 찾을 수 없습니다: {input_path}")
        sys.exit(1)

    # sheet가 숫자 형태라면 int로 변환, 아니면 문자열 시트명으로 사용
    sheet_name = int(args.sheet) if str(args.sheet).isdigit() else args.sheet

    try:
        process_excel(
            input_path=input_path,
            output_path=output_path,
            sheet_name=sheet_name,
            date_column=args.date_column,
            keep=args.keep,
            ascending=not args.descending,
        )
    except KeyError:
        print(f"[오류] 날짜 컬럼명을 찾을 수 없습니다: {args.date_column}")
        print("      엑셀 헤더(컬럼명) 오타 여부를 확인하세요.")
        sys.exit(1)
    except Exception as e:
        print(f"[오류] 처리 중 문제가 발생했습니다: {e}")
        sys.exit(1)

    print("완료! 결과 파일이 생성되었습니다:")
    print(output_path.resolve())


if __name__ == "__main__":
    main()
```

## 4) 실행 방법

예시 1) 기본 오름차순 정렬(오래된 날짜 -> 최신 날짜)

```bash
python excel_dedup_sort.py --input input.xlsx --date-column 주문일
```

예시 2) 내림차순 정렬(최신 날짜 -> 오래된 날짜)

```bash
python excel_dedup_sort.py --input input.xlsx --date-column 주문일 --descending
```

예시 3) 출력 파일명 지정

```bash
python excel_dedup_sort.py --input input.xlsx --date-column 주문일 --output result.xlsx
```

## 5) 결과 예시

입력(`input.xlsx`) 예시:

| 주문번호 | 고객명 | 주문일 |
|---|---|---|
| A001 | Kim | 2026-01-10 |
| A001 | Kim | 2026-01-10 |
| A002 | Lee | 2026-01-08 |
| A003 | Park | 2026-01-12 |

실행 후 출력(`output_cleaned.xlsx`) 예시:

| 주문번호 | 고객명 | 주문일 |
|---|---|---|
| A002 | Lee | 2026-01-08 |
| A001 | Kim | 2026-01-10 |
| A003 | Park | 2026-01-12 |

- 중복된 `A001` 행은 1개만 남음
- `주문일` 기준으로 날짜순 정렬됨

## 6) 실패할 때 확인할 점 3가지

1. **컬럼명 오타 확인**
   - `--date-column 주문일`의 컬럼명이 엑셀 헤더와 완전히 같아야 합니다.
2. **엑셀 파일 경로 확인**
   - `--input input.xlsx` 경로가 현재 폴더 기준인지 확인하세요.
   - 필요하면 절대경로를 사용하세요. 예: `C:\data\input.xlsx`
3. **패키지 설치 확인**
   - `ModuleNotFoundError`가 나면 가상환경 활성화 후 아래를 다시 실행하세요.
   - `pip install pandas openpyxl`
