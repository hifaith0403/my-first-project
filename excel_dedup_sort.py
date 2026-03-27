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
