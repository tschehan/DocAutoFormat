"""Simple local test entry."""

from pathlib import Path

from pipeline import run_pipeline


def main() -> None:
    """Run a simple local pipeline test."""
    input_file = "input.docx"
    output_file = f"{Path(input_file).stem}-formatted.docx"
    requirement_text = (
        "正文用宋体小四，1.5倍行距，两端对齐；"
        "论文标题用黑体三号加粗居中；页码不用处理。"
    )

    print("DocAutoFormat pipeline test")

    if not Path(input_file).exists():
        print("Please put input.docx in the current folder.")
        return

    result = run_pipeline(
        input_file=input_file,
        requirement_text=requirement_text,
        output_file=output_file,
        progress_callback=print,
    )

    print("\nPipeline result:")
    print(result)


if __name__ == "__main__":
    main()
