
from pathlib import Path
import zipfile

from mor_parser import detect_file, unpack_upload, combine_same_named_datasets


def main():
    import sys
    if len(sys.argv) < 2:
        print("Usage: python test_extractor.py <pdf-or-zip>")
        raise SystemExit(2)

    path = Path(sys.argv[1])
    raw = path.read_bytes()
    all_ds = []

    for name, data in unpack_upload(path.name, raw):
        all_ds.extend(detect_file(name, data))

    ds = combine_same_named_datasets(all_ds)

    for item in ds:
        print(item.name, item.dataframe.shape, item.confidence)
        if "DO (mg/L)" in item.dataframe.columns:
            print(item.dataframe[["Date", "DO (mg/L)"]].head(8).to_string(index=False))


if __name__ == "__main__":
    main()
