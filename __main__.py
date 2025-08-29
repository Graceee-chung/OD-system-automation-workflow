import argparse
import logging
from .logger_setup import *
from .document_handler import process_documents

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="自動處理待辦公文下載")
    parser.add_argument("--start_idx", type=int, default=0, help="起始索引（從 0 開始）")
    parser.add_argument("--end_idx", type=int, default=0, help="結束索引（不含，0 表示處理到結尾）")
    args = parser.parse_args()

    process_documents(start_idx=args.start_idx, end_idx=args.end_idx)
    process_documents(start_idx=args.start_idx, end_idx=args.end_idx)

    print("🎉 完成全部處理~！")
    logging.info("🎉 完成全部處理~！")
