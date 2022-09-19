import pandas as pd
import random
import xlTools
import argparse

write_file_loc = "/Users/charlie//Documents/personal_tools/Excel_Tools/excel_files/test_cells.xlsx"

def test(args):
    print(args.text)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers()

    write_frame_parser = subparsers.add_parser('writeFrame')
    write_frame_parser.add_argument("dataframe")
    write_frame_parser.add_argument("file_location")
    write_frame_parser.add_argument("offset")
    write_frame_parser.add_argument("--unsafe", action='store_true')
    write_frame_parser.set_defaults(func=xlTools.writeFrame)

    test_parser = subparsers.add_parser('test')
    test_parser.add_argument("text")
    test_parser.set_defaults(func=test)

    args = parser.parse_args()

    result = args.func(args)
    print()
    if result:
        print("Success!")
    else:
        print("Failed to update Excel file.")