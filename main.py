#!/usr/env/python3

import argparse
from report import factory


def main():
    parser = argparse.ArgumentParser(description='Generate report with data brazilian data')
    parser.add_argument('--report', nargs=1, type=str,
                        help='the report to be generated\'s name')
    parser.add_argument('--mindate', nargs=1, type=int,
                        help='yyyymmdd - the minimum date to be considered in the report')
    parser.add_argument('--maxdate', nargs=1, type=int,
                        help='yyyymmdd - the maximum date to be considered in the report')
    args = parser.parse_args()
    try:
        report = factory.create_report(args.report[0], args.mindate[0], args.maxdate[0])
        report.generate_report()
    except Exception as ex:
        print(ex)

if __name__ == '__main__':
    main()
