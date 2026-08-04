import sys
import logging

from report_engine import get_latest_excel, generate_report, ValidationError


def main():
    tickets_file = get_latest_excel()

    try:
        result = generate_report(tickets_file)
    except ValidationError as e:
        for msg in e.errors:
            logging.error(msg)
        logging.error("PLEASE CORRECT ERRORS IN THE EXCEL FILE TO CONTINUE EXECUTION...")
        sys.exit(1)

    logging.info(f"Report successfully generated at: {result['output_path']}")


if __name__ == "__main__":
    main()
