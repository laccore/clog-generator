import csv
import datetime
import mailbox
import os.path
import re
import timeit
from email.header import decode_header, make_header

import arrow
from gooey import Gooey, GooeyParser


class Email:
    def __init__(self, subject, from_address, to_address, date):
        self.valid_date = True
        self.valid_headers = True
        self.from_address = self._clean_header(from_address)
        self.to_address = self._clean_header(to_address)
        self.subject = self._clean_header(subject)
        self.date = self._validate_date(date)
        self.year = self._find_year(self.date)
        self.us_date = self.date.format("M/D/YY")

    def _validate_date(self, date):
        date_formats = [
            r"ddd,[\s+]D[\s+]MMM[\s+]YYYY[\s+]H:mm:ss[\s+]Z",
            r"ddd,[\s+]D[\s+]MMM[\s+]YYYY[\s+]H:mm:ss[\s+]ZZZ",
            r"ddd,[\s+]D[\s+]MMM[\s+]YYYY[\s+]H:mm:ss[\s+]",
            r"ddd,[\s+]DD[\s+]MMM[\s+]YYYY[\s+]HH:mm:ss",
            r"ddd[\s+]D[\s+]MMM[\s+]YYYY[\s+]H:mm:ss[\s+]Z",
            r"D[\s+]MMM[\s+]YYYY[\s+]HH:mm:ss[\s+]Z",
            r"ddd,[\s+]D[\s+]MMM[\s+]YYYY[\s+]H:mm[\s+]Z",
            r"MM/D/YY,[\s+]H[\s+]mm[.*]",
        ]

        arrow_date = None
        for date_format in date_formats:
            try:
                arrow_date = arrow.get(date, date_format)
                break
            except:
                continue
        if not arrow_date:
            self.valid_date = False
            return date
        else:
            return arrow_date

    def _find_year(self, date):
        if isinstance(date, arrow.Arrow):
            return date.year
        else:
            return None

    def _clean_header(self, header):
        try:
            return (
                str(make_header(decode_header(re.sub(r"\s\s+", " ", header))))
                if header
                else ""
            )
        except:
            self.valid_headers = False
            return header

    def __str__(self):
        out_str = [
            f"From:\t\t{self.from_address}",
            f"To:\t\t{self.to_address}",
            f"Subject:\t{self.subject}",
            f"Date:\t\t{self.date} ({self.year})",
            f"US Date:\t{self.us_date}",
            f"Valid Date:\t{self.valid_date}",
            f"Valid Headers:\t{self.valid_headers}",
        ]
        return "\n".join(out_str)

    def __iter__(self):
        return iter([self.subject, self.from_address, self.to_address, self.us_date])


def process_mbox(mbox_filename, year=None, verbose=False):
    count = 0
    emails = []

    if verbose and year:
        print(f"Excluding emails not from year {year}.")

    for message in mailbox.mbox(mbox_filename):
        count += 1
        emails.append(
            Email(message["Subject"], message["From"], message["To"], message["Date"])
        )

        if verbose and (count % 1000 == 0):
            print(f"INFO: {count} emails processed.")

    return emails


def export_emails(emails, output_filename):
    with open(output_filename, "w", newline="", encoding="utf-8") as out_file:
        writer = csv.writer(out_file, quoting=csv.QUOTE_MINIMAL)
        writer.writerows(emails)


@Gooey(program_name="CSDCO CLOG Generator")
def main():
    parser = GooeyParser(
        description="Export data (Subject, From, To, Date) from a .mbox file to a CSV"
    )
    parser.add_argument(
        "mbox",
        metavar=".mbox-file",
        widget="FileChooser",
        type=str,
        help="Name of mbox file",
    )
    parser.add_argument(
        "year",
        metavar="Year",
        help="Exclude emails not from this year",
        default=datetime.datetime.now().year - 1,
    )
    parser.add_argument(
        "-v",
        "--verbose",
        metavar="Verbose",
        action="store_true",
        help="Print troubleshooting information",
    )
    args = parser.parse_args()

    start_time = timeit.default_timer()

    mailbox_filename = args.mbox
    output_filename = mailbox_filename.replace(".mbox", ".csv")

    # Process data
    print(f"Beginning processing of {mailbox_filename}...")
    emails = process_mbox(mailbox_filename, args.year, args.verbose)

    invalid_dates = []
    invalid_headers = []
    validated_emails = []
    for email in emails:
        if not email.valid_date:
            invalid_dates.append(email)
        elif not email.valid_headers:
            invalid_headers.append(email)
        else:
            validated_emails.append(email)

    # Sort based on datetime
    validated_emails = sorted(validated_emails, key=lambda x: x.date)

    # Export data
    print(f"Beginning export of {len(validated_emails)} emails to {output_filename}...")
    export_emails(validated_emails, output_filename)

    # Check for invalid dates
    if invalid_dates:
        print("\n--------- ALERT ---------")
        print(
            f"Found {len(invalid_dates)} date(s) that do not match any expected format."
        )
        print("Please email the below information to the project maintainer to fix.")

        for invalid_email in invalid_dates:
            print(invalid_email.date)
        print()

    # Check for invalid headers
    if invalid_headers:
        print("\n--------- ALERT ---------")
        print(
            f"Found {len(invalid_headers)} headers that do not match any expected format."
        )
        print("Please email the below information to the project maintainer to fix.")

        for invalid_email in invalid_headers:
            print(invalid_email.header)
        print()

    num_emails_exported = len(emails) - (len(invalid_dates) + len(invalid_headers))
    print(
        f"{len(emails)} emails were found and {num_emails_exported} were exported to {output_filename}."
    )
    print(f"Completed in {round((timeit.default_timer()-start_time), 2)} seconds.")


if __name__ == "__main__":
    main()
