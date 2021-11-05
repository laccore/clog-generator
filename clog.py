import csv
import datetime
import mailbox
import os.path
import re
import timeit
from email.header import decode_header, make_header

import arrow
from flanker.addresslib import address
from gooey import Gooey, GooeyParser


class Email:
    def __init__(self, subject, from_address, to_address, date):
        self.valid_date = True
        self.valid_headers = True
        self.from_address = address.parse(self._clean_header(from_address))
        self.from_address_email = self.from_address.address
        self.from_address_name = self.from_address.display_name
        self.from_address_host = self.from_address.hostname
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
        return iter(
            [
                self.subject,
                self.from_address_name,
                self.from_address_email,
                self.to_address,
                self.us_date,
            ]
        )


def process_mbox(mbox_filename, verbose=False):
    count = 0
    emails = []

    for message in mailbox.mbox(mbox_filename):
        count += 1
        emails.append(
            Email(message["Subject"], message["From"], message["To"], message["Date"])
        )

        if verbose and (count % 1000 == 0):
            print(f"INFO: {count} emails processed.")

    return emails, count


def validate_and_sort_emails(emails, year=None, verbose=False):
    if verbose and year:
        print(f"Excluding emails not from year {year}.")

    # Sort emails if they had valid dates and headers
    bad_date_formats = []
    bad_header_formats = []
    validated_emails = []

    for email in emails:
        if not email.valid_date:
            bad_date_formats.append(email)
        elif not email.valid_headers:
            bad_header_formats.append(email)
        elif email.year != year:
            pass
        else:
            validated_emails.append(email)

    # Sort valid emails based on datetime
    validated_emails = sorted(validated_emails, key=lambda x: x.date)

    return validated_emails, bad_date_formats, bad_header_formats


def export_emails(emails, output_filename, exclude_subject=False):
    headers = ["Subject", "From Name", "From Email", "To", "Date"]
    with open(output_filename, "w", newline="", encoding="utf-8") as out_file:
        writer = csv.writer(out_file, quoting=csv.QUOTE_MINIMAL)
        if exclude_subject:
            headers.pop(0)
            emails = [list(email)[1:] for email in emails]
        writer.writerow(headers)
        writer.writerows(emails)


def export_bad_emails(bad_date_formats, bad_header_formats, output_filename):
    output_filename = output_filename.replace(".csv", "_bad_emails.csv")
    print("\nbad dates or headers found.")
    print(f"Please email file '{output_filename}' to the project maintainer to fix.\n")
    with open(output_filename, "w", newline="", encoding="utf-8") as out_file:
        writer = csv.writer(out_file, quoting=csv.QUOTE_MINIMAL)

        if bad_date_formats:
            for bad_email in bad_date_formats:
                writer.writerow(["Incorrect Date Format", bad_email.date])

        if bad_header_formats:
            for bad_email in bad_header_formats:
                writer.writerow(["Incorrect Header Format", bad_email.header])


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
        type=int,
    )
    parser.add_argument(
        "-ns",
        "--nosubject",
        metavar="No Subject",
        action="store_true",
        help="Exclude email Subject from exports",
        default=False,
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

    # Process mailbox
    print(f"Beginning processing of {mailbox_filename}...")
    emails, num_emails = process_mbox(mailbox_filename, args.verbose)

    # Validate and sort emails
    (
        validated_emails,
        bad_date_formats,
        bad_header_formats,
    ) = validate_and_sort_emails(emails, int(args.year))

    # Export data
    print(f"Beginning export of {len(validated_emails)} emails to {output_filename}...")
    export_emails(validated_emails, output_filename, args.nosubject)

    # Export bad dates/headers
    if bad_date_formats or bad_header_formats:
        export_bad_emails(bad_date_formats, bad_header_formats, output_filename)

    print(
        f"{num_emails} emails were found and {len(validated_emails)} were exported to {output_filename}."
    )
    print(f"Completed in {round((timeit.default_timer()-start_time), 2)} seconds.")


if __name__ == "__main__":
    main()
