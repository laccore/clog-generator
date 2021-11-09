import csv
import datetime
import mailbox
import os
import re
import timeit
from email.header import decode_header, make_header

import arrow
import requests
from dotenv import load_dotenv
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


def load_filters():
    print("Loading updated filters from Quickbase.")
    filters = {}
    load_dotenv()

    headers = {
        "QB-Realm-Hostname": os.environ.get("REALM"),
        "User-Agent": os.environ.get("UA"),
        "Authorization": os.environ.get("TOKEN"),
    }

    qb_ids = [
        [
            os.environ.get("DOMAINS_TABLE_ID"),
            os.environ.get("DOMAINS_COLUMN_ID"),
            "domains",
        ],
        [
            os.environ.get("EMAILS_TABLE_ID"),
            os.environ.get("EMAILS_COLUMN_ID"),
            "emails",
        ],
        [
            os.environ.get("KEYWORDS_TABLE_ID"),
            os.environ.get("KEYWORDS_COLUMN_ID"),
            "keywords",
        ],
    ]

    for table_id, column_id, name in qb_ids:
        body = {
            "from": table_id,
            "select": [column_id],
            "where": f"{{{column_id}.XEX.''}}",
            "options": {"skip": 0, "top": 0, "compareWithAppLocalTime": False},
        }

        # Make API call
        r = requests.post(
            f"https://api.quickbase.com/v1/records/query", headers=headers, json=body
        )

        # Check if request successfully returned
        if not r.ok:
            return None

        # Access the data we want, add to filters dict
        filters[name] = [record[column_id]["value"] for record in r.json()["data"]]

    return filters


def process_mbox(mbox_filename):
    count = 0
    emails = []

    for message in mailbox.mbox(mbox_filename):
        count += 1
        emails.append(
            Email(message["Subject"], message["From"], message["To"], message["Date"])
        )

        if count % 1000 == 0:
            print(f"  {count} emails processed.")

    return emails, count


def validate_and_sort_emails(emails, year=None):
    if year:
        print(f"Excluding emails not from year {year}.")

    # Check if emails had valid dates and headers
    valid_emails = []
    bad_formats = []

    for email in emails:
        if not email.valid_date or not email.valid_headers:
            bad_formats.append(email)
        elif email.year != year:
            pass
        else:
            valid_emails.append(email)

    # Sort valid emails based on datetime
    valid_emails = sorted(valid_emails, key=lambda x: x.date)

    return valid_emails, bad_formats


def export_emails(emails, output_filename, exclude_subject=False):
    if exclude_subject:
        print("Excluding email Subject field from export.")

    headers = ["Subject", "From Name", "From Email", "To", "Date"]
    with open(output_filename, "w", newline="", encoding="utf-8") as out_file:
        writer = csv.writer(out_file, quoting=csv.QUOTE_MINIMAL)
        if exclude_subject:
            headers = headers[1:]
            emails = [list(email)[1:] for email in emails]
        writer.writerow(headers)
        writer.writerows(emails)

    return None


def export_bad_emails(bad_formats, output_filename):
    output_filename = output_filename.replace(".csv", "_bad_emails.csv")
    print("\nInvalid dates or headers found.")
    print(f"Please email file '{output_filename}' to the project maintainer to fix.\n")
    with open(output_filename, "w", newline="", encoding="utf-8") as out_file:
        writer = csv.writer(out_file, quoting=csv.QUOTE_MINIMAL)

        for bad_email in bad_formats:
            if not bad_email.valid_date:
                writer.writerow(["Incorrect Date Format", bad_email.date])
            if not bad_email.valid_headers:
                writer.writerow(["Incorrect Header Format", bad_email.header])

    return None


@Gooey(program_name="CSD Contact Log (CLOG) Generator")
def main():
    parser = GooeyParser(
        description="Export data from a .mbox file to a csv for use in the CLOG",
    )
    parser.add_argument(
        "mbox",
        metavar="mbox File",
        widget="FileChooser",
        type=str,
        help="Path to your mbox file",
    )
    parser.add_argument(
        "year",
        metavar="Year",
        help="Exclude emails not from entered year",
        default=datetime.datetime.now().year - 1,
        type=int,
    )
    parser.add_argument(
        "-f",
        "--filter",
        metavar="Filter Emails",
        action="store_true",
        help="Filter out known bad domains, email addresses, and subjects from export.",
        default=True,
    )
    parser.add_argument(
        "-ns",
        "--nosubject",
        metavar="Exclude Subject",
        action="store_true",
        help="Exclude the Subject field from export file\n\nThis will reduce the amount of personal information, but make identifying unknown senders more difficult.",
        default=False,
    )

    start_time = timeit.default_timer()
    args = parser.parse_args()
    mailbox_filename = args.mbox
    output_filename = mailbox_filename.replace(".mbox", ".csv")
    year = int(args.year)
    exclude_subject = args.nosubject
    run_filters = args.filter

    if run_filters:
        # Load filter lists
        filters = load_filters()
        if not filters:
            print(
                "ERROR: Could not retrieve filters from Quickbase. Export will not be filtered."
            )
            run_filters = False

    # Process mailbox
    print(f"Beginning processing of {mailbox_filename}...")
    emails, num_emails = process_mbox(mailbox_filename)
    print(f"Processed mailbox {mailbox_filename}.")

    # Validate and sort emails
    (
        valid_emails,
        bad_formats,
    ) = validate_and_sort_emails(emails, year)

    # Export data
    print(f"Beginning export of emails to {output_filename}...")
    export_emails(valid_emails, output_filename, exclude_subject)
    if bad_formats:
        export_bad_emails(bad_formats, output_filename)

    # Done
    print(
        f"\n{num_emails} emails were found and {len(valid_emails)} were exported to {output_filename}.\n"
    )
    print(f"Completed in {round((timeit.default_timer()-start_time), 2)} seconds.")


if __name__ == "__main__":
    main()
