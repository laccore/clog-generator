import csv
import datetime
import mailbox
import re
import timeit
from email.header import decode_header, make_header

import arrow
from flanker.addresslib import address
from gooey import Gooey, GooeyParser

from filters import load_filters


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
        self.passed_filters = None
        self.filter_reason = None
        self.filter_value = None

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
            r"M/D/YY,[\s+]H:m",
            r"DD[\s+]MMM[\s+]YYYY[\s+]HH:mm:ss",
            r"MM/DD/YY,[\s+]mm[\s+]HH[\s+]YYYY[.*]",
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

    def filtered_iterable(self):
        return iter(
            [
                self.subject,
                self.from_address_name,
                self.from_address_email,
                self.to_address,
                self.us_date,
                self.filter_reason,
                self.filter_value,
            ]
        )

    def __str__(self):
        out_str = [
            f"From:\t\t{self.from_address}",
            f"To:\t\t{self.to_address}",
            f"Subject:\t{self.subject}",
            f"Date:\t\t{self.date} ({self.year})",
            f"US Date:\t{self.us_date}",
            f"Valid Date:\t{self.valid_date}",
            f"Valid Headers:\t{self.valid_headers}",
            f"Passed Filters:\t{self.passed_filters}",
        ]
        if not self.passed_filters:
            out_str += [
                f"Filter Reason:\t{self.filter_reason}",
                f"Filter Value:\t{self.filter_value}",
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


def check_against_filters(email, filters):
    domains = filters["domains"]
    emails = filters["emails"]
    keywords = filters["keywords"]
    staff = filters["staff"]

    if email.from_address_host in domains:
        email.passed_filters = False
        email.filter_reason = "Domain in filter list"
        email.filter_value = email.from_address_host
    elif email.from_address_email in staff:
        email.passed_filters = False
        email.filter_reason = "Staff"
        email.filter_value = email.from_address_email
    elif email.from_address_email in emails:
        email.passed_filters = False
        email.filter_reason = "Email address in filter list"
        email.filter_value = email.from_address_email
    elif any([keyword in email.subject for keyword in keywords]):
        email.passed_filters = False
        email.filter_reason = "Subject contains keyword in filter list"
        email.filter_value = email.subject
    else:
        email.passed_filters = True

    return email


def validate_and_sort_emails(emails, year=None, filters=False):
    if year:
        print(f"Excluding emails not from year {year}.", "\n")

    if filters:
        emails = [check_against_filters(email, filters) for email in emails]

    # Check if emails had valid dates and headers
    valid_emails = []
    bad_formats = []
    filtered_emails = []

    for email in emails:
        if not email.valid_date or not email.valid_headers:
            bad_formats.append(email)
        elif email.year != year:
            email.passed_filters = False
            email.filter_reason = "Incorrect Year"
            email.filter_value = email.year
            filtered_emails.append(email)
        elif filters and not email.passed_filters:
            filtered_emails.append(email)
        else:
            valid_emails.append(email)

    if filters:
        # Remove staff emails from filtered emails
        filtered_emails = [
            email
            for email in filtered_emails
            if email.filter_reason not in ("Staff", "Incorrect Year")
        ]
        filtered_emails = sorted(filtered_emails, key=lambda x: x.date)

    # Sort valid emails based on datetime
    valid_emails = sorted(valid_emails, key=lambda x: x.date)

    return valid_emails, bad_formats, filtered_emails


def export_emails(emails, output_filename, exclude_subject=False):
    if exclude_subject:
        print("Excluding email Subject field from export.")

    headers = ["Subject", "Gmail Name", "From Email", "To", "Date"]
    with open(output_filename, "w", newline="", encoding="utf-8") as out_file:
        writer = csv.writer(out_file, quoting=csv.QUOTE_MINIMAL)
        if exclude_subject:
            headers = headers[1:]
            emails = [list(email)[1:] for email in emails]
        writer.writerow(headers)
        writer.writerows(emails)

    return None


def process_filter_stats(emails, print_output=True):
    email_filter_reasons = [email.filter_reason for email in emails]
    filter_reasons = set(email_filter_reasons)
    filter_counts = {
        filter_reason: email_filter_reasons.count(filter_reason)
        for filter_reason in filter_reasons
    }
    # Sort dict in descending order
    filter_counts = {
        k: v for k, v in sorted(filter_counts.items(), key=lambda x: x[1], reverse=True)
    }

    print(f"{type(filter_counts)=}")

    if print_output:
        print()
        print(f"Number of emails filtered: {len(emails)}")
        print("Number filtered by reason:")
        for f in filter_counts:
            print(f"\t{f}: {filter_counts[f]}")


def export_filtered_emails(filtered_emails, output_filename):
    headers = [
        "Subject",
        "Gmail Name",
        "From Email",
        "To",
        "Date",
        "Filter Reason",
        "Filter Value",
    ]
    with open(output_filename, "w", newline="", encoding="utf-8") as out_file:
        writer = csv.writer(out_file, quoting=csv.QUOTE_MINIMAL)
        writer.writerow(headers)
        filtered_emails = [email.filtered_iterable() for email in filtered_emails]
        writer.writerows(filtered_emails)

    return output_filename


def export_bad_emails(bad_formats, output_filename):
    with open(output_filename, "w", newline="", encoding="utf-8") as out_file:
        writer = csv.writer(out_file, quoting=csv.QUOTE_MINIMAL)

        for bad_email in bad_formats:
            if not bad_email.valid_date:
                writer.writerow(["Incorrect Date Format", bad_email.date])
            if not bad_email.valid_headers:
                writer.writerow(["Incorrect Header Format", bad_email.header])

    return output_filename


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
        print("Loading updated filters from Quickbase.")
        filter_start_time = timeit.default_timer()
        filters = load_filters()
        if not filters:
            print(
                "ERROR: Could not retrieve filters from Quickbase. Export will not be filtered."
            )
            run_filters = False
        print(
            f"Completed loading filters ({round((timeit.default_timer()-filter_start_time), 2)}s).",
            "\n",
        )

    # Process mailbox
    print(f"Beginning processing of {mailbox_filename}...")
    emails, num_emails = process_mbox(mailbox_filename)
    print(f"Completed mailbox processing.", "\n")

    # Validate and sort emails
    (valid_emails, bad_formats, filtered_emails) = validate_and_sort_emails(
        emails, year, filters
    )

    # Export data
    export_emails(valid_emails, output_filename, exclude_subject)
    print(f"Exported valid emails to '{output_filename}'.")

    if run_filters:
        filtered_output_filename = output_filename.replace(
            ".csv", "_filtered_emails.csv"
        )
        print(f"Exported filtered emails to '{filtered_output_filename}'.")
        export_filtered_emails(filtered_emails, filtered_output_filename)
        process_filter_stats(filtered_emails)

    if bad_formats:
        bad_formats_output_filename = output_filename.replace(".csv", "_bad_emails.csv")
        print("\n", "WARNING: Invalid dates or headers found.")
        print(
            f"Please email '{bad_formats_output_filename}' to the project maintainer to fix.",
            "\n",
        )
        export_bad_emails(bad_formats, bad_formats_output_filename)

    # Done
    print()
    print(
        f"{num_emails} emails were found and {len(valid_emails)} were exported to '{output_filename}' ({round((timeit.default_timer()-start_time), 2)}s).\n",
    )


if __name__ == "__main__":
    main()
