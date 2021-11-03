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
        self.valid_date = False
        self.valid_headers = True
        self.from_address = clean_header(from_address)
        self.to_address = clean_header(to_address)
        self.subject = clean_header(subject)
        self.date = self._validate_date(date)

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

        a_date = None
        for d_format in date_formats:
            try:
                a_date = arrow.get(date, d_format)
                break
            except:
                continue

        if not a_date:
            print(
                f"ALERT: '{date}' does not match any expected format. Excluding email with subject '{self.subject}'."
            )
            return date
        else:
            self.valid_date = True
            return a_date

    def clean_header(self, header):
        try:
            return (
                str(make_header(decode_header(re.sub(r"\s\s+", " ", header))))
                if header
                else ""
            )
        except:
            # if verbose:
            #     print(f"Failed to properly decode/reencode header:")
            #     print(header)
            self.valid_headers = False
            return header

    def __str__(self):
        out_str = [
            f"From:\t\t{self.from_address}",
            f"To:\t\t{self.to_address}",
            f"Subject:\t{self.subject}",
            f"Date:\t\t{self.date}",
            f"Valid Date:\t{self.valid_date}",
            f"Valid Headers:\t{self.valid_headers}",
        ]
        return "\n".join(out_str)


def clean_header(header, verbose=False):
    try:
        return (
            str(make_header(decode_header(re.sub(r"\s\s+", " ", header))))
            if header
            else ""
        )
    except:
        if verbose:
            print(f"Failed to properly decode/reencode header:")
            print(header)
        return header


def process_mbox(mbox_filename, year=None, verbose=False):
    count = 0
    ignored = 0
    emails = []

    if verbose and year:
        print(f"Excluding emails not from year {year}.")

    for message in mailbox.mbox(mbox_filename):
        count += 1
        emails.append(
            Email(message["Subject"], message["From"], message["To"], message["Date"])
        )

        # for email in mailbox:
        #   instantiate Email
        #           (clean headers, validate date)
        #   validate date

        # else:
        #     if year and (a_date.format("YYYY") != year):
        #         ignored += 1
        #         if verbose:
        #             print(f"INFO: Invalid year found ({a_date.format('YYYY')}).")

        #     else:
        #         data = [
        #             clean_header(message["Subject"], verbose),
        #             clean_header(message["From"], verbose),
        #             clean_header(message["To"], verbose),
        #             a_date,
        #         ]

        #         emails.append(data)

        if verbose and (count % 1000 == 0):
            print(f"INFO: {count} emails processed.")

    # # TODO fix sorting  issue
    # # Sort based on arrow object
    # emails = sorted(emails, key=lambda x: x[-1])

    # # Convert list to desired string format
    # date_format = "M/D/YY"
    # emails = [[*email[:-1], email[-1].format(date_format)] for email in emails]

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

    invalid_dates = [email for email in emails if email.valid_date]
    invalid_headers = [email for email in emails if email.valid_headers]

    emails = [email for email in emails if email.valid_headers and email.valid_date]

    # TODO Delete this
    for email in emails:
        print(email)
        print()

    # # TODO Fix Exporting
    # # Export data
    # print(f"Beginning export of {len(emails)} emails to {output_filename}...")
    # export_emails(emails, output_filename)

    print(
        f"{len(emails)} emails were found and {'0000000'} were exported to {output_filename}."
    )
    print(f"Completed in {round((timeit.default_timer()-start_time), 2)} seconds.")


if __name__ == "__main__":
    main()
