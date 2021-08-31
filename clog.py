import csv
import datetime
import mailbox
import re
import timeit
from email.header import decode_header, make_header

import arrow
import chardet
from flanker.addresslib import address
from gooey import Gooey, GooeyParser


def guess_file_encoding(input_file, verbose=False):
    with open(
        input_file,
        "rb",
    ) as f:
        rawdata = f.read(1024)
        enc = chardet.detect(rawdata)
        if verbose:
            print(f'Guess on encoding of {input_file}: {enc["encoding"]}')
    return enc["encoding"]


def import_ignore_lists(file_list, verbose=False):
    filters = {}
    for ignore_list in file_list:
        ignore_type = ignore_list.split("_")[-1].replace(".csv", "")
        enc = guess_file_encoding(ignore_list)
        with open(ignore_list, "r+", encoding=enc) as f:
            items = set(f.read().splitlines())
            filters[ignore_type] = items
            if verbose:
                print(f"Imported ignore list: {ignore_type}")
    return filters


def check_date_format(message):
    # When downloading from Google Takeout, there are a few different datetime formats
    # So far, they've all matched one of the below options (with lots of regex for extra whitespace...)
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

    date_string = message["Date"]

    a_date = None
    for d_format in date_formats:
        try:
            a_date = arrow.get(date_string, d_format)
            break
        except:
            continue
    return a_date


def check_year(date, year, verbose=False):
    if date.format("YYYY") == year:
        return True
    else:
        return False


def check_filters(email, filters):
    subject = email[0]
    from_address = address.parse(email[1])

    if from_address.address in filters["emails"]:
        fail_reason = ["From Address", from_address.address]
        return (False, fail_reason)
    elif from_address.hostname in filters["domains"]:
        fail_reason = ["Domain", from_address.hostname]
        return (False, fail_reason)
    else:
        for keyword in filters["subjects"]:
            if keyword in subject:
                fail_reason = ["Keyword", keyword]
                return (False, fail_reason)
    return (True, "")


def clean_header(header, verbose=False):
    try:
        return (
            str(make_header(decode_header(re.sub(r"\s\s+", " ", header))))
            if header
            else ""
        )
    except:
        if verbose:
            print(f"Failed to properly decode/reencode header. Send text below to Alex.")
            print("Header:")
            print(header)
        return header


def process_mbox(mbox_filename, filters={}, year=None, verbose=False):
    count = 0
    emails = []
    filtered_emails = []

    if verbose and year:
        print(f"Excluding emails not from year {year}.")

    for message in mailbox.mbox(mbox_filename):
        count += 1

        a_date = check_date_format(message)

        email = [
            clean_header(message["Subject"], verbose),
            clean_header(message["From"], verbose),
            clean_header(message["To"], verbose),
            a_date,
        ]

        if not a_date:
            # Invalid date format
            filter_reason = ["Invalid date format. Send this line to Alex to fix.", ""]
            filtered_emails.append(email + filter_reason)
        elif not check_year(a_date, year):
            # Valid date format, invalid year
            filter_reason = ["Year", a_date.format("YYYY")]
            filtered_emails.append(email + filter_reason)
        else:
            # Valid date format, valid year
            # Check against filters, if using filters
            if filters:
                passed_filters, filter_reason = check_filters(email, filters)
                if passed_filters:
                    emails.append(email)
                else:
                    filtered_emails.append(email + filter_reason)
            else:
                emails.append(email)

        if verbose and (count % 1000 == 0):
            print(f"INFO: {count} emails processed.")

    # Sort based on arrow object
    emails = sorted(emails, key=lambda x: x[-1])
    filtered_emails = sorted(filtered_emails, key=lambda x: x[-3])

    # Convert list to desired string format
    date_format = "M/D/YY"
    emails = [[*email[:-1], email[-1].format(date_format)] for email in emails]

    emails.insert(0, ["Subject", "From", "To", "Date"])
    filtered_emails.insert(
        0, ["Subject", "From", "To", "DateTime", "Filter Reason", "Filter Value"]
    )

    return [emails, count, filtered_emails]


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
        metavar=".mbox file",
        widget="FileChooser",
        type=str,
        help="Name of mbox file",
    )
    parser.add_argument(
        "year",
        metavar="Year",
        help="Exclude emails not from this year",
        default=datetime.datetime.now().year - 1,
    ),
    parser.add_argument(
        "-f",
        "--filter",
        metavar="Filter",
        action="store_true",
        help="Filter emails against known ignore lists",
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
    filtered_filename = mailbox_filename.replace(".mbox", "_filtered.csv")

    # test ignore lists
    filters_file_list = [
        "filter_lists/ignore_emails.csv",
        "filter_lists/ignore_domains.csv",
        "filter_lists/ignore_subjects.csv",
    ]
    filters = import_ignore_lists(filters_file_list, args.verbose)

    # Process data
    print(f"Beginning processing of {mailbox_filename}...")
    emails, message_count, filtered_emails = process_mbox(
        mailbox_filename, filters, args.year, args.verbose
    )

    # Export data
    print(f"Beginning export of {len(emails)-1} emails to {output_filename}...")
    export_emails(emails, output_filename)
    export_emails(filtered_emails, filtered_filename)

    print(
        f"{message_count} emails were found and {len(emails)-1} were exported to {output_filename}."
    )
    print(f"Completed in {round((timeit.default_timer()-start_time), 2)} seconds.")


if __name__ == "__main__":
    main()
